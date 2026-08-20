/* Delta — local-only version history for Excel.
 *
 * The engine works like a camera:
 *  1. A snapshot photographs every sheet: each used cell's value + formula.
 *  2. "Review changes" photographs again and compares, cell by cell.
 *  3. "Save version" stores the new photo in a local version history,
 *     together with the user's message ("what did you change?").
 *
 * Every cell change carries a type — "add", "edit", or "del" (emptied) —
 * which drives both the panel rendering and the in-grid highlight colors.
 * Everything is stored locally on the user's computer — no data leaves it.
 */

const MAX_CHANGES_SHOWN = 300;
const MAX_HIGHLIGHTS = 200;
const MAX_VERSIONS = 20;

// In-grid highlight colors, by change type — applied to the cell border and
// font only, never the fill.
const HL_COLORS = {
  add: "#188038",    // green — cell was added
  edit: "#B26A00",   // amber — cell was edited
  del: "#D93025",    // red — cell was emptied / removed
  recalc: "#1A73E8", // blue — formula result changed indirectly
};
const HL_LABELS = { add: "Added", edit: "Edited", del: "Removed", recalc: "Recalculated" };

// Sheet-tab tints are deliberately pastel: on a light fill Excel keeps the tab
// label as crisp dark text, while a saturated fill switches it to fuzzy
// bold-white-on-color.
const TAB_COLORS = {
  add: "#CDEBD8",  // soft green — new sheet
  edit: "#F6E3B4", // soft amber — modified sheet
};

let workbookKey = "esr:untitled";
let lastDiff = null;              // diff currently shown in the panel
let lastCurrentSnapshot = null;   // snapshot the diff was computed from
let lastBaseVersion = null;       // version the diff was computed against
let lastIsPastVersion = false;
let highlightState = null;        // original fills to restore, or null
let activeSheetTab = null;        // sheet id currently shown in the Tabs layout

const SETTINGS_KEY = "esr:settings";
let settings = { diffView: "stacked" };
try {
  settings = { ...settings, ...(JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {}) };
} catch (e) { /* corrupted settings: keep defaults */ }
// Older builds stored "expanded"/"collapsed" — both fold into the stacked view.
if (settings.diffView !== "tabs") settings.diffView = "stacked";

const $ = (id) => document.getElementById(id);

function icon(name, cls) {
  return `<svg class="ico ${cls || ""}"><use href="#i-${name}"/></svg>`;
}

/* ---------- workbook identity ---------- */
// Everything is stored locally, on this computer only — nothing ever leaves it.

let workbookId = null;    // permanent id stored inside the workbook file
let workbookName = "Untitled workbook";

// A UUID stored in the workbook's own settings — survives renames and copies
// (persisted into the file when the user saves it).
function ensureWorkbookId() {
  try {
    const s = Office.context.document.settings;
    let id = s.get("esrWorkbookId");
    if (!id) {
      id = (crypto.randomUUID ? crypto.randomUUID() : "wb-" + Date.now() + "-" + Math.floor(Math.random() * 1e9))
        .replace(/[^a-zA-Z0-9-]/g, "");
      s.set("esrWorkbookId", id);
      s.saveAsync(() => {});
    }
    return id;
  } catch (e) {
    return null;
  }
}

// File line above "Review changes": name + folder, and a "Save file" chip
// while the workbook has never been saved (its embedded history id only
// becomes permanent once the file exists on disk).
function refreshFileLine() {
  const url = Office.context.document.url || "";
  workbookName = url ? url.split("/").pop() : "Untitled workbook";
  $("file-name").textContent = workbookName;
  $("file-dir").textContent = url.includes("/") ? url.slice(0, url.lastIndexOf("/")) : "";
  $("file-line").title = url || workbookName;
  const btn = $("btn-save-file");
  btn.textContent = url ? "Save" : "Save file";
  btn.title = url
    ? "Save the workbook. (Add-ins can't open Save As — use Cmd+Shift+S in Excel to rename or move the file.)"
    : "This workbook has never been saved — its version history can't reconnect to it until it exists as a file.";
}

async function onSaveFile() {
  const wasUnsaved = !Office.context.document.url;
  try {
    await Excel.run(async (context) => {
      // "prompt" asks for name/location only for a never-saved file; for a
      // saved one it saves in place (Office.js has no Save As).
      context.workbook.save(Excel.SaveBehavior ? Excel.SaveBehavior.prompt : "Prompt");
      await context.sync();
    });
    if (wasUnsaved) {
      // The save dialog is asynchronous — watch for the file getting a path.
      let tries = 0;
      const timer = setInterval(() => {
        refreshFileLine();
        if (Office.context.document.url || ++tries > 40) clearInterval(timer);
      }, 1500);
    } else {
      setStatus("Workbook saved.", "ok");
    }
  } catch (e) {
    setStatus("Could not save from here — save with Cmd+S instead. (" + (e.message || e) + ")", "err");
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    $("snapshot-info").textContent = "This add-in only works in Excel.";
    return;
  }
  const url = Office.context.document.url || "";
  workbookName = url ? url.split("/").pop() : "Untitled workbook";
  workbookId = ensureWorkbookId();
  // History is keyed by a per-workbook UUID, so
  // every workbook (saved or not) has its own history. URL is the fallback
  // only if document settings are unavailable.
  workbookKey = workbookId ? "esr:wb:" + workbookId : "esr:" + (url || "untitled");
  // File name + folder shown above the "Review changes" button, with a
  // "Save file" nudge while the workbook has never been saved to disk.
  refreshFileLine();
  $("file-line").hidden = false;
  $("btn-save-file").onclick = onSaveFile;
  migrateOldFormat();
  clearStaleHighlight();
  // Closing the pane (Excel's own X) kills this runtime — kick off the restore
  // the instant unload starts. If it doesn't get to finish, the persisted
  // state means clearStaleHighlight finishes the job on the next open.
  const unloadCleanup = () => {
    if (!highlightState) return;
    const originals = highlightState;
    highlightState = null;
    try {
      restoreFills(originals)
        .then(() => localStorage.removeItem(hlKey()))
        .catch(() => { highlightState = originals; });
    } catch (e) { highlightState = originals; }
  };
  window.addEventListener("pagehide", unloadCleanup);
  window.addEventListener("beforeunload", unloadCleanup);
  // With the shared runtime (manifest <Runtimes lifetime="long">) the add-in
  // keeps running after the pane's X is pressed; this event fires instead of
  // an unload, and the restore runs reliably in the background.
  if (Office.addin && Office.addin.onVisibilityModeChanged) {
    Office.addin.onVisibilityModeChanged((args) => {
      if (args.visibilityMode === "Hidden") onClearHighlights();
    });
  }
  $("btn-snapshot").disabled = false;
  $("btn-review").disabled = false;
  $("btn-snapshot").onclick = onFirstSnapshot;
  $("btn-review").onclick = () => onReview(null);
  $("btn-dismiss").onclick = hideResults;
  $("btn-highlight").onclick = onHighlight;
  $("btn-clear-highlight").onclick = onClearHighlights;

  $("btn-settings").onclick = () => {
    const panel = $("settings-panel");
    panel.hidden = !panel.hidden;
  };
  for (const radio of document.querySelectorAll('input[name="diffview"]')) {
    radio.checked = radio.value === settings.diffView;
    radio.onchange = () => {
      settings.diffView = radio.value;
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
      // Re-render the open diff with the new layout.
      if (lastDiff && !$("results").hidden) renderDiff(lastDiff, lastBaseVersion, lastIsPastVersion);
    };
  }

  refreshUi();
});

/* ---------- snapshot engine ---------- */

// Convert 0-based column index to Excel letters: 0->A, 27->AB
function colLetter(n) {
  let s = "";
  n = n + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Photograph the whole workbook.
// Returns { takenAt, sheets: { <sheetId>: { name, cells: { "B3": {v, f?} } } } }
async function takeSnapshot() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,items/name");
    await context.sync();

    const pending = [];
    for (const ws of sheets.items) {
      const range = ws.getUsedRangeOrNullObject(true); // true = values only
      range.load("values,formulas,rowIndex,columnIndex,isNullObject");
      pending.push({ ws, range });
    }
    await context.sync();

    const snapshot = { takenAt: new Date().toISOString(), sheets: {} };
    for (const { ws, range } of pending) {
      const cells = {};
      if (!range.isNullObject) {
        const { values, formulas, rowIndex, columnIndex } = range;
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const v = values[r][c];
            const f = formulas[r][c];
            const isFormula = typeof f === "string" && f.startsWith("=");
            if ((v === "" || v === null) && !isFormula) continue; // skip empty cells
            const addr = colLetter(columnIndex + c) + (rowIndex + r + 1);
            const cell = { v: String(v) };
            if (isFormula) cell.f = f;
            cells[addr] = cell;
          }
        }
      }
      snapshot.sheets[ws.id] = { name: ws.name, cells };
    }
    return snapshot;
  });
}

/* ---------- diff engine ---------- */

// Compare two snapshots. Sheets are matched by internal sheet id, so a renamed
// sheet is detected as a rename — not as delete + add.
// Cell changes carry type: "add" | "edit" | "del". Sheet events carry
// kind: "add" | "del" | "rename".
function diffSnapshots(oldSnap, newSnap) {
  const result = { sheetEvents: [], sheetBlocks: [], totalChanges: 0 };

  const oldIds = Object.keys(oldSnap.sheets);
  const newIds = Object.keys(newSnap.sheets);

  for (const id of oldIds) {
    if (!newSnap.sheets[id]) {
      result.sheetEvents.push({ kind: "del", text: `Sheet "${oldSnap.sheets[id].name}" was deleted` });
      result.totalChanges++;
    }
  }
  for (const id of newIds) {
    if (!oldSnap.sheets[id]) {
      result.sheetEvents.push({ kind: "add", id, text: `Sheet "${newSnap.sheets[id].name}" was added` });
      result.totalChanges++;
    } else if (oldSnap.sheets[id].name !== newSnap.sheets[id].name) {
      result.sheetEvents.push({
        kind: "rename",
        id,
        text: `Sheet "${oldSnap.sheets[id].name}" was renamed to "${newSnap.sheets[id].name}"`,
      });
      result.totalChanges++;
    }
  }

  for (const id of newIds) {
    // A sheet that didn't exist in the base compares against empty, so all of
    // its cells show up as additions (the "sheet added" event is separate).
    const isNew = !oldSnap.sheets[id];
    const oldSheet = oldSnap.sheets[id] || { cells: {} };
    const newSheet = newSnap.sheets[id];
    const edited = [];
    const recalculated = [];

    const addrs = new Set([...Object.keys(oldSheet.cells), ...Object.keys(newSheet.cells)]);
    for (const addr of addrs) {
      const a = oldSheet.cells[addr];
      const b = newSheet.cells[addr];
      const aText = a ? (a.f || a.v) : "";
      const bText = b ? (b.f || b.v) : "";

      if ((a?.f || null) !== (b?.f || null) || (!a?.f && !b?.f && aText !== bText)) {
        // The user actually typed something different (value or formula).
        const type = aText === "" ? "add" : bText === "" ? "del" : "edit";
        edited.push({ addr, old: aText, new: bText, type });
      } else if (a?.f && b?.f && a.v !== b.v) {
        // Same formula, different result — changed because another cell changed.
        recalculated.push({ addr, old: a.v, new: b.v, type: "recalc" });
      }
    }

    if (edited.length || recalculated.length) {
      sortByAddress(edited);
      sortByAddress(recalculated);
      result.sheetBlocks.push({ id, name: newSheet.name, isNew, edited, recalculated });
      result.totalChanges += edited.length;
    }
  }

  return result;
}

function sortByAddress(list) {
  const parse = (addr) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return [Number(m[2]), col];
  };
  list.sort((x, y) => {
    const [xr, xc] = parse(x.addr);
    const [yr, yc] = parse(y.addr);
    return xr - yr || xc - yc;
  });
}

/* ---------- in-grid highlighting ---------- */

// Outline changed cells in the grid — colored border + font by change type,
// never the fill — remembering each cell's original border and font color so
// "Un-highlight" can restore them.
const HL_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];

async function highlightDiff(diff) {
  return Excel.run(async (context) => {
    // Load the whole collection once and use its real proxies — id lookups via
    // getItemOrNullObject can miss sheets on some hosts, which left every
    // sheet but the first un-highlighted.
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,items/name");
    await context.sync();
    const byId = new Map(sheets.items.map((w) => [w.id, w]));
    const byName = new Map(sheets.items.map((w) => [w.name, w]));

    const targets = [];
    const sheetInfo = new Map(); // ws proxy -> { sheetId, isNew }
    for (const block of diff.sheetBlocks) {
      const ws = byId.get(block.id) || byName.get(block.name);
      if (!ws) continue; // sheet no longer exists
      sheetInfo.set(ws, { sheetId: ws.id, isNew: !!block.isNew });
      for (const ch of [...block.edited, ...block.recalculated]) {
        if (targets.length >= MAX_HIGHLIGHTS) break;
        targets.push({ ws, addr: ch.addr, color: HL_COLORS[ch.type], type: ch.type, sheetId: ws.id });
      }
    }
    // Sheets whose change is the sheet itself, not cells: a brand-new sheet
    // (green) or a renamed one (amber) still gets its tab tinted even when it
    // has no cell changes.
    for (const ev of diff.sheetEvents || []) {
      if (!ev.id || (ev.kind !== "add" && ev.kind !== "rename")) continue;
      const ws = byId.get(ev.id);
      if (ws && !sheetInfo.has(ws)) sheetInfo.set(ws, { sheetId: ws.id, isNew: ev.kind === "add" });
    }

    // Pass 1: read original font color and border edges, plus each affected
    // sheet's tab color so the tab itself can be tinted too.
    for (const ws of sheetInfo.keys()) ws.load("tabColor");
    const reads = targets.map((t) => {
      const range = t.ws.getRange(t.addr);
      range.load("format/font/color");
      const borders = HL_EDGES.map((edge) => {
        const b = range.format.borders.getItem(edge);
        b.load("style,color,weight");
        return b;
      });
      return { ...t, range, borders };
    });
    await context.sync();

    // Pass 2: remember originals, then outline.
    // Formatting that exactly matches our highlight palette is almost certainly
    // a leftover frame from a lost session — record it as "no formatting" so
    // re-highlighting heals it instead of making it permanent.
    const hlSet = new Set(
      [...Object.values(HL_COLORS), ...Object.values(TAB_COLORS)].map((c) => c.toUpperCase())
    );
    const isOurs = (color) => hlSet.has(String(color || "").toUpperCase());
    const originals = [];
    const typesUsed = new Set();
    for (const t of reads) {
      originals.push({
        sheetId: t.sheetId,
        addr: t.addr,
        fontColor: isOurs(t.range.format.font.color) ? "#000000" : t.range.format.font.color,
        borders: t.borders.map((b, i) =>
          b.style === "Continuous" && isOurs(b.color)
            ? { edge: HL_EDGES[i], style: "None" }
            : { edge: HL_EDGES[i], style: b.style, color: b.color, weight: b.weight }
        ),
      });
      t.range.format.font.color = t.color;
      for (const edge of HL_EDGES) {
        const b = t.range.format.borders.getItem(edge);
        b.style = "Continuous";
        b.color = t.color;
        b.weight = "Medium";
      }
      typesUsed.add(t.type);
    }
    // Tab color: green for a brand-new sheet, amber for a modified one —
    // regardless of what kind of cell changes it contains.
    for (const [ws, info] of sheetInfo) {
      originals.push({ sheetId: info.sheetId, isTab: true, tabColor: isOurs(ws.tabColor) ? "" : ws.tabColor });
      ws.tabColor = info.isNew ? TAB_COLORS.add : TAB_COLORS.edit;
    }
    await context.sync();
    return { originals, typesUsed: [...typesUsed] };
  });
}

async function restoreFills(originals) {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id");
    await context.sync();
    const byId = new Map(sheets.items.map((w) => [w.id, w]));
    for (const o of originals) {
      const ws = byId.get(o.sheetId);
      if (!ws) continue; // sheet was deleted since highlighting
      if (o.isTab) {
        ws.tabColor = o.tabColor || ""; // "" = default (no tab color)
        continue;
      }
      const range = ws.getRange(o.addr);
      // Excel reports default text as black; writing it back is visually identical.
      range.format.font.color = o.fontColor || "#000000";
      for (const ob of o.borders) {
        const b = range.format.borders.getItem(ob.edge);
        if (!ob.style || ob.style === "None") {
          b.style = "None";
        } else {
          b.style = ob.style;
          b.color = ob.color;
          b.weight = ob.weight;
        }
      }
    }
    await context.sync();
  });
}

/* ---------- restore engine ---------- */

// Rebuild the workbook to match a snapshot: recreate missing sheets, delete
// sheets that didn't exist then, restore names, clear contents and write back
// every stored cell. Restores values + formulas only — not formatting.
async function restoreSnapshot(snap) {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,items/name");
    await context.sync();

    const targetIds = Object.keys(snap.sheets);
    const currentById = new Map(sheets.items.map((ws) => [ws.id, ws]));

    // Recreate sheets that no longer exist (temp names; final names set below).
    const mapping = [];
    targetIds.forEach((id, i) => {
      const target = snap.sheets[id];
      let ws = currentById.get(id);
      if (!ws) ws = sheets.add("esr_new_" + i);
      mapping.push({ ws, target });
    });
    // Delete sheets that didn't exist in the target version. Safe: every
    // target sheet already exists at this point, so the workbook keeps >=1.
    for (const [id, ws] of currentById) {
      if (!snap.sheets[id]) ws.delete();
    }
    await context.sync();

    // Two-phase rename so name swaps between sheets can't collide.
    mapping.forEach((m, i) => { m.ws.name = "esr_tmp_" + i; });
    await context.sync();
    mapping.forEach((m) => { m.ws.name = m.target.name; });

    // Clear current contents (keep formatting), then write stored cells back.
    const usedRanges = mapping.map((m) => {
      const used = m.ws.getUsedRangeOrNullObject();
      used.load("isNullObject");
      return used;
    });
    await context.sync();
    usedRanges.forEach((used) => {
      if (!used.isNullObject) used.clear(Excel.ClearApplyTo.contents);
    });
    for (const m of mapping) {
      for (const [addr, cell] of Object.entries(m.target.cells)) {
        m.ws.getRange(addr).formulas = [[cell.f || cell.v]];
      }
    }
    await context.sync();
  });
}

async function onRevert(version) {
  try {
    setStatus("Reverting…");
    if (highlightState) await onClearHighlights();
    hideResults();
    const before = latestVersion();
    await restoreSnapshot(version.snapshot);
    // The revert becomes the newest version, so history only moves forward.
    const current = await takeSnapshot();
    const changeCount = before ? diffSnapshots(before.snapshot, current).totalChanges : 0;
    saveVersion(current, `Revert to “${version.message}”`, changeCount);
    refreshUi();
    setStatus("Reverted — saved as the newest version. (Values and formulas restored; formatting is not tracked.)", "ok");
  } catch (e) {
    setStatus("Revert failed: " + (e.message || e), "err");
  }
}

// Preview a revert before doing it: the diff runs in reverse (current state →
// that version), so it shows exactly what the revert will delete, restore and
// change. The save card becomes the confirm button, and the version's history
// row shows "Cancel revert" until the preview is left.
let revertPreviewVersion = null;

async function cancelRevertPreview() {
  revertPreviewVersion = null;
  renderHistory();
  // Back to the normal forward diff with the regular Save version card.
  await onReview(null);
}

async function onRevertPreview(version) {
  try {
    setStatus("Computing what a revert would change…");
    if (highlightState) await onClearHighlights();
    const current = await takeSnapshot();
    const diff = diffSnapshots(current, version.snapshot);
    lastDiff = diff;
    lastCurrentSnapshot = current;
    lastBaseVersion = version;
    lastIsPastVersion = false;
    renderDiff(diff, version, false);
    const note = $("compare-note");
    note.hidden = false;
    note.textContent = diff.totalChanges === 0
      ? "Reverting to this version would change nothing — the workbook already matches it."
      : `Revert preview — this is what going back to “${version.message}” (${new Date(version.takenAt).toLocaleString()}) will do to the current workbook.`;
    $("version-message").hidden = true;
    const btn = $("btn-accept");
    btn.classList.add("revert");
    btn.innerHTML = icon("undo") + " Revert to this version";
    btn.onclick = () => onRevert(version);
    revertPreviewVersion = version;
    renderHistory();
    setStatus("");
    if (diff.totalChanges > 0) await flashHighlight(5);
  } catch (e) {
    setStatus("Revert preview failed: " + (e.message || e), "err");
  }
}

/* ---------- storage: version history ---------- */

// Stored shape: { versions: [ {takenAt, message, changeCount, snapshot}, ... ] }
// newest first, capped at MAX_VERSIONS.
function loadStore() {
  const raw = localStorage.getItem(workbookKey);
  return raw ? JSON.parse(raw) : { versions: [] };
}

function persistStore(store) {
  // If the browser storage quota is hit, drop oldest versions until it fits.
  while (true) {
    try {
      localStorage.setItem(workbookKey, JSON.stringify(store));
      return;
    } catch (e) {
      if (store.versions.length <= 1) throw e;
      store.versions.pop();
    }
  }
}

// Phase A stored a bare snapshot ({takenAt, sheets}); wrap it as version #1.
function migrateOldFormat() {
  const raw = localStorage.getItem(workbookKey);
  if (!raw) return;
  const parsed = JSON.parse(raw);
  if (parsed.sheets) {
    persistStore({
      versions: [{ takenAt: parsed.takenAt, message: "(first snapshot)", changeCount: 0, snapshot: parsed }],
    });
  }
}

function latestVersion() {
  return loadStore().versions[0] || null;
}

function saveVersion(snapshot, message, changeCount) {
  const store = loadStore();
  store.versions.unshift({ takenAt: snapshot.takenAt, message, changeCount, snapshot });
  if (store.versions.length > MAX_VERSIONS) store.versions.length = MAX_VERSIONS;
  persistStore(store);
}

/* ---------- UI ---------- */

function timeAgo(iso) {
  const s = Math.max(0, (Date.now() - new Date(iso).getTime()) / 1000);
  if (s < 60) return "just now";
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  if (s < 172800) return "yesterday";
  return new Date(iso).toLocaleDateString();
}

function refreshUi() {
  const store = loadStore();
  const latest = store.versions[0] || null;
  $("empty-state").hidden = !!latest;
  $("main-actions").hidden = !latest;
  $("snapshot-info").textContent = latest
    ? `${store.versions.length} version${store.versions.length === 1 ? "" : "s"} · ${timeAgo(latest.takenAt)} — “${latest.message}”`
    : "Not tracking this workbook yet";
  // The line ellipsizes in the narrow pane — hovering reveals the full text.
  $("snapshot-info").title = $("snapshot-info").textContent;
  renderHistory();
}

let statusTimer = null;
function setStatus(text, cls) {
  clearTimeout(statusTimer);
  $("status").textContent = text;
  $("status").className = "status " + (cls || "");
  // Success messages fade away on their own; errors stay until replaced.
  if (cls === "ok" && text) {
    statusTimer = setTimeout(() => {
      $("status").textContent = "";
      $("status").className = "status";
    }, 10000);
  }
}

function hideResults() {
  $("results").hidden = true;
  if (revertPreviewVersion) {
    revertPreviewVersion = null;
    renderHistory();
  }
  if (highlightState) onClearHighlights();
}

async function onFirstSnapshot() {
  try {
    setStatus("Photographing workbook…");
    const snap = await takeSnapshot();
    saveVersion(snap, "(first snapshot)", 0);
    hideResults();
    refreshUi();
    setStatus("First snapshot saved. Edit the workbook, then click Review changes.", "ok");
  } catch (e) {
    setStatus("Snapshot failed: " + (e.message || e), "err");
  }
}

// compareVersion = a history entry to diff against, or null for the latest.
async function onReview(compareVersion) {
  try {
    const base = compareVersion || latestVersion();
    if (!base) {
      return onFirstSnapshot();
    }
    if (highlightState) await onClearHighlights();
    setStatus("Comparing…");
    const current = await takeSnapshot();
    const diff = diffSnapshots(base.snapshot, current);
    lastDiff = diff;
    lastCurrentSnapshot = current;
    lastBaseVersion = base;
    lastIsPastVersion = !!compareVersion;
    renderDiff(diff, base, !!compareVersion);
    setStatus("");
    if (diff.totalChanges > 0) await flashHighlight(5);
  } catch (e) {
    setStatus("Review failed: " + (e.message || e), "err");
  }
}

// The original formatting is also persisted, so a highlight left behind when
// the pane closes can be cleaned up on the next open.
function hlKey() { return workbookKey + ":hl"; }

// "Review changes" flashes the highlight for a few seconds; the un-highlight
// button counts them down and any manual action cancels the timer.
let hlTimer = null;
let hlTicker = null;

function stopHlTimer() {
  clearTimeout(hlTimer);
  clearInterval(hlTicker);
  hlTimer = hlTicker = null;
  $("btn-clear-highlight").innerHTML = icon("x") + " Un-highlight";
}

async function flashHighlight(seconds) {
  await onHighlight();
  if (!highlightState || !highlightState.length) return;
  const btn = $("btn-clear-highlight");
  let remaining = seconds;
  btn.innerHTML = icon("x") + ` Un-highlight (${remaining}s)`;
  hlTicker = setInterval(() => {
    remaining--;
    if (remaining > 0) btn.innerHTML = icon("x") + ` Un-highlight (${remaining}s)`;
  }, 1000);
  hlTimer = setTimeout(() => onClearHighlights(), seconds * 1000);
}

async function onHighlight() {
  if (!lastDiff) return;
  try {
    // Never sample "original" formatting on top of a leftover highlight.
    await clearStaleHighlight();
    const result = await highlightDiff(lastDiff);
    highlightState = result.originals;
    try { localStorage.setItem(hlKey(), JSON.stringify(result.originals)); } catch (e) { /* best effort */ }
    $("btn-highlight").hidden = true;
    $("btn-clear-highlight").hidden = false;
    renderLegend(result.typesUsed);
  } catch (e) {
    setStatus("Highlight failed: " + (e.message || e), "err");
  }
}

async function onClearHighlights() {
  stopHlTimer();
  if (!highlightState) return;
  try {
    await restoreFills(highlightState);
    highlightState = null;
    localStorage.removeItem(hlKey());
    $("btn-highlight").hidden = false;
    $("btn-clear-highlight").hidden = true;
    $("hl-legend").hidden = true;
  } catch (e) {
    setStatus("Could not restore cell formatting: " + (e.message || e), "err");
  }
}

// A highlight left behind when the pane was last closed: restore the original
// formatting now that we're back.
async function clearStaleHighlight() {
  const raw = localStorage.getItem(hlKey());
  if (!raw) return;
  try {
    await restoreFills(JSON.parse(raw));
  } catch (e) {
    setStatus("Could not remove a leftover highlight: " + (e.message || e), "err");
  }
  localStorage.removeItem(hlKey());
}

// Small color legend shown directly under the highlight button.
function renderLegend(typesUsed) {
  const legend = $("hl-legend");
  legend.innerHTML = "";
  for (const type of ["edit", "add", "del", "recalc"]) {
    if (!typesUsed.includes(type)) continue;
    const item = document.createElement("span");
    item.className = "lg-item";
    const dot = document.createElement("i");
    dot.className = "lg-dot";
    dot.style.background = HL_COLORS[type];
    item.appendChild(dot);
    item.appendChild(document.createTextNode(HL_LABELS[type]));
    legend.appendChild(item);
  }
  legend.hidden = false;
}

function renderDiff(diff, baseVersion, isPastVersion) {
  const summary = $("summary");
  const events = $("sheet-events");
  const cellBox = $("cell-changes");
  events.innerHTML = "";
  cellBox.innerHTML = "";
  $("hl-legend").hidden = true;

  const note = $("compare-note");
  note.hidden = !isPastVersion;
  if (isPastVersion) {
    note.textContent = `Comparing against the version from ${new Date(baseVersion.takenAt).toLocaleString()} (“${baseVersion.message}”), not the latest.`;
  }

  // Summary pills: total + breakdown by kind of change.
  summary.innerHTML = "";
  const pill = (html, cls) => {
    const el = document.createElement("span");
    el.className = "pill " + cls;
    el.innerHTML = html;
    summary.appendChild(el);
  };
  let added = 0, edited = 0, emptied = 0;
  for (const block of diff.sheetBlocks) {
    for (const ch of block.edited) {
      if (ch.type === "add") added++;
      else if (ch.type === "del") emptied++;
      else edited++;
    }
  }
  if (diff.totalChanges === 0) {
    pill(icon("check") + " No changes", "zero");
  } else {
    // Diffstat strip: total on the left, a thin proportional color bar in the
    // middle, compact +/~/− counters on the right.
    pill(`${diff.totalChanges} change${diff.totalChanges === 1 ? "" : "s"}`, "total");

    const bar = document.createElement("span");
    bar.className = "stat-bar";
    for (const [n, cls] of [[added, "add"], [edited, "edit"], [emptied, "del"]]) {
      if (!n) continue;
      const seg = document.createElement("i");
      seg.className = "seg seg-" + cls;
      seg.style.flexGrow = n;
      bar.appendChild(seg);
    }
    summary.appendChild(bar);

    const nums = document.createElement("span");
    nums.className = "stat-nums";
    const num = (html, cls, title) => {
      const b = document.createElement("b");
      b.className = "sn-" + cls;
      b.innerHTML = html;
      b.title = title;
      nums.appendChild(b);
    };
    if (added) num("+" + added, "add", added + " added");
    if (edited) num("~" + edited, "edit", edited + " edited");
    if (emptied) num("−" + emptied, "del", emptied + " removed");
    if (diff.sheetEvents.length) {
      num(icon("grid") + diff.sheetEvents.length, "sheet",
        diff.sheetEvents.length + " sheet-level change" + (diff.sheetEvents.length === 1 ? "" : "s"));
    }
    summary.appendChild(nums);
  }
  $("btn-highlight").hidden = diff.totalChanges === 0;
  $("btn-clear-highlight").hidden = true;

  const EVENT_ICON = { add: "plus", del: "trash", rename: "pencil" };
  for (const ev of diff.sheetEvents) {
    const div = document.createElement("div");
    div.className = "sheet-event ev-" + ev.kind;
    div.innerHTML = icon(EVENT_ICON[ev.kind]) + `<span></span>`;
    div.lastChild.textContent = ev.text;
    events.appendChild(div);
  }

  // Two layouts. "stacked": every sheet in one list, unfolded, each header
  // folds its sheet away. "tabs": a strip like Excel's own sheet tabs — one
  // sheet visible at a time, click a tab to move between them.
  const tabsView = settings.diffView === "tabs";
  let tabBar = null;
  let tabFrame = null;
  if (tabsView && diff.sheetBlocks.length) {
    // One frame, GitHub-style: the tab strip is the frame's header and the
    // chosen sheet's changes open directly below it, inside the same border.
    tabFrame = document.createElement("div");
    tabFrame.className = "tab-frame";
    tabBar = document.createElement("div");
    tabBar.className = "sheet-tabs";
    // Strip chrome: ‹ › arrows appear only when there are more tabs than fit.
    const tabsRow = document.createElement("div");
    tabsRow.className = "tabs-row";
    const navBtn = (dir) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tab-nav" + (dir < 0 ? " left" : "");
      b.innerHTML = icon("chevron");
      b.title = dir < 0 ? "Scroll tabs left" : "Scroll tabs right";
      b.onclick = () => tabBar.scrollBy({ left: dir * 140, behavior: "smooth" });
      return b;
    };
    const navL = navBtn(-1);
    const navR = navBtn(1);
    tabsRow.append(navL, tabBar, navR);
    tabFrame.appendChild(tabsRow);
    cellBox.appendChild(tabFrame);
    // Runs after layout (and after the results section is unhidden below).
    requestAnimationFrame(() => {
      navL.hidden = navR.hidden = tabBar.scrollWidth <= tabBar.clientWidth + 1;
      const act = tabBar.querySelector(".sheet-tab.active");
      if (act) act.scrollIntoView({ inline: "nearest", block: "nearest" });
    });
    // Keep the previously selected tab across re-renders when it still exists.
    if (!diff.sheetBlocks.some((b) => b.id === activeSheetTab)) {
      activeSheetTab = diff.sheetBlocks[0].id;
    }
  }
  let shown = 0;
  for (const block of diff.sheetBlocks) {
    const wrap = document.createElement("div");
    wrap.className = "sheet-block";
    if (tabsView) {
      wrap.classList.add("tab-page");
      if (block.id === activeSheetTab) wrap.classList.add("active");
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "sheet-tab" + (block.id === activeSheetTab ? " active" : "");
      tab.innerHTML = `<span></span><span class="tab-count"></span>`;
      tab.firstChild.textContent = block.name;
      tab.lastChild.textContent = block.edited.length;
      tab.title = block.name + " — " + block.edited.length + " change" + (block.edited.length === 1 ? "" : "s");
      tab.onclick = () => {
        activeSheetTab = block.id;
        for (const el of tabBar.children) el.classList.toggle("active", el === tab);
        for (const el of cellBox.querySelectorAll(".tab-page")) el.classList.toggle("active", el === wrap);
        tab.scrollIntoView({ inline: "nearest", block: "nearest", behavior: "smooth" });
      };
      tabBar.appendChild(tab);
    } else {
      const title = document.createElement("div");
      title.className = "sheet-name";
      // Stacked: open by default, header folds the sheet.
      wrap.classList.add("collapsible");
      title.innerHTML = icon("chevron", "chev") + `<span></span>`;
      title.children[1].textContent = block.name;
      const count = document.createElement("span");
      count.className = "sheet-count";
      const n = block.edited.length;
      count.textContent = `${n} change${n === 1 ? "" : "s"}`;
      title.appendChild(count);
      title.onclick = () => wrap.classList.toggle("collapsed");
      wrap.appendChild(title);
    }

    for (const ch of block.edited) {
      if (shown++ >= MAX_CHANGES_SHOWN) break;
      wrap.appendChild(renderChange(ch, block.id));
    }
    if (block.recalculated.length) {
      const noteEl = document.createElement("div");
      noteEl.className = "recalc-note";
      noteEl.innerHTML = icon("refresh") + `<span></span>`;
      noteEl.lastChild.textContent = ` ${block.recalculated.length} cell(s) recalculated because of the edits above:`;
      wrap.appendChild(noteEl);
      for (const ch of block.recalculated.slice(0, 20)) {
        wrap.appendChild(renderChange(ch, block.id));
      }
    }
    (tabsView ? tabFrame : cellBox).appendChild(wrap);
  }
  if (shown >= MAX_CHANGES_SHOWN) {
    const more = document.createElement("p");
    more.className = "truncate-note";
    more.textContent = `…only the first ${MAX_CHANGES_SHOWN} changes are shown.`;
    cellBox.appendChild(more);
  }

  // Reset the save card — a revert preview may have repurposed it. (A running
  // onRevertPreview re-sets the flag right after this renders.)
  if (revertPreviewVersion) {
    revertPreviewVersion = null;
    renderHistory();
  }
  $("version-message").value = "";
  $("version-message").hidden = false;
  $("btn-accept").classList.remove("revert");
  $("btn-accept").innerHTML = icon("check") + " Save version";
  $("btn-accept").onclick = async () => {
    const message = $("version-message").value.trim() || "(no message)";
    if (highlightState) await onClearHighlights();
    saveVersion(lastCurrentSnapshot, message, diff.totalChanges);
    hideResults();
    refreshUi();
    setStatus("Version saved.", "ok");
  };

  $("results").hidden = false;
}

// Jump Excel to a cell: activate its sheet and select the range.
// keepPaneFocus: the pane is about to open its inline editor for this cell,
// so keyboard focus must stay here instead of being handed to the grid.
async function goToCell(sheetId, addr, keepPaneFocus) {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItemOrNullObject(sheetId);
      ws.activate();
      ws.getRange(addr).select();
      await context.sync();
    });
    if (!keepPaneFocus) {
      // Hand keyboard focus back to the grid so the first keystroke starts
      // editing the selected cell (there is no API to enter edit mode itself).
      if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
      window.blur();
    }
    return true;
  } catch (e) {
    setStatus("Could not jump to " + addr + " — the sheet may have been deleted.", "err");
    return false;
  }
}

function renderChange(ch, sheetId) {
  const row = document.createElement("div");
  row.className = "change" + (ch.type === "recalc" ? " recalc" : "");
  const addr = document.createElement("button");
  addr.className = "addr t-" + ch.type;
  addr.textContent = ch.addr;
  addr.title = ch.type === "recalc"
    ? "Jump to " + ch.addr
    : "Jump to " + ch.addr + " and edit its value";
  // Don't let the button take keyboard focus on click — the editor (or the
  // grid, for recalc rows) gets it instead.
  addr.onmousedown = (e) => e.preventDefault();
  const body = document.createElement("span");
  // "del" (emptied) shows only the removed value — the cell is now truly empty.
  // "add" shows only the new value. "edit"/"recalc" show old → new.
  if (ch.old !== "" && ch.type !== "add") {
    const oldEl = document.createElement("span");
    oldEl.className = "old";
    oldEl.textContent = ch.old;
    body.appendChild(oldEl);
  }
  let newEl = null;
  if (ch.new !== "" && ch.type !== "del") {
    newEl = document.createElement("span");
    newEl.className = "new";
    newEl.textContent = ch.new;
    // recalc rows show a formula's result — editing that would overwrite the
    // formula with a literal, so they stay read-only
    if (ch.type !== "recalc") makeValueEditable(newEl, sheetId, ch.addr);
    body.appendChild(newEl);
  }
  // A task pane can never hand keyboard focus to the grid, so "jump then type"
  // can't work. Instead, the address chip jumps AND opens the pane's own
  // inline editor for that cell — type, Enter, and it's written to the cell.
  addr.onclick = async () => {
    if (ch.type === "recalc") return goToCell(sheetId, ch.addr);
    let target = newEl;
    if (!target) {
      // "del" row: the cell is empty now — give it an editor to type a new
      // value into; removed again if nothing is entered.
      target = document.createElement("span");
      target.className = "new";
      makeValueEditable(target, sheetId, ch.addr);
      body.appendChild(target);
      target._ephemeral = true;
    }
    if (!(await goToCell(sheetId, ch.addr, true))) return;
    openCellEditor(target, sheetId, ch.addr);
  };
  row.appendChild(addr);
  row.appendChild(body);
  return row;
}

// The task pane cannot hand keyboard focus to the grid, so editing happens
// here instead: click the current value (or the address chip), type, Enter
// writes it to the cell.
function makeValueEditable(newEl, sheetId, addr) {
  newEl.title = "Click to edit " + addr;
  newEl.onclick = () => openCellEditor(newEl, sheetId, addr);
}

function openCellEditor(newEl, sheetId, addr) {
  const existing = newEl.querySelector("input");
  if (existing) return existing.focus();
  const current = newEl.textContent;
  const input = document.createElement("input");
  input.className = "cell-edit";
  input.value = current;
  newEl.textContent = "";
  newEl.appendChild(input);
  input.focus();
  input.select();

  let finished = false;
  const finish = async (commit) => {
    if (finished) return;
    finished = true;
    const val = input.value;
    input.remove();
    // An editor conjured onto a "del" row disappears again unless a value
    // was actually written.
    if (newEl._ephemeral && (!commit || val === "")) return newEl.remove();
    newEl.textContent = commit ? val : current;
    if (!commit || val === current) return;
    try {
      await writeCell(sheetId, addr, val);
      setStatus(addr + " updated.", "ok");
      showUndoChip(newEl, sheetId, addr, current);
    } catch (e) {
      if (newEl._ephemeral) newEl.remove();
      else newEl.textContent = current;
      setStatus("Could not update " + addr + ": " + (e.message || e), "err");
    }
  };
  input.onkeydown = (e) => {
    if (e.key === "Enter") { e.preventDefault(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); finish(false); }
  };
  input.onblur = () => finish(false);
}

// Write one value into a cell, the way typing would: "=" starts a formula,
// numbers land as numbers, "" clears the cell.
async function writeCell(sheetId, addr, val) {
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItemOrNullObject(sheetId);
    const range = ws.getRange(addr);
    if (val === "") {
      range.clear(Excel.ClearApplyTo.contents);
    } else if (val.startsWith("=")) {
      range.formulas = [[val]];
    } else {
      const num = Number(val);
      range.values = [[val.trim() !== "" && !isNaN(num) ? num : val]];
    }
    range.select();
    await context.sync();
  });
}

// After an edit made from the pane, an ↩ chip appears on that row: one click
// restores what the cell held before the last pane edit. It lives only as
// long as the review panel — deeper time travel is the version history's job.
function showUndoChip(newEl, sheetId, addr, prevVal) {
  const row = newEl.closest(".change");
  if (!row) return;
  let btn = row.querySelector(".undo-edit");
  if (!btn) {
    btn = document.createElement("button");
    btn.className = "undo-edit";
    btn.innerHTML = icon("undo");
    row.appendChild(btn);
  }
  btn.title = "Undo this edit — restore " + (prevVal === "" ? "the empty cell" : "“" + prevVal + "”");
  btn.onclick = async () => {
    btn.disabled = true;
    try {
      await writeCell(sheetId, addr, prevVal);
      btn.remove();
      // A "del"-row editor that was reverted to empty vanishes again.
      if (newEl._ephemeral && prevVal === "") newEl.remove();
      else newEl.textContent = prevVal;
      setStatus(addr + " restored to what it was before your edit.", "ok");
    } catch (e) {
      btn.disabled = false;
      setStatus("Could not restore " + addr + ": " + (e.message || e), "err");
    }
  };
}

function renderHistory() {
  const store = loadStore();
  const section = $("history-section");
  const list = $("history-list");
  list.innerHTML = "";

  const rows = [...store.versions];
  rows.sort((a, b) => (a.takenAt < b.takenAt ? 1 : -1));

  section.hidden = rows.length === 0;
  $("history-count").textContent = rows.length ? `(${rows.length})` : "";

  rows.forEach((v, i) => {
    const row = document.createElement("div");
    row.className = "history-row" + (i === 0 ? " latest" : "");

    const info = document.createElement("div");
    info.className = "history-info";
    const msg = document.createElement("div");
    msg.className = "history-msg";
    msg.textContent = v.message;
    msg.title = v.message;
    const when = document.createElement("div");
    when.className = "history-when";
    when.textContent =
      timeAgo(v.takenAt) +
      (i === 0 ? " · latest" : v.changeCount ? ` · ${v.changeCount} change${v.changeCount === 1 ? "" : "s"}` : "");
    when.title = new Date(v.takenAt).toLocaleString();
    info.appendChild(msg);
    info.appendChild(when);

    const actions = document.createElement("div");
    actions.className = "history-actions";

    const btnDiff = document.createElement("button");
    btnDiff.className = "ghost small";
    btnDiff.textContent = "Diff vs now";
    btnDiff.title = "Show what changed between this version and the workbook right now";
    btnDiff.onclick = () => onReview(v);

    // Revert arms on first click, executes on second (no popups in the pane).
    const btnRevert = document.createElement("button");
    btnRevert.className = "ghost small";
    if (revertPreviewVersion && revertPreviewVersion.takenAt === v.takenAt) {
      btnRevert.classList.add("danger");
      btnRevert.innerHTML = icon("x") + " Cancel revert";
      btnRevert.title = "Leave the revert preview without changing anything.";
      btnRevert.onclick = () => cancelRevertPreview();
    } else {
      btnRevert.innerHTML = icon("undo") + " Revert";
      btnRevert.title = "Preview what reverting to this version will change, then confirm.";
      btnRevert.onclick = () => onRevertPreview(v);
    }

    actions.appendChild(btnDiff);
    actions.appendChild(btnRevert);
    row.appendChild(info);
    row.appendChild(actions);
    list.appendChild(row);
  });
}

/* ---------- save-bar float detection ---------- */

// The save bar renders as a floating card only while it is pinned to the
// bottom of the viewport. A zero-height sentinel sits right below its natural
// position: sentinel off-screen -> the bar is stuck -> card chrome on.
(() => {
  const confirmEl = document.querySelector(".confirm");
  if (!confirmEl || typeof IntersectionObserver === "undefined") return;
  const sentinel = document.createElement("div");
  confirmEl.after(sentinel);
  new IntersectionObserver(([entry]) => {
    confirmEl.classList.toggle("floating", !entry.isIntersecting);
  }).observe(sentinel);
})();

// Same trick for the summary bar, pinned to the top: sentinel above its
// natural position scrolled off-screen -> the bar is stuck -> lifted chrome.
(() => {
  const bar = document.querySelector(".summary-bar");
  if (!bar || typeof IntersectionObserver === "undefined") return;
  const sentinel = document.createElement("div");
  bar.before(sentinel);
  new IntersectionObserver(([entry]) => {
    bar.classList.toggle("floating", !entry.isIntersecting);
  }).observe(sentinel);
})();
