# Save & Review — Excel add-in prototype (Phase C)

An Excel add-in that shows you **exactly what changed** in your workbook —
cell by cell, old value → new value, plus added/renamed/deleted sheets —
before you accept it as a new version. Like reviewing a diff on GitHub,
but inside Excel.

This is the Phase A prototype: the snapshot + diff engine with a task pane UI.
Versions are stored locally; the cloud backend comes in Phase C.

---

## How it works (plain English)

The engine is a **camera**:

1. **📸 Take snapshot** — photographs every sheet: each used cell's value and
   formula. This photo is your *baseline*.
2. You edit the workbook normally.
3. **🔍 Review changes** — takes a new photo and compares it to the baseline:
   - Cells you actually edited (value or formula changed) — shown red → green.
   - Cells that only *recalculated* because of your edits — shown separately, dimmed.
   - Sheets that were added, deleted, or renamed. Renames are detected properly
     (sheets are matched by Excel's internal sheet ID, not by name).
4. **✅ Save version** — type a short message ("what did you change?") and
   accept; the new photo joins the **version history**. In the product version,
   this is the moment a version uploads to the cloud.

Phase B features:

- **Version history** — every accepted save is kept (newest first, up to 20),
  each with its timestamp, message, and change count. **Diff vs now** on any
  past version shows everything that changed since it.
- **Highlight in sheet** — paints the changed cells directly in the grid,
  colored by change type (amber = edited, green = added, red = removed,
  blue = recalculated), with a color legend under the button. Original cell
  fills are restored on "Un-highlight" or when the review closes.
- **Revert** — every history entry can restore the workbook to that version's
  state (sheets recreated/deleted/renamed, all cell values and formulas
  written back). The revert is saved as the *newest* version, so history only
  moves forward. Restores values + formulas only — formatting is not tracked.
- **Settings (sliders icon)** — choose the diff layout: Expanded (everything
  visible) or Compact (click a sheet header to unfold its changes).

## Phase C: the cloud

A second local service — the **version server** (`server/index.js`, port 3003)
— stores versions per user account:

- **Sign in** from the settings panel (sliders icon → Cloud account). Dev-mode
  sign-in for now: email only, no password. The production version swaps this
  for Microsoft sign-in; the API shape stays the same.
- **Every "Save version" uploads to the server.** The pane keeps a local cache
  (last 20 versions in `localStorage`) and works fully offline — anything
  saved while the server is unreachable syncs up automatically later, and
  versions made before signing in are uploaded on first sign-in.
- **History merges both worlds**: cached versions plus older cloud-only ones
  (marked with a cloud icon). Diff/revert on a cloud-only version fetches its
  snapshot on demand.
- **Workbook identity**: a UUID stored in the workbook's own document settings
  (inside the file), so renaming or moving the file keeps its history.
- Storage is JSON files under `server/data/` (gitignored). Deliberately
  simple — swaps for object storage when this moves off localhost.

Run it alongside the pane server:

```bash
npm run server
```

### GitHub mirror (optional, per user)

The pragmatic combination: platform storage is the default for everyone; users
who want to own their data can additionally connect **their own GitHub** from
the settings panel. The server then:

1. validates the pasted personal access token (dev flow — production uses
   GitHub OAuth),
2. creates a private repo `save-review-versions` in the user's account,
3. backfills all existing versions, and from then on commits every new
   version as `<workbook>-<id>/<timestamp>.json`.

Versions are immutable, so mirroring is idempotent (an existing file means
"already mirrored"). GitHub's own UI doubles as a version browser. Mirror
status is shown under the connect button in settings.

## Project layout

| File | What it is |
|---|---|
| `manifest.xml` | The "ID card" that tells Excel about the add-in: its name, its ribbon button, and where its web page lives (`https://localhost:3002`). |
| `src/taskpane.html/.css/.js` | The task pane — a small web page that runs in Excel's side panel. `taskpane.js` holds the snapshot + diff engine. |
| `server.js` | Tiny HTTPS server that serves those files to Excel while developing. |
| `sideload.js` | Copies the manifest into Excel-for-Mac's magic folder so Excel picks it up. |
| `assets/` | Ribbon icons. |

## Running it (first time — three steps)

**Step 1 — trust a local certificate** (one time only). Excel only loads
add-ins over HTTPS, so we create a certificate for `localhost` that your Mac
trusts. It will ask for your Mac password once:

```bash
cd excel-save-review
npm run certs
```

**Step 2 — start the dev server** (leave it running while you use the add-in):

```bash
npm run start
```

Check it works: open https://localhost:3002 in your browser — you should see
the Save & Review panel (with "This add-in only works in Excel", which is fine).

**Step 3 — tell Excel about the add-in** (one time only):

```bash
npm run sideload
```

Then **fully quit Excel (Cmd+Q) and reopen it**. Open any workbook and look for
the **Save & Review** button on the **Home** tab (if you don't see it:
**Insert tab → My Add-ins → dropdown arrow → Developer Add-ins → Save & Review**).

## Daily use

1. `npm run start` (if not already running)
2. Open your workbook → click **Save & Review** on the ribbon
3. Take a snapshot → edit cells → **Review changes** → see the diff → accept

## Known limits (by design — prototype)

- Versions are stored locally in the task pane's storage, not in the cloud
  (Phase C), and capped at 20 per workbook.
- Row/column *inserts* show up as many changed cells (everything shifted),
  not as "row inserted". Proper move detection is a later refinement.
- Very large sheets: the review lists the first 300 changes.
- It does not intercept Ctrl+S — the review is a deliberate button click
  (this is also the plan for the store version; Microsoft doesn't allow
  store add-ins to block the built-in save).

## Roadmap

- **Phase B** — ✅ done: version message box, in-grid highlighting, local
  version history with diff-vs-now.
- **Phase C** — cloud backend: sign-in, version history list, view/restore any
  past version.
- **Phase D** — Microsoft AppSource submission (Partner Center), billing.
