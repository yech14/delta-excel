// Save & Review — version server (Phase C).
//
// A small HTTPS API the task pane talks to:
//   POST /api/auth/login                          {email} -> {email, token}
//   GET  /api/workbooks                           list this user's workbooks
//   GET  /api/workbooks/:wb/versions              version metadata (no snapshots)
//   POST /api/workbooks/:wb/versions              {takenAt, message, changeCount, snapshot, name}
//   GET  /api/workbooks/:wb/versions/:id          one full version incl. snapshot
//
// Storage is plain JSON files under server/data/ — deliberately simple; swap
// for object storage / a database when this moves off localhost.
//
// AUTH IS DEV-MODE ONLY: email in, token out, no verification. The production
// version replaces this with Microsoft sign-in (the token check stays the same
// shape, which is why the rest of the API won't need to change).

const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");

const PORT = 3003;
const PANE_ORIGINS = ["https://localhost:3002", "https://127.0.0.1:3002"];
const MAX_BODY = 25 * 1024 * 1024; // 25 MB per request
const DATA_DIR = path.join(__dirname, "data");
const USERS_FILE = path.join(DATA_DIR, "users.json");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

fs.mkdirSync(DATA_DIR, { recursive: true });

function loadJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    return fallback;
  }
}
function saveJson(file, obj) {
  // Atomic write: a crash mid-write leaves the old file intact (rename is
  // atomic on macOS/Linux), never a half-written history.
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(obj));
  fs.renameSync(tmp, file);
}

/* ---------- users ---------- */

function getOrCreateUser(email) {
  const users = loadJson(USERS_FILE, {});
  const key = email.trim().toLowerCase();
  if (!users[key]) {
    users[key] = {
      id: crypto.createHash("sha256").update(key).digest("hex").slice(0, 16),
      token: crypto.randomBytes(24).toString("hex"),
    };
    saveJson(USERS_FILE, users);
  }
  return { email: key, ...users[key] };
}

function userFromToken(token) {
  if (!token) return null;
  const users = loadJson(USERS_FILE, {});
  for (const [email, u] of Object.entries(users)) {
    if (u.token === token) return { email, ...u };
  }
  return null;
}

/* ---------- GitHub mirror (optional, per user) ---------- */
// Mirrors every version into a private repo on the USER'S OWN GitHub account,
// in addition to platform storage. Dev flow: user pastes a personal access
// token; production swaps this for GitHub OAuth. Tokens live in users.json —
// acceptable for localhost dev only.

const GH_API = "https://api.github.com";
const GH_REPO_NAME = "save-review-versions";
const mirrorStatus = {}; // userId -> {state: "syncing"|"ok"|"error", detail}

// Central mirror: when server/config.json exists ({token, owner, repo}), ALL
// users' versions are mirrored into that one repo — one folder per user under
// history/. Users then never deal with tokens; they just sign in with email.
const CONFIG = loadJson(path.join(__dirname, "config.json"), null);
const CENTRAL = CONFIG && CONFIG.token
  ? { token: CONFIG.token, login: CONFIG.owner, repo: CONFIG.repo || GH_REPO_NAME }
  : null;

function emailSlug(email) {
  return String(email).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}

async function ghFetch(ghToken, method, apiPath, body) {
  return fetch(GH_API + apiPath, {
    method,
    headers: {
      Authorization: "Bearer " + ghToken,
      Accept: "application/vnd.github+json",
      "User-Agent": "save-and-review-dev",
      ...(body ? { "Content-Type": "application/json" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
}

function setUserGh(email, gh) {
  const users = loadJson(USERS_FILE, {});
  if (!users[email]) return;
  if (gh) users[email].gh = gh;
  else delete users[email].gh;
  saveJson(USERS_FILE, users);
}

function mirrorPath(wbName, wbId, v, userFolder) {
  const slug =
    String(wbName || "workbook")
      .toLowerCase()
      .replace(/\.[a-z0-9]+$/, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 40) || "workbook";
  // Everything lives under history/ so a mirror into a repo that also holds
  // other content (e.g. source code) stays tidy. Central mirroring adds one
  // folder per user.
  const base = userFolder ? `history/${userFolder}` : "history";
  return `${base}/${slug}-${String(wbId).slice(0, 8)}/${v.takenAt.replace(/[:.]/g, "-")}.json`;
}

// Which GitHub credentials mirror this user's versions: the central repo when
// configured, else the user's own connected account, else none.
function mirrorTarget(user) {
  if (CENTRAL) return { gh: CENTRAL, userFolder: emailSlug(user.email) };
  if (user.gh) return { gh: user.gh, userFolder: null };
  return null;
}

// Returns "mirrored" | "exists". Versions are immutable, so an already-existing
// file (422) simply means this version was mirrored before.
async function mirrorVersion(target, wbName, wbId, v) {
  const p = mirrorPath(wbName, wbId, v, target.userFolder);
  const res = await ghFetch(target.gh.token, "PUT",
    `/repos/${target.gh.login}/${target.gh.repo}/contents/${encodeURI(p)}`, {
      message: `${v.message} — ${wbName || wbId} (${v.takenAt})`,
      content: Buffer.from(JSON.stringify(
        { takenAt: v.takenAt, message: v.message, changeCount: v.changeCount, snapshot: v.snapshot },
        null, 2)).toString("base64"),
    });
  if (res.status === 201 || res.status === 200) return "mirrored";
  if (res.status === 422) return "exists";
  const err = await res.json().catch(() => ({}));
  throw new Error(`GitHub ${res.status}: ${err.message || "mirror failed"}`);
}

async function mirrorAll(user) {
  const target = mirrorTarget(user);
  if (!target) return;
  mirrorStatus[user.id] = { state: "syncing", detail: "backfilling versions…" };
  try {
    const dir = path.join(DATA_DIR, user.id);
    let files = [];
    try { files = fs.readdirSync(dir).filter((f) => f.endsWith(".json")); } catch (e) { /* none */ }
    let mirrored = 0, existed = 0;
    for (const f of files) {
      const wb = loadJson(path.join(dir, f), { versions: [] });
      const wbId = f.replace(/\.json$/, "");
      for (const v of [...wb.versions].reverse()) { // oldest first: sane commit order
        (await mirrorVersion(target, wb.name, wbId, v)) === "mirrored" ? mirrored++ : existed++;
      }
    }
    mirrorStatus[user.id] = { state: "ok", detail: `${mirrored} pushed, ${existed} already in the repo` };
  } catch (e) {
    mirrorStatus[user.id] = { state: "error", detail: e.message };
  }
}

// With a central mirror, backfill every known user shortly after startup.
async function mirrorAllUsersCentral() {
  const users = loadJson(USERS_FILE, {});
  for (const [email, u] of Object.entries(users)) {
    await mirrorAll({ email, ...u });
  }
}

/* ---------- workbook storage ---------- */

function workbookFile(userId, wbId) {
  const safe = String(wbId).replace(/[^a-zA-Z0-9_-]/g, "");
  if (!safe) return null;
  return path.join(DATA_DIR, userId, safe + ".json");
}

function versionMeta(v) {
  return { id: v.id, takenAt: v.takenAt, message: v.message, changeCount: v.changeCount };
}

/* ---------- request handling ---------- */

function send(res, code, obj, origin) {
  res.writeHead(code, {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  });
  res.end(JSON.stringify(obj));
}

async function handle(req, res, body, origin) {
  const parts = req.url.split("?")[0].split("/").filter(Boolean); // ["api", ...]
  if (parts[0] !== "api") return send(res, 404, { error: "not found" }, origin);

  // POST /api/auth/login
  if (parts[1] === "auth" && parts[2] === "login" && req.method === "POST") {
    const email = body && body.email;
    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return send(res, 400, { error: "valid email required" }, origin);
    }
    const user = getOrCreateUser(email);
    return send(res, 200, { email: user.email, token: user.token }, origin);
  }

  // Everything else needs a token.
  const token = (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  const user = userFromToken(token);
  if (!user) return send(res, 401, { error: "sign in required" }, origin);

  // GitHub mirror management
  if (parts[1] === "github") {
    if (CENTRAL && (parts[2] === "connect" || parts[2] === "disconnect")) {
      return send(res, 400, { error: "mirroring is managed centrally on this server" }, origin);
    }
    // POST /api/github/connect {token}
    if (parts[2] === "connect" && req.method === "POST") {
      const ghToken = body && body.token;
      if (!ghToken) return send(res, 400, { error: "token required" }, origin);
      const who = await ghFetch(ghToken, "GET", "/user");
      if (!who.ok) return send(res, 400, { error: "GitHub rejected that token" }, origin);
      const login = (await who.json()).login;
      // Mirror repo is user-choosable; default save-review-versions.
      const repoName =
        String(body.repo || GH_REPO_NAME).trim().replace(/[^a-zA-Z0-9._-]/g, "").slice(0, 100) || GH_REPO_NAME;
      // Use the repo if it already exists (works with a minimal token scoped to
      // just that repo). Only try to create it when it's missing — that needs a
      // token with repo-creation rights.
      const repoCheck = await ghFetch(ghToken, "GET", `/repos/${login}/${repoName}`);
      if (repoCheck.status === 404) {
        const created = await ghFetch(ghToken, "POST", "/user/repos", {
          name: repoName,
          private: true,
          auto_init: true,
          description: "Version history mirrored from Save & Review",
        });
        if (!created.ok && created.status !== 422) {
          return send(res, 400, {
            error: `repo "${repoName}" doesn't exist and this token can't create it. ` +
              `Either create a private repo named "${repoName}" on GitHub yourself ` +
              `(then a token scoped to only that repo, Contents read/write, is enough), ` +
              `or use a token allowed to create repositories.`,
          }, origin);
        }
      } else if (!repoCheck.ok) {
        return send(res, 400, { error: `this token can't access ${login}/${repoName} (GitHub said ${repoCheck.status})` }, origin);
      }
      const gh = { token: ghToken, login, repo: repoName };
      setUserGh(user.email, gh);
      mirrorAll({ ...user, gh }); // fire-and-forget backfill
      return send(res, 200, { login, repo: GH_REPO_NAME }, origin);
    }
    // POST /api/github/disconnect
    if (parts[2] === "disconnect" && req.method === "POST") {
      setUserGh(user.email, null);
      delete mirrorStatus[user.id];
      return send(res, 200, { ok: true }, origin);
    }
    // GET /api/github/status
    if (parts[2] === "status" && req.method === "GET") {
      if (CENTRAL) {
        return send(res, 200, {
          connected: true,
          central: true,
          login: CENTRAL.login,
          repo: CENTRAL.repo,
          mirror: mirrorStatus[user.id] || null,
        }, origin);
      }
      return send(res, 200, {
        connected: !!user.gh,
        login: user.gh?.login,
        repo: user.gh?.repo,
        mirror: mirrorStatus[user.id] || null,
      }, origin);
    }
  }

  // GET /api/workbooks
  if (parts[1] === "workbooks" && parts.length === 2 && req.method === "GET") {
    const dir = path.join(DATA_DIR, user.id);
    let files = [];
    try { files = fs.readdirSync(dir); } catch (e) { /* no workbooks yet */ }
    const list = files
      .filter((f) => f.endsWith(".json"))
      .map((f) => {
        const wb = loadJson(path.join(dir, f), {});
        return { id: f.replace(/\.json$/, ""), name: wb.name || "(unnamed)", versions: (wb.versions || []).length };
      });
    return send(res, 200, list, origin);
  }

  // /api/workbooks/:wb/versions[...]
  if (parts[1] === "workbooks" && parts[3] === "versions") {
    const file = workbookFile(user.id, parts[2]);
    if (!file) return send(res, 400, { error: "bad workbook id" }, origin);
    const wb = loadJson(file, { name: "", versions: [] });

    // GET list (metadata only)
    if (parts.length === 4 && req.method === "GET") {
      return send(res, 200, wb.versions.map(versionMeta), origin);
    }

    // POST new version
    if (parts.length === 4 && req.method === "POST") {
      if (!body || !body.snapshot || !body.takenAt) {
        return send(res, 400, { error: "takenAt and snapshot required" }, origin);
      }
      // Idempotent: same takenAt twice is the same version.
      const existing = wb.versions.find((v) => v.takenAt === body.takenAt);
      if (existing) return send(res, 200, versionMeta(existing), origin);
      const version = {
        id: crypto.randomBytes(8).toString("hex"),
        takenAt: body.takenAt,
        message: String(body.message || "(no message)").slice(0, 200),
        changeCount: Number(body.changeCount) || 0,
        snapshot: body.snapshot,
      };
      wb.versions.unshift(version);
      wb.versions.sort((a, b) => (a.takenAt < b.takenAt ? 1 : -1));
      if (body.name) wb.name = String(body.name).slice(0, 200);
      saveJson(file, wb);
      const target = mirrorTarget(user);
      if (target) {
        // Mirror in the background (central repo or the user's own).
        mirrorVersion(target, wb.name, parts[2], version)
          .then(() => { mirrorStatus[user.id] = { state: "ok", detail: "last version mirrored" }; })
          .catch((e) => { mirrorStatus[user.id] = { state: "error", detail: e.message }; });
      }
      return send(res, 200, versionMeta(version), origin);
    }

    // GET one full version
    if (parts.length === 5 && req.method === "GET") {
      const version = wb.versions.find((v) => v.id === parts[4] || v.takenAt === parts[4]);
      if (!version) return send(res, 404, { error: "version not found" }, origin);
      return send(res, 200, version, origin);
    }
  }

  return send(res, 404, { error: "not found" }, origin);
}

/* ---------- server ---------- */

let options;
try {
  options = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch (e) {
  console.error("No dev certificate found — run `npm run certs` first.");
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    const origin = PANE_ORIGINS.includes(req.headers.origin) ? req.headers.origin : PANE_ORIGINS[0];
    if (req.method === "OPTIONS") return send(res, 204, {}, origin);

    let raw = "";
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY) {
        send(res, 413, { error: "payload too large" }, origin);
        req.destroy();
        return;
      }
      raw += chunk;
    });
    req.on("end", () => {
      if (res.writableEnded) return;
      let body = null;
      if (raw) {
        try { body = JSON.parse(raw); } catch (e) {
          return send(res, 400, { error: "invalid JSON" }, origin);
        }
      }
      handle(req, res, body, origin).catch((e) => {
        console.error(e);
        if (!res.writableEnded) send(res, 500, { error: "server error" }, origin);
      });
    });
  })
  .listen(PORT, () => {
    console.log(`Save & Review version server running at https://localhost:${PORT}`);
    console.log(`Data directory: ${DATA_DIR}`);
    if (CENTRAL) {
      console.log(`Central mirror: ${CENTRAL.login}/${CENTRAL.repo} (backfilling all users)`);
      mirrorAllUsersCentral();
    } else {
      console.log("Central mirror: off (no server/config.json) — per-user GitHub connect available");
    }
  });
