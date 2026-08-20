// Tiny HTTPS static file server for local add-in development.
// Uses the trusted localhost certificate created by:  npm run certs
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3002;
const ROOT = __dirname;
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const MIME = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".xml": "text/xml",
  ".json": "application/json",
};

let options;
try {
  options = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch (e) {
  console.error(
    "\nNo dev certificate found. Run this first (it will ask for your Mac password\n" +
      "once, to trust the certificate):\n\n    npm run certs\n"
  );
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(req.url.split("?")[0]);
    let filePath = path.normalize(path.join(ROOT, urlPath === "/" ? "/src/taskpane.html" : urlPath));
    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      return res.end("Forbidden");
    }
    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end("Not found: " + urlPath);
      }
      res.writeHead(200, { "Content-Type": MIME[path.extname(filePath)] || "application/octet-stream" });
      res.end(data);
    });
  })
  .listen(PORT, () => {
    console.log(`Save & Review dev server running at https://localhost:${PORT}`);
    console.log("Leave this running while you use the add-in in Excel.");
  });
