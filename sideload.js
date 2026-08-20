// Copies manifest.xml into Excel-for-Mac's sideload folder ("wef").
// After running this, the add-in appears in Excel under:
//   Insert tab -> Add-ins (or "My Add-ins" dropdown) -> Developer Add-ins
const fs = require("fs");
const path = require("path");
const os = require("os");

const wef = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Excel/Data/Documents/wef"
);

fs.mkdirSync(wef, { recursive: true });
fs.copyFileSync(path.join(__dirname, "manifest.xml"), path.join(wef, "save-review-manifest.xml"));
console.log("Manifest copied to:\n  " + wef);
console.log("\nNow fully quit Excel (Cmd+Q) and reopen it, then look for the");
console.log('"Delta" button on the Home tab (or Insert -> My Add-ins).');
