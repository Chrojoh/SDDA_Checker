import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";

const root = process.cwd();
const webDir = join(root, "www");
await mkdir(webDir, { recursive: true });

const files = [
  "mobile-v2.html",
  "app.js",
  "styles.css",
  "games.css",
  "mobile-dashboard.css",
  "manifest.webmanifest",
  "sw.js"
];

for (const file of files) {
  await cp(join(root, file), join(webDir, file));
}

// Launch directly into the mobile interface inside the Android app.
const mobileHtml = await readFile(join(root, "mobile-v2.html"), "utf8");
await writeFile(join(webDir, "index.html"), mobileHtml, "utf8");

for (const directory of ["vendor", "icons"]) {
  await cp(join(root, directory), join(webDir, directory), { recursive: true });
}

console.log("Prepared the mobile SDDA Ledger web assets for Capacitor.");
