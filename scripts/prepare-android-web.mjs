import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

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


const { build } = await import("esbuild");
await build({
  entryPoints: [join(root, "scripts/android-pdf-bridge.js")],
  bundle: true,
  platform: "browser",
  format: "iife",
  target: ["es2020"],
  outfile: join(webDir, "android-pdf-bridge.js")
});
const bundledHtml = mobileHtml.replace(
  /<script src="app\\.js(?:\\?[^"]*)?"><\\/script>/,
  match => '<script src="android-pdf-bridge.js"></script>' + match
);
if (bundledHtml === mobileHtml) {
  throw new Error("Could not inject the Android PDF bridge before app.js.");
}
await writeFile(join(webDir, "index.html"), bundledHtml, "utf8");
console.log("Prepared the mobile SDDA Ledger web assets and native PDF bridge for Capacitor.");

