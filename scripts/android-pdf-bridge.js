import { Filesystem, Directory } from "@capacitor/filesystem";
import { FileOpener } from "@capacitor-community/file-opener";

function isNative() {
  return !!(window.Capacitor?.isNativePlatform?.() && window.Capacitor?.getPlatform?.() === "android");
}
function toBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}
async function saveAndOpen(bytes, filename) {
  if (!isNative()) throw new Error("Native Android file access is unavailable.");
  const data = toBase64(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const result = await Filesystem.writeFile({
    path: filename,
    data,
    directory: Directory.Documents,
    recursive: true
  });
  if (!result?.uri) throw new Error("Android did not return a saved PDF file location.");
  await FileOpener.open({ filePath: result.uri, contentType: "application/pdf" });
  return { uri: result.uri, filename };
}
window.SDDAAndroidPdf = { isNative, saveAndOpen };
