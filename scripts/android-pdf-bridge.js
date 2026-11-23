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

  const path = filename;
  let existing = false;
  try {
    await Filesystem.stat({ path, directory: Directory.Documents });
    existing = true;
  } catch (error) {
    // A missing file is expected on the first export. Other failures will be
    // surfaced by the subsequent write rather than silently treated as saved.
    if (error?.message && !/not found|does not exist|no such file/i.test(error.message)) {
      throw error;
    }
  }

  if (existing) {
    const replace = window.confirm(
      `A PDF named "${filename}" already exists in Documents. Replace it with the newly generated report?`
    );
    if (!replace) return { cancelled: true, filename };
    await Filesystem.deleteFile({ path, directory: Directory.Documents });
  }

  const data = toBase64(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const saved = await Filesystem.writeFile({
    path,
    data,
    directory: Directory.Documents,
    recursive: true
  });
  if (!saved?.uri) throw new Error("Android did not return the saved PDF file location.");

  // Always explicitly request Android's app chooser/default PDF handler.
  // If opening fails, the file remains saved and the caller receives the error.
  await FileOpener.open({
    filePath: saved.uri,
    contentType: "application/pdf",
    openWithDefault: false
  });

  return { uri: saved.uri, filename, replaced: existing, opened: true };
}

window.SDDAAndroidPdf = { isNative, saveAndOpen };
