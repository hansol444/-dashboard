import { writeFile, readFile, mkdir } from "fs/promises";
import path from "path";
import os from "os";

/**
 * Resolve a file path: if it's a URL (Blob), download to /tmp and return local path.
 * If it's already a local path, return as-is.
 */
export async function resolveFile(filePathOrUrl: string): Promise<string> {
  if (filePathOrUrl.startsWith("http://") || filePathOrUrl.startsWith("https://")) {
    // Download from Blob URL to /tmp
    const res = await fetch(filePathOrUrl);
    if (!res.ok) throw new Error(`Failed to download: ${res.status}`);
    const buffer = Buffer.from(await res.arrayBuffer());

    const urlPath = new URL(filePathOrUrl).pathname;
    const filename = path.basename(urlPath);
    const tmpDir = os.tmpdir();
    const localPath = path.join(tmpDir, `dl_${Date.now()}_${filename}`);

    await mkdir(tmpDir, { recursive: true }).catch(() => {});
    await writeFile(localPath, buffer);
    return localPath;
  }
  return filePathOrUrl;
}

/**
 * Save data to a temp file and return the path.
 */
export async function saveTempFile(filename: string, data: Buffer | string): Promise<string> {
  const tmpDir = os.tmpdir();
  const tmpPath = path.join(tmpDir, `${Date.now()}_${filename}`);
  await mkdir(tmpDir, { recursive: true }).catch(() => {});
  await writeFile(tmpPath, data);
  return tmpPath;
}

/**
 * Read a temp file.
 */
export async function readTempFile(filePath: string): Promise<Buffer> {
  return readFile(filePath);
}
