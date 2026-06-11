import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = path.resolve(root, process.env.MANIFEST_PATH || "manifest.xml");
const targetDir = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Excel/Data/Documents/wef"
);
const targetPath = path.join(targetDir, "reference-checker-manifest.xml");

await fs.mkdir(targetDir, { recursive: true });
await fs.copyFile(manifestPath, targetPath);

console.log(`Installed manifest: ${targetPath}`);
console.log(`Source manifest: ${manifestPath}`);
