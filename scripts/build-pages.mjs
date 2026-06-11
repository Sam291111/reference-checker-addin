import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const baseUrl = normalizeBaseUrl(process.env.ADDIN_BASE_URL);

if (!baseUrl) {
  throw new Error("Set ADDIN_BASE_URL, for example: ADDIN_BASE_URL=https://USERNAME.github.io/REPO npm run build:pages");
}

await fs.rm(dist, { force: true, recursive: true });
await fs.mkdir(dist, { recursive: true });

await Promise.all([
  copyFile("taskpane.html"),
  copyFile("review.html"),
  copyFile("taskpane.css"),
  copyFile("README.md"),
  copyDirectory("assets"),
  copyDirectory("src")
]);

await writeManifest("manifest.xml");
await writeManifest("manifest-basic.xml");

console.log(`Built GitHub Pages files in ${dist}`);
console.log(`Add-in base URL: ${baseUrl}`);
console.log(`Hosted manifest will be: ${baseUrl}/manifest.xml`);

function normalizeBaseUrl(value) {
  if (!value) {
    return "";
  }
  return value.replace(/\/+$/, "");
}

async function copyFile(relativePath) {
  await fs.copyFile(path.join(root, relativePath), path.join(dist, relativePath));
}

async function copyDirectory(relativePath) {
  await fs.cp(path.join(root, relativePath), path.join(dist, relativePath), {
    recursive: true,
    filter: (source) => !source.endsWith(".DS_Store")
  });
}

async function writeManifest(relativePath) {
  const sourcePath = path.join(root, relativePath);
  const targetPath = path.join(dist, relativePath);
  const source = await fs.readFile(sourcePath, "utf8");
  const output = source.replaceAll("https://localhost:3000", baseUrl);
  await fs.writeFile(targetPath, output);
}
