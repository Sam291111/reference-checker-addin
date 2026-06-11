import http from "node:http";
import https from "node:https";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT || 3000);
const protocol = process.env.USE_HTTP === "1" ? "http" : "https";

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml; charset=utf-8",
  ".xml": "application/xml; charset=utf-8"
};

function safePath(urlPath) {
  const requested = decodeURIComponent(urlPath.split("?")[0]);
  const filePath = requested === "/" ? "/taskpane.html" : requested;
  const resolved = path.resolve(root, `.${filePath}`);
  if (!resolved.startsWith(root)) {
    return null;
  }
  return resolved;
}

async function handleRequest(request, response) {
  const resolved = safePath(request.url || "/");
  if (!resolved) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  try {
    const data = await fs.readFile(resolved);
    response.writeHead(200, {
      "Content-Type": contentTypes[path.extname(resolved)] || "application/octet-stream",
      "Cache-Control": "no-store"
    });
    response.end(data);
  } catch (error) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
  }
}

async function createServer() {
  if (protocol === "http") {
    return http.createServer(handleRequest);
  }

  const officeCertDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const officeCert = {
    key: path.join(officeCertDir, "localhost.key"),
    cert: path.join(officeCertDir, "localhost.crt")
  };
  const localCert = {
    key: path.join(root, "certs/localhost.key"),
    cert: path.join(root, "certs/localhost.crt")
  };
  const certPaths = await fileExists(officeCert.cert) && await fileExists(officeCert.key)
    ? officeCert
    : localCert;
  const [key, cert] = await Promise.all([
    fs.readFile(certPaths.key),
    fs.readFile(certPaths.cert)
  ]);
  return https.createServer({ key, cert }, handleRequest);
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

const server = await createServer();

server.listen(port, "127.0.0.1", () => {
  console.log(`Reference Checker add-in is running at ${protocol}://localhost:${port}/taskpane.html`);
  console.log(`Manifest: ${path.join(root, "manifest.xml")}`);
});
