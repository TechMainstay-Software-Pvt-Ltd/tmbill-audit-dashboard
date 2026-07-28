import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import { extname, resolve, sep } from "node:path";
import { Readable } from "node:stream";
import worker from "../dist/server/index.js";

const port = Number(process.argv[2] || 3001);
const assetsRoot = resolve("dist/client");
const mime = {
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

const assets = {
  async fetch(request) {
    const pathname = decodeURIComponent(new URL(request.url).pathname);
    const target = resolve(assetsRoot, `.${pathname}`);
    if (target !== assetsRoot && !target.startsWith(`${assetsRoot}${sep}`)) {
      return new Response("Forbidden", { status: 403 });
    }
    try {
      const info = await stat(target);
      if (!info.isFile()) return new Response("Not Found", { status: 404 });
      const body = await readFile(target);
      return new Response(body, {
        headers: {
          "content-type": mime[extname(target).toLowerCase()] || "application/octet-stream",
          "cache-control": pathname.startsWith("/assets/") ? "public, max-age=31536000, immutable" : "no-cache",
        },
      });
    } catch {
      return new Response("Not Found", { status: 404 });
    }
  },
};

createServer(async (req, res) => {
  try {
    const url = `http://${req.headers.host || `localhost:${port}`}${req.url || "/"}`;
    const request = new Request(url, { method: req.method, headers: req.headers });
    const pathname = new URL(url).pathname;
    if (pathname.startsWith("/assets/") || pathname.startsWith("/data/") || pathname.startsWith("/downloads/") || /\.[a-z0-9]{2,5}$/i.test(pathname)) {
      const assetResponse = await assets.fetch(request);
      if (assetResponse.status !== 404) {
        res.writeHead(assetResponse.status, Object.fromEntries(assetResponse.headers));
        if (assetResponse.body) Readable.fromWeb(assetResponse.body).pipe(res);
        else res.end();
        return;
      }
    }
    const context = { waitUntil() {}, passThroughOnException() {} };
    const response = await worker.fetch(request, { ASSETS: assets }, context);
    res.writeHead(response.status, Object.fromEntries(response.headers));
    if (response.body) Readable.fromWeb(response.body).pipe(res);
    else res.end();
  } catch (error) {
    res.writeHead(500, { "content-type": "text/plain; charset=utf-8" });
    res.end(error instanceof Error ? error.stack : String(error));
  }
}).listen(port, "0.0.0.0", () => {
  console.log(`TMBill production dashboard: http://localhost:${port}/`);
});
