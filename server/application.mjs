// HTTP surface of the Office engine service.
//
//   GET  /health                         liveness + pool stats (no auth)
//   POST /engine/open                    {docId, name, version, hash, source}  -> session
//   POST /engine/:session/rpc            {channel, args, sequence, operationId} -> result
//   POST /engine/:session/close          -> {ok}
//   POST /convert/csv                    body: CSV text, X-Office-Filename -> XLSX bytes
//
// Every non-health request needs a platform engine token. Responses never
// include stack traces; unknown errors are logged and reported generically.
import http from "node:http";

import { HttpError, reject } from "./auth.mjs";
import { isKind, wireDecode, wireEncode } from "./engine-pool.mjs";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function body(req, limit) {
  let total = 0;
  const parts = [];
  for await (const part of req) {
    total += part.length;
    if (total > limit) reject(413, "Request exceeds its size limit.");
    parts.push(part);
  }
  return Buffer.concat(parts);
}
async function jsonBody(req, limit = 16_384) {
  const raw = await body(req, limit);
  try {
    return JSON.parse(raw.toString("utf8"));
  } catch {
    reject(400, "Invalid JSON request.");
  }
}
function json(res, status, value) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  res.end(JSON.stringify(value));
}

export function createEngineServer({ auth, pool, engine, version }) {
  const server = http.createServer(async (req, res) => {
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("Referrer-Policy", "no-referrer");
    res.setHeader("Cache-Control", "no-store");
    try {
      const url = new URL(req.url, "http://engine.local");
      const path = url.pathname;
      const method = req.method;

      if (!auth.cors(req, res)) reject(403, "Origin is not allowed.");
      if (method === "OPTIONS") {
        res.writeHead(204);
        return res.end();
      }
      if (path === "/health" && method === "GET") {
        return json(res, 200, { ok: true, version, ...pool.stats() });
      }

      const actor = await auth.actor(req);
      if (path === "/engine/open" && method === "POST") {
        const r = await jsonBody(req, 8192);
        if (typeof r?.docId !== "string" || typeof r?.source !== "string") reject(422, "docId and source are required.");
        if (r.kind !== undefined && !isKind(r.kind)) reject(422, "Unsupported document kind.");
        return json(res, 200, wireEncode(await pool.open(actor, r)));
      }
      const engineMatch = path.match(/^\/engine\/([a-f0-9-]{36})\/(rpc|close)$/);
      if (engineMatch && method === "POST") {
        if (engineMatch[2] === "close") {
          await pool.close(actor, engineMatch[1]);
          return json(res, 200, { ok: true });
        }
        const r = wireDecode(await jsonBody(req, 26_000_000));
        return json(res, 200, wireEncode(await pool.rpc(actor, engineMatch[1], r)));
      }
      if (path === "/convert/csv" && method === "POST") {
        if (!actor.write) reject(403, "Editing is not authorized.");
        let name = "Untitled.csv";
        try {
          name = decodeURIComponent(String(req.headers["x-office-filename"] || "")) || name;
        } catch {
          /* keep default */
        }
        const bytes = await body(req, 16 * 1024 * 1024);
        const out = await engine.importCsv(new Uint8Array(bytes), name);
        res.writeHead(200, {
          "Content-Type": XLSX_MIME,
          "Content-Length": out.bytes.byteLength,
          "X-Office-Filename": encodeURIComponent(out.name),
        });
        return res.end(Buffer.from(out.bytes));
      }
      reject(404, "Office engine operation not found.");
    } catch (error) {
      if (res.headersSent) {
        res.end();
        return;
      }
      const status = Number.isInteger(error?.status) ? error.status : 500;
      if (status === 500) console.error("[engine] request failed:", error);
      json(res, status, {
        error: status === 500 || !(error instanceof HttpError || error?.status) ? "The document engine encountered an internal error." : error.message,
      });
    }
  });
  server.requestTimeout = 120_000;
  server.headersTimeout = 15_000;
  server.maxHeadersCount = 80;
  return server;
}
