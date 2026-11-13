// Entrypoint. Configuration (environment):
//   PORT                          listen port (default 8790)
//   HOST                          bind address (default 0.0.0.0; use 127.0.0.1 locally)
//   OFFICE_PLATFORM_URL           platform origin that mints tokens and stores revisions (required)
//   OFFICE_ENGINE_JWKS_URL        override JWKS URL (default <platform>/api/office/jwks)
//   OFFICE_ENGINE_JWT_ISSUER      override issuer (default platform origin)
//   OFFICE_ENGINE_JWT_AUDIENCE    default seegerweiss-office-engine
//   OFFICE_ENGINE_TENANT          default seegerweiss
//   OFFICE_ENGINE_ALLOWED_ORIGINS comma-separated browser origins for CORS (dev only)
//   OFFICE_ENGINE_SCRATCH         scratch directory (default ./.local/scratch)
//   XLSX_SIDECAR_PATH             path to the Rust sidecar binary
//   OFFICE_ENGINE_MAX_SESSIONS    open workbooks per task (default 24)
//   OFFICE_ENGINE_MAX_PER_USER    open workbooks per user (default 4)
import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import { createRequire } from "node:module";
import { resolve } from "node:path";

import { createEngineServer } from "./application.mjs";
import { loadAuthConfig, TokenAuth } from "./auth.mjs";
import { EnginePool } from "./engine-pool.mjs";

const require = createRequire(import.meta.url);
const root = resolve(import.meta.dirname, "..");
const { version } = require("../package.json");

const port = Number(process.env.PORT || 8790);
if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error("Set a valid PORT.");
const host = process.env.HOST || "0.0.0.0";

const binaryPath =
  process.env.XLSX_SIDECAR_PATH ||
  resolve(root, "native/xlsx-engine/target/release", process.platform === "win32" ? "xlsx-sidecar.exe" : "xlsx-sidecar");
if (!existsSync(binaryPath)) throw new Error(`xlsx-sidecar binary not found at ${binaryPath}. Run npm run build:native or set XLSX_SIDECAR_PATH.`);
if (!existsSync(resolve(root, ".build/native-worker.cjs"))) throw new Error("Engine bundle missing. Run npm run build.");

const scratchRoot = resolve(process.env.OFFICE_ENGINE_SCRATCH || resolve(root, ".local/scratch"));
await mkdir(scratchRoot, { recursive: true });

const config = loadAuthConfig();
const auth = new TokenAuth(config);
const pool = new EnginePool({
  root,
  scratchRoot,
  binaryPath,
  platformUrl: config.platformUrl,
  maxSessions: Number(process.env.OFFICE_ENGINE_MAX_SESSIONS || 24),
  maxPerUser: Number(process.env.OFFICE_ENGINE_MAX_PER_USER || 4),
});
const engine = require("../.build/engine.cjs");
const server = createEngineServer({ auth, pool, engine, version });

server.on("error", (error) => {
  console.error("Office engine could not start:", error.code || error.message);
  process.exitCode = 1;
});
server.listen(port, host, () => {
  console.log(`Office engine ${version} listening on http://${host}:${port}`);
  console.log(`Platform: ${config.platformUrl}  JWKS: ${config.jwksUrl}  sidecar: ${binaryPath}`);
  if (config.allowedOrigins.length) console.log(`CORS origins: ${config.allowedOrigins.join(", ")}`);
});

let stopping = false;
for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    if (stopping) return;
    stopping = true;
    console.log("Office engine shutting down...");
    void pool.dispose().finally(() => server.close(() => process.exit(0)));
    setTimeout(() => process.exit(0), 10_000).unref();
  });
}
