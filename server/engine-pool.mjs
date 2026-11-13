// One Node worker thread per open workbook, each driving its own Rust
// xlsx-sidecar. Ported from the Office server's engine-pool with the storage
// swapped: the workbook is fetched from a presigned S3 grant at open, and a
// save writes the new revision back through the platform's document API with
// the caller's engine token. This service holds no bucket credentials and no
// database.
import { Worker } from "node:worker_threads";
import { mkdir, writeFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { randomUUID, createHash } from "node:crypto";

import { reject } from "./auth.mjs";
import { operationKind, safeValue } from "./boundary.mjs";

const hash = (v) => createHash("sha256").update(JSON.stringify(v)).digest("hex");
const MAX_WORKBOOK_BYTES = 30 * 1024 * 1024;
const OP_TIMEOUT_MS = 45_000;
const IDLE_SESSION_MS = 30 * 60_000;

/** Per document kind: which retained engine hosts it and how it saves. */
export const KINDS = Object.freeze({
  xlsx: Object.freeze({
    app: "sheets",
    worker: ".build/native-worker.cjs",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    /** Sheets returns {canceled:false} on a completed save. */
    saved: (result) => result?.canceled === false,
    /** The workbook file name must keep its extension for the sidecar. */
    ext: ".xlsx",
  }),
  pptx: Object.freeze({
    app: "slides",
    worker: ".build/slides-worker.cjs",
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    /** Slides returns {ok:true} on a completed save. */
    saved: (result) => result?.ok === true,
    ext: ".pptx",
  }),
});
export const isKind = (v) => typeof v === "string" && Object.hasOwn(KINDS, v);

export class EnginePool {
  /**
   * @param {{root:string, scratchRoot:string, binaryPath:string, platformUrl:string, maxSessions:number, maxPerUser:number}} options
   */
  constructor(options) {
    Object.assign(this, options);
    this.sessions = new Map();
    this.reaper = setInterval(() => void this.reapIdle(), 60_000);
    this.reaper.unref();
  }

  async fetchGrant(source, expectedHash) {
    const url = new URL(source);
    if (url.protocol !== "https:" && !(url.protocol === "http:" && ["127.0.0.1", "localhost"].includes(url.hostname))) {
      reject(422, "The document grant must be an HTTPS URL.");
    }
    const res = await fetch(url, { redirect: "error", signal: AbortSignal.timeout(60_000) });
    if (!res.ok) reject(502, `The document could not be fetched from storage (HTTP ${res.status}).`);
    const length = Number(res.headers.get("content-length") ?? "0");
    if (length > MAX_WORKBOOK_BYTES) reject(413, "The document exceeds the size limit.");
    const bytes = new Uint8Array(await res.arrayBuffer());
    if (bytes.byteLength > MAX_WORKBOOK_BYTES) reject(413, "The document exceeds the size limit.");
    if (expectedHash && createHash("sha256").update(bytes).digest("hex") !== expectedHash) {
      reject(409, "The fetched document does not match the expected revision. Reload the document.");
    }
    return bytes;
  }

  /**
   * Open a document session for the actor.
   * @param actor verified token actor (doc-scoped)
   * @param input {docId, kind, name, version, hash, source}
   */
  async open(actor, input) {
    if (input.docId !== actor.doc) reject(403, "The token is not scoped to this document.");
    const kind = isKind(input.kind) ? KINDS[input.kind] : KINDS.xlsx;
    // One live session per user and document: reopening (reload, second tab)
    // replaces the previous session instead of leaking it until the reaper.
    for (const existing of [...this.sessions.values()]) {
      if (existing.subject === actor.subject && existing.docId === input.docId) await this.destroy(existing);
    }
    const mine = [...this.sessions.values()].filter((x) => x.subject === actor.subject);
    if (this.sessions.size >= this.maxSessions) reject(503, "The spreadsheet engine is at capacity. Try again shortly.");
    if (mine.length >= this.maxPerUser) reject(429, "Close an existing workbook tab before opening another.");
    const fallback = `Untitled${kind.ext}`;
    const name = String(input.name || fallback).replace(/[^\w.\- ()]/g, "_").slice(0, 160) || fallback;
    const bytes = await this.fetchGrant(input.source, input.hash);

    const sessionId = randomUUID();
    const scratch = join(this.scratchRoot, sessionId);
    const documentPath = join(scratch, name.toLowerCase().endsWith(kind.ext) ? name : `${name}${kind.ext}`);
    await mkdir(scratch, { recursive: true, mode: 0o700 });
    await writeFile(documentPath, bytes, { flag: "wx", mode: 0o600 });

    const env = {};
    for (const k of ["PATH", "Path", "SystemRoot", "WINDIR", "TEMP", "TMP", "USERPROFILE", "HOME", "LOCALAPPDATA"]) {
      if (process.env[k]) env[k] = process.env[k];
    }
    env.OFFICE_ENGINE_SCRATCH = scratch;
    const worker = new Worker(join(this.root, kind.worker), {
      workerData: {
        root: this.root,
        scratch,
        documentPath,
        documentToken: "office:" + input.docId,
        app: kind.app,
        binaryPath: this.binaryPath,
      },
      env,
      resourceLimits: { maxOldGenerationSizeMb: 512, maxYoungGenerationSizeMb: 48 },
      stdout: true,
      stderr: true,
    });
    const record = {
      sessionId,
      subject: actor.subject,
      docId: input.docId,
      kind,
      version: Number(input.version),
      token: actor.token,
      sequence: 0,
      worker,
      scratch,
      pending: new Map(),
      operations: new Map(),
      chain: Promise.resolve(),
      dead: false,
      openedAt: Date.now(),
      touchedAt: Date.now(),
    };
    this.sessions.set(sessionId, record);
    worker.stdout.on("data", () => {});
    worker.stderr.on("data", (d) => console.error(`[worker ${sessionId.slice(0, 8)}]`, String(d).slice(0, 500)));
    const fail = () => {
      record.dead = true;
      for (const p of record.pending.values()) {
        clearTimeout(p.timer);
        p.reject(new Error("The document engine stopped. Reload the saved document; unsaved changes may require recovery."));
      }
      record.pending.clear();
    };
    worker.on("error", (e) => {
      console.error(`[worker ${sessionId.slice(0, 8)}] error`, e?.message);
      fail();
    });
    worker.on("exit", fail);
    worker.on("message", (r) => {
      const p = record.pending.get(r.id);
      if (!p) return;
      record.pending.delete(r.id);
      clearTimeout(p.timer);
      r.ok ? p.resolve(r) : p.reject(new Error(r.error || "Document operation failed."));
    });
    try {
      const response = await this.call(record, "host:open", []);
      return { sessionId, version: record.version, sequence: 0, ...response };
    } catch (error) {
      await this.close(actor, sessionId);
      throw error;
    }
  }

  get(actor, id) {
    const r = this.sessions.get(id);
    if (!r || r.subject !== actor.subject || r.docId !== actor.doc) reject(404, "Document session not found.");
    if (r.dead) reject(409, "Document session stopped. Reload your saved document.");
    // Refresh the token used for platform callbacks with the caller's current one.
    r.token = actor.token;
    return r;
  }

  call(record, channel, args) {
    if (record.dead) return Promise.reject(new Error("Document engine is no longer available."));
    const id = randomUUID();
    return new Promise((resolve, rejectCall) => {
      const timer = setTimeout(() => {
        record.pending.delete(id);
        record.dead = true;
        void record.worker.terminate();
        rejectCall(new Error("The document operation timed out. Reload a saved revision before continuing."));
      }, OP_TIMEOUT_MS);
      record.pending.set(id, { resolve, reject: rejectCall, timer });
      record.worker.postMessage({ id, channel, args });
    });
  }

  /** Save the current document bytes as a new revision through the platform API. */
  async persist(record, bytes, operationId) {
    const url = `${this.platformUrl}/api/public/office/engine/docs/${record.docId}/content`;
    // A transport error can arrive after the platform committed the revision.
    // Retry only that ambiguity, once, with the exact same operation and body.
    // A fresh operation ID would turn a successful save into a version conflict.
    const platformRequest = async init => {
      for (let attempt = 0; attempt < 2; attempt++) {
        let response;
        try {
          response = await fetch(url, { ...init, redirect: 'error', signal: AbortSignal.timeout(60_000) });
        } catch (error) {
          if (attempt === 0 && (error instanceof TypeError || error?.name === 'TimeoutError')) continue;
          throw error;
        }
        if (attempt === 0 && [502, 503, 504].includes(response.status)) {
          await response.body?.cancel().catch(() => undefined);
          continue;
        }
        let result;
        try { result = await response.json(); }
        catch (error) {
          if (response.ok && attempt === 0 && (error instanceof TypeError || error?.name === 'TimeoutError')) continue;
          reject(response.ok ? 502 : response.status, 'The platform returned an invalid save response.');
        }
        if (!response.ok) reject(response.status >= 400 && response.status < 500 ? response.status : 502,
          result?.error || 'The platform could not store the revision.');
        return result;
      }
    };
    const confirmedRevision = result => {
      if (!result || !Number.isSafeInteger(result.version) || result.version !== record.version + 1)
        reject(502, 'The platform did not confirm the expected saved revision. Reopen before continuing.');
      return result;
    };
    if (bytes.byteLength > 3 * 1024 * 1024) {
      const input = { expectedVersion: record.version, operationId, size: bytes.byteLength,
        sha256: createHash('sha256').update(bytes).digest('hex') };
      const control = method => platformRequest({method,
        headers:{authorization:`Bearer ${record.token}`,'content-type':'application/json'}, body:JSON.stringify(input)});
      const grant = await control('POST');
      // Storage receives only native bytes and the checksum, never the engine JWT.
      const uploaded = await fetch(grant.url,{method:'PUT',headers:{'x-amz-checksum-sha256':grant.checksum},
        body:bytes,redirect:'error',signal:AbortSignal.timeout(120_000)});
      if (!uploaded.ok) reject(502,'The document upload failed. Retry Save.');
      return confirmedRevision(await control('PUT'));
    }
    return confirmedRevision(await platformRequest({
      method: "PUT",
      headers: {
        authorization: `Bearer ${record.token}`,
        "content-type": record.kind.mime,
        "if-match": String(record.version),
        "idempotency-key": operationId,
      },
      body: bytes,
    }));
  }

  async rpc(actor, sessionId, { channel, args, sequence, operationId }) {
    const r = this.get(actor, sessionId);
    const kind = operationKind(r.kind.app, channel);
    safeValue(args);
    if (!Array.isArray(args) || args.length > 8) reject(422, "Invalid operation arguments.");
    if (channel === "workbook:save" && args[0]?.mode === "save-as") {
      reject(501, "Save As is not available here. Save a revision, then use Download or Save a copy.");
    }
    if (kind !== "read" && !actor.write) reject(403, "Editing is not authorized.");
    if (!/^[a-f0-9-]{36}$/.test(operationId ?? "")) reject(422, "A unique operation identifier is required.");
    const fingerprint = hash({ channel, args, sequence });
    const prior = r.operations.get(operationId);
    if (prior) {
      if (prior.fingerprint !== fingerprint) reject(409, "Operation ID was reused with different content.");
      return prior.promise;
    }
    const run = r.chain.then(async () => {
      if (sequence !== r.sequence) reject(409, "The editing session changed. Reload before repeating this operation.");
      const response = await this.call(r, channel, args);
      let saved = null;
      if (kind === "save" && r.kind.saved(response.result)) {
        const raw = await this.call(r, "host:bytes", []);
        saved = await this.persist(r, Buffer.from(raw.result), operationId);
        r.version = saved.version;
      }
      if (kind !== "read") r.sequence++;
      r.touchedAt = Date.now();
      return { ...response, operationKind: kind, sequence: r.sequence, version: r.version, document: saved };
    });
    r.chain = run.catch(() => {});
    r.operations.set(operationId, { fingerprint, promise: run });
    if (r.operations.size > 120) r.operations.delete(r.operations.keys().next().value);
    return run;
  }

  async close(actor, id) {
    const record = this.sessions.get(id);
    if (!record || record.subject !== actor.subject) return;
    await this.destroy(record);
  }

  async destroy(record) {
    this.sessions.delete(record.sessionId);
    // Give the worker a short grace period to stop its sidecar cleanly, then
    // terminate regardless: a stuck close must never hold the slot.
    await Promise.race([
      this.call(record, "host:close", []).catch(() => undefined),
      new Promise((resolve) => setTimeout(resolve, 3_000)),
    ]);
    await record.worker.terminate().catch(() => {});
    await rm(record.scratch, { recursive: true, force: true }).catch(() => {});
  }

  async reapIdle() {
    const now = Date.now();
    for (const r of [...this.sessions.values()]) {
      if (now - r.touchedAt > IDLE_SESSION_MS) {
        console.log(`[pool] closing idle session ${r.sessionId.slice(0, 8)} for doc ${r.docId}`);
        await this.destroy(r);
      }
    }
  }

  stats() {
    return { sessions: this.sessions.size, users: new Set([...this.sessions.values()].map((r) => r.subject)).size };
  }

  async dispose() {
    clearInterval(this.reaper);
    for (const r of [...this.sessions.values()]) await this.destroy(r);
  }
}

export function wireEncode(value) {
  if (value instanceof ArrayBuffer) return { $officeBytes: Buffer.from(value).toString("base64"), kind: "arraybuffer" };
  if (value instanceof Uint8Array) return { $officeBytes: Buffer.from(value).toString("base64"), kind: "uint8" };
  if (Array.isArray(value)) return value.map(wireEncode);
  if (value && typeof value === "object") {
    const out = {};
    for (const [k, v] of Object.entries(value)) out[k] = wireEncode(v);
    return out;
  }
  return value;
}

export function wireDecode(value) {
  if (Array.isArray(value)) return value.map(wireDecode);
  if (value && typeof value === "object") {
    if (typeof value.$officeBytes === "string") {
      if (value.$officeBytes.length > 24_000_000 || !/^[A-Za-z0-9+/]*={0,2}$/.test(value.$officeBytes)) {
        reject(422, "Invalid byte payload.");
      }
      const b = Buffer.from(value.$officeBytes, "base64");
      return value.kind === "arraybuffer" ? b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) : new Uint8Array(b);
    }
    const out = {};
    for (const [k, v] of Object.entries(value)) {
      if (["__proto__", "constructor", "prototype"].includes(k)) reject(422, "Unsafe property.");
      out[k] = wireDecode(v);
    }
    return out;
  }
  return value;
}
