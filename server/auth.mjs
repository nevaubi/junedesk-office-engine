// Request authentication for the Office engine service.
//
// Every request carries `Authorization: Bearer <token>` where the token was
// minted by the Seeger Weiss platform (RS256, 15 minute lifetime) for one user
// and one document. Verification uses the platform's JWKS (createRemoteJWKSet
// caches keys and refetches on unknown kid). There are no cookies and no
// sessions of its own: the actor is derived from the token on each call, and
// the engine's workbook sessions are keyed by that actor.
import { createRemoteJWKSet, jwtVerify } from "jose";

export class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}
export const reject = (status, message) => {
  throw new HttpError(status, message);
};

export function loadAuthConfig(env = process.env) {
  const platformUrl = env.OFFICE_PLATFORM_URL;
  if (!platformUrl) throw new Error("OFFICE_PLATFORM_URL is required (the platform origin that mints engine tokens).");
  const origin = new URL(platformUrl);
  const production = env.NODE_ENV === "production";
  if (production && origin.protocol !== "https:") throw new Error("OFFICE_PLATFORM_URL must be https in production.");
  return {
    platformUrl: origin.origin,
    issuer: env.OFFICE_ENGINE_JWT_ISSUER || origin.origin,
    jwksUrl: env.OFFICE_ENGINE_JWKS_URL || `${origin.origin}/api/public/office/jwks`,
    audience: env.OFFICE_ENGINE_JWT_AUDIENCE || "local-office-engine",
    tenant: env.OFFICE_ENGINE_TENANT || "local",
    requiredScope: env.OFFICE_ENGINE_REQUIRED_SCOPE || "office:write",
    /** Browser origins allowed to call the engine (CORS). Same-origin deployments leave this empty. */
    allowedOrigins: (env.OFFICE_ENGINE_ALLOWED_ORIGINS || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean),
  };
}

export class TokenAuth {
  constructor(config) {
    this.config = config;
    this.keys = createRemoteJWKSet(new URL(config.jwksUrl), { timeoutDuration: 5000, cooldownDuration: 30_000 });
  }

  /** Verify the bearer token; returns the actor or throws HttpError(401/403). */
  async actor(req) {
    const header = String(req.headers.authorization ?? "");
    const token = header.startsWith("Bearer ") ? header.slice(7) : "";
    if (!token || token.length > 16_000) reject(401, "Sign in to the workspace.");
    let payload;
    try {
      ({ payload } = await jwtVerify(token, this.keys, {
        issuer: this.config.issuer,
        audience: this.config.audience,
        algorithms: ["RS256", "EdDSA"],
        clockTolerance: 5,
      }));
    } catch (error) {
      // The jose error code (no token contents) is enough to diagnose issuer,
      // audience, key or clock problems without leaking anything sensitive.
      const code = error?.code || error?.name || "verify_failed";
      console.warn(`[auth] token rejected: ${code}${error?.claim ? ` (${error.claim})` : ""}`);
      reject(401, `The workspace token could not be verified (${code}). Reload the document.`);
    }
    const scopes = new Set(String(payload.scope ?? "").split(" "));
    if (!payload.sub || !payload.exp || !scopes.has(this.config.requiredScope) || payload.token_use !== "access") {
      reject(403, "The workspace token does not authorize this engine.");
    }
    if (payload.tenant !== this.config.tenant) reject(403, "The workspace tenant is not authorized.");
    if (typeof payload.doc !== "string" || !/^(?:[0-9A-HJKMNP-TV-Z]{26}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/.test(payload.doc)) {
      reject(403, "The workspace token is not scoped to a document.");
    }
    return {
      tenant: this.config.tenant,
      subject: payload.sub,
      name: typeof payload.name === "string" ? payload.name : "Workspace user",
      write: scopes.has("office:write"),
      doc: payload.doc,
      expires: payload.exp * 1000,
      token,
    };
  }

  /** CORS for a browser on another origin (local development); same-origin deployments never send Origin. */
  cors(req, res) {
    const origin = req.headers.origin;
    if (!origin) return true;
    if (!this.config.allowedOrigins.includes(origin)) return false;
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
    res.setHeader("Access-Control-Allow-Headers", "authorization, content-type, x-office-filename");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Max-Age", "600");
    return true;
  }
}
