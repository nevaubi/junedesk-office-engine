# Office engine service

Hosts the retained document engines for the Seeger Weiss platform: one Node
worker thread per open document, running the Sheets workbook engine (driving
the Rust `xlsx-sidecar`) or the Slides deck engine (pptx-engine + pptx-render
with harfbuzz text shaping). The platform on Lambda cannot keep stateful worker
sessions, so this runs as an always-on service (ECS Fargate in AWS, a plain
Node process locally).

What it does and does not do:

- Verifies platform-minted RS256 tokens (15 min, scoped to one user and one
  document) against the platform JWKS at `/api/public/office/jwks`.
- Loads a document from a presigned S3 URL the platform issues at open time.
  `POST /engine/open` carries `kind` (`xlsx` | `pptx`), which selects the
  worker bundle, the operation allow-list prefix and the save contract
  (`server/engine-pool.mjs` `KINDS`).
- Runs document operations (workbook reads/formulas/recalc/save-edits; slide
  reads/edits/structure/save) with the same operation allow-list as the Office
  edition (`server/boundary.mjs`).
- On save, PUTs the new package to the platform
  (`/api/public/office/engine/docs/:id/content`) with the user's token,
  `If-Match`, an idempotency key and the kind's MIME type. The platform owns
  revisions.
- Holds no database, no bucket credentials, and never calls a model.

## Layout

```
server/                HTTP app, token auth, engine pool, worker hosts (Electron shim)
server/host/worker.ts  Sheets worker (workbook:* channels, xlsx-sidecar)
server/host/worker-slides.ts  Slides worker (slides:* channels, in-process pptx engine)
server/stubs/          inert replacements for desktop-only modules (AI IPC, cloud search)
vendor/sheets/         retained Sheets engine sources (main, gateway, domain, shared)
vendor/slides/         retained Slides engine sources (main incl. ops/, shared)
vendor/packages/       the @genoffice packages the engines import (full pptx-engine,
                       pptx-render, bundled Carlito fonts from ui/)
native/xlsx-engine     Rust sidecar crate (pinned toolchain 1.94.1)
tests/                 package-fidelity gate (xlsx change-plan writer, pptx text edit)
build.mjs              esbuild: .build/native-worker.cjs, .build/slides-worker.cjs,
                       .build/engine.cjs, .build/fidelity.cjs
Dockerfile             Linux x86_64 image: rust build stage + node runtime + fonts
```

## Slides notes

- Text layout is shaped with real font files. The image installs Liberation
  (Arial/Times/Courier metrics), Carlito/Caladea (Calibri/Cambria metrics) and
  DejaVu; the engine also carries the bundled Carlito faces. Non-Latin scripts
  fall back to DejaVu until `fonts-noto-cjk` / `fonts-noto-core` are added to
  the Dockerfile (larger image).
- Desktop-only paths (native file dialogs, OS clipboard, presenter audience
  window, video/3D insertion, cloud page generation) are served by the browser
  host in the platform (`src/office/slides/web/host/session.ts`) or reported
  as unavailable; the engine never opens dialogs.
- PDF export and printing are browser-side raster output of the rendered
  slides (`rendered-pdf.mjs`), the same as the Office web edition.

## Unsaved-work copies

The engine holds a document's unsaved state (a Slides deck, a Sheets edit session), and
`pagehide` closes the session, so the platform keeps a separate recovery copy of it.
Two read-only operations hand the package to the browser, which stores it through the
platform's recovery route (`/api/office/docs/:id/recovery`, `If-Match` = the saved revision):

- `slides:recovery-bytes`: the deck as a package when it has unsaved edits.
- `workbook:recovery-bytes`: the copy written by the preceding `workbook:write-recovery`.

Both answer `{ found, dirty, size, bytes? }`, with `bytes` only up to 3 MB
(`RECOVERY_MAX_BYTES` in `server/boundary.mjs`; larger work reports its size only). They
are `read` operations, so they never advance the editing sequence or mark the document
dirty, and they never touch the saved revisions. Restoring opens the recovery copy as the
working document (`POST /engine/open` with the recovery grant); only a user Save writes a
revision. An engine that predates these operations answers 403 and the browser simply
writes no recovery copy for that session.

## Local development

```powershell
cd services/office-engine
npm install --ignore-scripts
node build.mjs
# Windows binary from the Office checkout, or build one: npm run build:native
$env:OFFICE_PLATFORM_URL="http://localhost:8080"
$env:OFFICE_ENGINE_ALLOWED_ORIGINS="http://localhost:8080"
$env:HOST="127.0.0.1"; $env:PORT="8790"
$env:XLSX_SIDECAR_PATH="<path>\xlsx-sidecar.exe"
node server/start.mjs
```

Platform `.env`: `OFFICE_ENGINE_URL=http://127.0.0.1:8790`. Without
`OFFICE_ENGINE_JWT_PRIVATE_KEY` the platform dev server signs with an ephemeral
key per process; restart the engine if the dev server restarts and tokens stop
verifying.

## Container

```bash
docker build -t office-engine:local services/office-engine
docker run --rm -p 8790:8790 -e OFFICE_PLATFORM_URL=https://<app-domain> office-engine:local
```

The Rust stage compiles the sidecar's bzip2/zstd C sources with the image's
gcc; no prebuilt binary is checked in.

## AWS deployment (as deployed for testing)

1. **Image**: `infra/office-engine/office-engine-build.cfn.yaml` creates the ECR
   repository and a CodeBuild project. Zip this directory (honoring
   `.dockerignore`) to `s3://<artifact bucket>/office-engine/source.zip`, deploy
   the stack, then `aws codebuild start-build --project-name
   <prefix>-<env>-office-engine --environment-variables-override
   name=IMAGE_TAG,value=<tag>`. The Dockerfile's Rust stage and the fidelity
   gate run inside the build (about 5 minutes on BUILD_GENERAL1_LARGE).
2. **Service**: deploy `infra/office-engine/office-engine.cfn.yaml` with
   `infra/office-engine/parameters/<env>.parameters.json` plus
   `OriginKeyHeaderValue` (a random shared secret, passed at deploy time, never
   committed). With no ALB certificate the stack exposes an HTTP listener on
   port 80 reachable only from CloudFront's origin-facing prefix list and only
   with the origin key header; CloudFront then uses `http-only` to this origin.
   Attach an ACM certificate (`AlbCertificateArn`) as soon as a validated domain
   exists and switch the platform's `OfficeEngineOriginProtocol` to
   `https-only`. In a default VPC without NAT set `AssignPublicIp=ENABLED`.
3. **Secrets**: an RS256 PKCS#8 key as `{"privateKeyPem": ..., "kid": ...}` in
   Secrets Manager (`litai/<env>/office-engine-jwt`); the same origin key stored
   as `officeEngineOriginKey` in the platform's CloudFront origin secret.
4. **Platform** (`infra/app/app-runtime.cfn.yaml` parameters):
   `OfficeEngineOriginDomain` = the stack's `LoadBalancerDnsName`,
   `OfficeEngineOriginProtocol`, `OfficeEnginePublicUrl` = the platform origin
   (browsers reach the engine through its `/engine/*` behavior; it is also the
   token issuer), `OfficeEngineJwtSecretArn`. The Lambda reads the signing key
   from Secrets Manager at runtime (`OFFICE_ENGINE_JWT_SECRET_ARN`), never from
   an environment variable (4 KB cap).

Limits to know: API Gateway REST caps request bodies at 10 MB, so packages
above that cannot be uploaded or saved back through the platform until the
origin is fronted differently.

## Limits

Per task: 24 open documents (`OFFICE_ENGINE_MAX_SESSIONS`), 4 per user, 45 s
per operation, idle sessions closed after 30 minutes, 30 MB packages. ALB
stickiness keeps a browser on the task that holds its session. Measured
locally, Slides read and edit operations round-trip in 3–12 ms; deployed, add
the browser-to-ALB latency.

## Tests

```powershell
node build.mjs && node --test server/fidelity.test.mjs
```

The gate applies one surgical edit per fixture (five workbooks, one two-slide
deck) and asserts every other package entry survives byte-identical. It also
runs inside the Docker build.
