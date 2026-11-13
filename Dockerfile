# Office engine service: Node 22 host + Rust xlsx-sidecar, Linux x86_64.
# Stage 1 builds the sidecar (bzip2/zstd C sources compile with the image's gcc).
FROM rust:1.94.1-bookworm AS native
WORKDIR /src
COPY native/xlsx-engine ./native/xlsx-engine
RUN cargo build --release --locked --manifest-path native/xlsx-engine/Cargo.toml \
 && strip native/xlsx-engine/target/release/xlsx-sidecar

# Stage 2 bundles the engine.
FROM node:22-bookworm-slim AS build
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci --ignore-scripts --no-audit --no-fund
COPY build.mjs tsconfig.json ./
COPY server ./server
COPY vendor ./vendor
COPY tests ./tests
# The package-fidelity gate runs in every image build: one surgical edit per
# fixture must leave every other ZIP entry byte-identical.
RUN node build.mjs && node --test server/fidelity.test.mjs && npm prune --omit=dev

# Stage 3 runtime.
FROM node:22-bookworm-slim
ENV NODE_ENV=production PORT=8790 OFFICE_ENGINE_SCRATCH=/scratch \
    XLSX_SIDECAR_PATH=/app/native/xlsx-sidecar
# Slides shapes text with real font files (harfbuzz + opentype) so layout and
# line breaks match PowerPoint. Liberation covers Arial/Times/Courier metrics,
# Carlito/Caladea cover Calibri/Cambria, DejaVu is the last-resort fallback.
# Non-Latin scripts (CJK, Arabic, Devanagari) fall back to DejaVu until
# fonts-noto-cjk / fonts-noto-core are added here.
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      fontconfig fonts-liberation fonts-liberation2 fonts-crosextra-carlito fonts-crosextra-caladea fonts-dejavu-core \
 && fc-cache -f \
 && rm -rf /var/lib/apt/lists/* \
 && groupadd -r office && useradd -r -g office -d /app office \
 && mkdir -p /scratch && chown office:office /scratch
WORKDIR /app
COPY --from=build --chown=office:office /app/package.json ./
COPY --from=build --chown=office:office /app/node_modules ./node_modules
COPY --from=build --chown=office:office /app/.build ./.build
COPY --from=build --chown=office:office /app/server ./server
COPY --from=native --chown=office:office /src/native/xlsx-engine/target/release/xlsx-sidecar ./native/xlsx-sidecar
USER office
EXPOSE 8790
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s CMD node -e "fetch('http://127.0.0.1:8790/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"
CMD ["node", "server/start.mjs"]
