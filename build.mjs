// Bundles the retained Sheets engine for the Node worker host and the small
// helpers the HTTP layer needs (blank workbook, CSV import, registry). The
// Electron surface is replaced by server/host/electron.ts; the desktop-only
// AI IPC and Genspark search are replaced by inert stubs because inference
// runs on the platform, never in this service.
import { build } from "esbuild";
import { mkdir, readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { join, relative, resolve } from "node:path";

const root = resolve(import.meta.dirname);
const req = createRequire(join(root, "package.json"));
const vendor = (p) => join(root, "vendor", p);

const assets = {
  name: "native-assets",
  setup(b) {
    b.onResolve({ filter: /\?(asset|raw)$/ }, async (args) => {
      const raw = args.path.endsWith("?raw");
      const bare = args.path.replace(/\?(asset|raw)$/, "");
      // Bundled fonts (Slides text metrics) live in the vendored ui package.
      const uiFont = bare.match(/^@genoffice\/ui\/fonts\/(.+)$/);
      const found = uiFont
        ? vendor(`packages/ui/src/fonts/${uiFont[1]}`)
        : bare.startsWith(".")
          ? resolve(args.resolveDir, bare)
          : req.resolve(bare);
      return { path: found, namespace: raw ? "raw-source" : "native-asset" };
    });
    b.onLoad({ filter: /.*/, namespace: "native-asset" }, (args) => ({
      contents: `import {resolve} from 'node:path'; export default resolve(__dirname, ${JSON.stringify(
        "../" + relative(root, args.path).replaceAll("\\", "/"),
      )});`,
      loader: "js",
    }));
    b.onLoad({ filter: /.*/, namespace: "raw-source" }, async (args) => ({
      contents: `export default ${JSON.stringify(await readFile(args.path, "utf8"))};`,
      loader: "js",
    }));
  },
};

const alias = {
  electron: join(root, "server/host/electron.ts"),
  "@genoffice/agent-core": vendor("packages/agent-core/src/index.ts"),
  "@genoffice/ai-provider": vendor("packages/ai-provider/src/index.ts"),
  "@genoffice/ai-search": join(root, "server/stubs/ai-search.ts"),
  "@genoffice/electron-utils": vendor("packages/electron-utils/src/index.ts"),
  "@genoffice/file-parse": vendor("packages/file-parse/src/index.ts"),
  "@genoffice/i18n": vendor("packages/i18n/src/index.ts"),
  "@genoffice/project-store": vendor("packages/project-store/src/index.ts"),
  "@genoffice/docx-engine": vendor("packages/docx-engine/src/index.ts"),
  "@genoffice/docx-engine/math": vendor("packages/docx-engine/src/math.ts"),
  "@genoffice/docx-engine/metafile": vendor("packages/docx-engine/src/metafile.ts"),
  "@genoffice/pptx-engine": vendor("packages/pptx-engine/src/index.ts"),
  "@genoffice/pptx-engine/custgeom": vendor("packages/pptx-engine/src/custgeom.ts"),
  "@genoffice/pptx-engine/table-grid": vendor("packages/pptx-engine/src/table-grid.ts"),
  "@genoffice/pptx-engine/identity": vendor("packages/pptx-engine/src/identity.ts"),
  "@genoffice/pptx-engine/background-promote": vendor("packages/pptx-engine/src/background-promote.ts"),
  "@genoffice/pptx-render": vendor("packages/pptx-render/src/index.ts"),
  "@genoffice/pptx-render/preset-geometry": vendor("packages/pptx-render/src/preset-geometry.ts"),
};

/** Desktop-only modules the worker must never execute here. */
const stubs = {
  name: "engine-stubs",
  setup(b) {
    b.onResolve({ filter: /(^|\/)sw-ipc(\.ts)?$/ }, () => ({ path: join(root, "server/stubs/sw-ipc.ts") }));
    b.onResolve({ filter: /(^|\/)sw-connection(\.ts)?$/ }, () => ({
      path: join(root, "server/stubs/sw-connection.ts"),
    }));
  },
};

await mkdir(join(root, ".build"), { recursive: true });
const shared = {
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node22",
  logLevel: "warning",
  plugins: [stubs, assets],
  alias,
  // Desktop-only code paths (PDF export, PDF parsing) stay unresolved at
  // bundle time; they are dynamic imports that the browser edition never calls.
  external: ["@embedpdf/pdfium", "harfbuzzjs", "pdfjs-dist", "pdf-lib"],
  define: { "import.meta.url": "__fileUrl", "process.env.GENOFFICE_DEBUG_HOOKS": '"0"' },
  banner: { js: "const __fileUrl=require('node:url').pathToFileURL(__filename).href;" },
};
await build({ ...shared, entryPoints: [join(root, "server/host/worker.ts")], outfile: join(root, ".build/native-worker.cjs") });
await build({ ...shared, entryPoints: [join(root, "server/host/worker-slides.ts")], outfile: join(root, ".build/slides-worker.cjs") });
await build({ ...shared, entryPoints: [join(root, "server/engine.ts")], outfile: join(root, ".build/engine.cjs") });
await build({ ...shared, entryPoints: [join(root, "tests/fidelity.ts")], outfile: join(root, ".build/fidelity.cjs") });
console.log("Office engine built: .build/native-worker.cjs, .build/slides-worker.cjs, .build/engine.cjs, .build/fidelity.cjs");
