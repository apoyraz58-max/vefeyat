// Builds the Office add-in into dist/addin.
//   BASE_URL  where taskpane.html will be served (default: GitHub Pages).
//             Example for a local server: BASE_URL=https://localhost:3000/ npm run build
import { build } from "esbuild";
import { existsSync } from "node:fs";
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";

const root = new URL("../", import.meta.url);
const out = new URL("dist/addin/", root);
const pkg = JSON.parse(await readFile(new URL("package.json", root), "utf8"));

let baseUrl = process.env.BASE_URL ?? "https://apoyraz58-max.github.io/vefeyat/";
if (!baseUrl.endsWith("/")) baseUrl += "/";
if (!baseUrl.startsWith("https://")) throw new Error("BASE_URL must use https:// (Office requires it).");

// Manifest versions are four numbers: 0.4.<CI run>.0
// The third number is the CI run number, so every published build has a new
// version and Word replaces its cached copy of the ribbon.
const numbers = pkg.version.split(/[.-]/).filter((part) => /^\d+$/.test(part)).slice(0, 2);
while (numbers.length < 2) numbers.push("0");
const manifestVersion = [...numbers, process.env.GITHUB_RUN_NUMBER ?? "0", "0"].join(".");

await rm(out, { recursive: true, force: true });
await mkdir(new URL("assets/", out), { recursive: true });

const version = process.env.GITHUB_RUN_NUMBER ? `${pkg.version.replace(/-dev$/, "")}.${process.env.GITHUB_RUN_NUMBER}` : pkg.version;

for (const entry of ["taskpane", "commands"]) {
  await build({
    entryPoints: [new URL(`src/addin/${entry}.ts`, root).pathname],
    outfile: new URL(`${entry}.js`, out).pathname,
    bundle: true,
    format: "iife",
    target: ["es2018"],
    minify: true,
    sourcemap: true,
    define: { __VERSION__: JSON.stringify(version) },
    logLevel: "warning",
  });
}

// Word keeps add-in files in its own cache: every reference carries the
// version, and version.json tells an open pane that a newer one exists.
await cp(new URL("src/addin/taskpane.css", root), new URL("taskpane.css", out));
for (const file of ["taskpane.html", "commands.html"]) {
  const html = (await readFile(new URL(`src/addin/${file}`, root), "utf8"))
    .replace(/(src|href)="(taskpane|commands)\.(js|css)"/g, `$1="$2.$3?v=${version}"`);
  await writeFile(new URL(file, out), html);
}
await writeFile(new URL("version.json", out), JSON.stringify({ version }));
await cp(new URL("assets/icons/", root), new URL("assets/icons/", out), { recursive: true });
await cp(new URL("assets/logo/", root), new URL("assets/logo/", out), { recursive: true });
for (const size of [16, 32, 64, 80, 128]) {
  await cp(new URL(`assets/icon-${size}.png`, root), new URL(`assets/icon-${size}.png`, out));
}

// Catalog published with the add-in (loaded automatically by the pane).
const catalog = new URL("data/catalog.csv", root);
if (existsSync(catalog)) {
  await cp(catalog, new URL("catalog.csv", out));
  console.log("Embedded catalog: data/catalog.csv");
} else {
  console.log("No data/catalog.csv: users choose a CSV on first use.");
}

// manifest.xml is the regular add-in; manifest-kisayol.xml the keyboard-shortcut
// trial (shared runtime), which also needs shortcuts.json next to it.
for (const name of ["manifest", "manifest-kisayol"]) {
  const template = await readFile(new URL(`manifest/${name}.template.xml`, root), "utf8");
  const manifest = template.replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{VERSION}}", manifestVersion);
  await writeFile(new URL(`${name}.xml`, out), manifest);
}
await cp(new URL("manifest/shortcuts.json", root), new URL("shortcuts.json", out));

console.log(`Built Vefeyât ${pkg.version} (manifest ${manifestVersion}) for ${baseUrl}`);

// Test thesis for trying every command in real Word (scripts/test-doc.mjs).
const testDoc = new URL("docs/vefeyat-deneme.docx", root);
if (existsSync(testDoc)) await cp(testDoc, new URL("vefeyat-deneme.docx", out));
