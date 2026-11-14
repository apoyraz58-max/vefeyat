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

// Manifest versions are four numbers: 0.4.0-dev -> 0.4.0.0
const numbers = pkg.version.split(/[.-]/).filter((part) => /^\d+$/.test(part)).slice(0, 4);
while (numbers.length < 4) numbers.push("0");
const manifestVersion = numbers.join(".");

await rm(out, { recursive: true, force: true });
await mkdir(new URL("assets/", out), { recursive: true });

for (const entry of ["taskpane", "commands"]) {
  await build({
    entryPoints: [new URL(`src/addin/${entry}.ts`, root).pathname],
    outfile: new URL(`${entry}.js`, out).pathname,
    bundle: true,
    format: "iife",
    target: ["es2018"],
    minify: true,
    sourcemap: true,
    define: { __VERSION__: JSON.stringify(pkg.version) },
    logLevel: "warning",
  });
}

for (const file of ["taskpane.html", "taskpane.css", "commands.html"]) {
  await cp(new URL(`src/addin/${file}`, root), new URL(file, out));
}
await cp(new URL("assets/icons/", root), new URL("assets/icons/", out), { recursive: true });
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

const template = await readFile(new URL("manifest/manifest.template.xml", root), "utf8");
const manifest = template.replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{VERSION}}", manifestVersion);
await writeFile(new URL("manifest.xml", out), manifest);

console.log(`Built Vefeyât ${pkg.version} (manifest ${manifestVersion}) for ${baseUrl}`);
