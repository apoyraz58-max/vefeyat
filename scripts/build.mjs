// Builds the Office add-in into dist/addin.
//   BASE_URL  where taskpane.html will be served (default: GitHub Pages).
//             Example for a local server: BASE_URL=https://localhost:3000/ npm run build
import { build } from "esbuild";
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

await build({
  entryPoints: [new URL("src/addin/taskpane.ts", root).pathname],
  outfile: new URL("taskpane.js", out).pathname,
  bundle: true,
  format: "iife",
  target: ["es2018"],
  minify: true,
  sourcemap: true,
  define: { __VERSION__: JSON.stringify(pkg.version) },
  logLevel: "warning",
});

for (const file of ["taskpane.html", "taskpane.css", "commands.html"]) {
  await cp(new URL(`src/addin/${file}`, root), new URL(file, out));
}
await cp(new URL("assets/", root), new URL("assets/", out), { recursive: true });

const template = await readFile(new URL("manifest/manifest.template.xml", root), "utf8");
const manifest = template.replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{VERSION}}", manifestVersion);
await writeFile(new URL("manifest.xml", out), manifest);

console.log(`Built Vefeyât ${pkg.version} (manifest ${manifestVersion}) for ${baseUrl}`);
