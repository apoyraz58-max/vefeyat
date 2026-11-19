// Builds the Office add-in into dist/addin.
//   BASE_URL  where taskpane.html will be served (default: GitHub Pages).
//             Example for a local server: BASE_URL=https://localhost:3000/ npm run build
//   CHANNEL   "beta": the closed-beta add-in (own Id and name, so it sits beside
//             the development one), numbered by BETA_NUMBER instead of the CI run.
//   OUT_DIR   output folder (default dist/addin/).
import { build } from "esbuild";
import { strToU8, zipSync } from "fflate";
import { existsSync } from "node:fs";
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";

const root = new URL("../", import.meta.url);
const channel = process.env.CHANNEL === "beta" ? "beta" : "dev";
const out = process.env.OUT_DIR ? new URL(`file://${process.env.OUT_DIR.replace(/\/?$/, "/")}`) : new URL("dist/addin/", root);
const pkg = JSON.parse(await readFile(new URL("package.json", root), "utf8"));

let baseUrl = process.env.BASE_URL ?? "https://apoyraz58-max.github.io/vefeyat/";
if (!baseUrl.endsWith("/")) baseUrl += "/";
if (!baseUrl.startsWith("https://")) throw new Error("BASE_URL must use https:// (Office requires it).");

// Manifest versions are four numbers: 0.4.<CI run>.0
// The third number is the CI run number, so every published build has a new
// version and Word replaces its cached copy of the ribbon.
const numbers = pkg.version.split(/[.-]/).filter((part) => /^\d+$/.test(part)).slice(0, 2);
while (numbers.length < 2) numbers.push("0");
// The beta keeps its number until it is promoted again: BETA_NUMBER is the beta branch's commit count.
const build_number = channel === "beta" ? (process.env.BETA_NUMBER ?? "0") : (process.env.GITHUB_RUN_NUMBER ?? "0");
const manifestVersion = [...numbers, build_number, "0"].join(".");

await rm(out, { recursive: true, force: true });
await mkdir(new URL("assets/", out), { recursive: true });

const base = pkg.version.replace(/-dev$/, "");
const version =
  channel === "beta" ? `${base}-beta.${build_number}` : process.env.GITHUB_RUN_NUMBER ? `${base}.${process.env.GITHUB_RUN_NUMBER}` : pkg.version;

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

// manifest.xml: Word 2408 and later (shared runtime, keyboard shortcuts; it names
// shortcuts.json, published next to it). manifest-uyumlu.xml: older Word, the
// same add-in without shortcuts. The beta is another add-in beside the
// development one: own Id, name and tab.
const BETA = { id: "6f3c2a4e-8b1d-4c7a-9e52-0d9a7b3e5f23", name: "Vefeyât (beta)" };
for (const [file, source] of [["manifest.xml", "manifest.template.xml"], ["manifest-uyumlu.xml", "manifest-uyumlu.template.xml"]]) {
  let manifest = (await readFile(new URL(`manifest/${source}`, root), "utf8")).replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{VERSION}}", manifestVersion);
  if (channel === "beta") {
    manifest = manifest
      .replace(/<Id>[^<]+<\/Id>/, `<Id>${BETA.id}</Id>`)
      .replace('<DisplayName DefaultValue="Vefeyât"/>', `<DisplayName DefaultValue="${BETA.name}"/>`)
      .replace('<bt:String id="Tab" DefaultValue="Vefeyât"/>', `<bt:String id="Tab" DefaultValue="${BETA.name}"/>`);
    if (!manifest.includes(BETA.id) || !manifest.includes(`DefaultValue="${BETA.name}"`)) throw new Error(`beta ${file}: Id or name not replaced`);
  }
  await writeFile(new URL(file, out), manifest);
}
await cp(new URL("manifest/shortcuts.json", root), new URL("shortcuts.json", out));

// Windows installers: double-click to register the add-in with Word (HKCU WEF\Developer,
// no administrator rights), and to remove it. Windows line endings for cmd.exe.
const addinId = channel === "beta" ? BETA.id : /<Id>([^<]+)<\/Id>/.exec(await readFile(new URL("manifest/manifest.template.xml", root), "utf8"))[1];
const addinName = channel === "beta" ? BETA.name : "Vefeyât";
const installers = {};
for (const file of ["Vefeyat-Kur.cmd", "Vefeyat-Kaldir.cmd"]) {
  const text = (await readFile(new URL(`installer/${file}`, root), "utf8"))
    .replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{ID}}", addinId).replaceAll("{{NAME}}", addinName)
    .replace(/\r?\n/g, "\r\n");
  if (/\{\{/.test(text)) throw new Error(`${file}: placeholder left`);
  await writeFile(new URL(file, out), text);
  installers[file] = strToU8(text);
}
// Browsers often refuse to download a bare .cmd: the same files in a zip.
installers["OKUBENI.txt"] = strToU8(
  [
    `${addinName} kurulumu`,
    "",
    "1. Bu ZIP dosyasina sag tiklayip 'Tumunu ayikla' deyin (ya da acip dosyalari masaustune surukleyin).",
    "2. Word'u kapatin.",
    "3. Vefeyat-Kur.cmd dosyasina cift tiklayin. Windows 'bilgisayarinizi korudu' derse: Ek bilgi -> Yine de calistir.",
    "4. Pencerede 'Tamam' yazinca Word'u acin.",
    "5. Seritte sekme gorunmezse (yalniz ilk seferde): Giris -> Eklentiler (ya da Ekle -> Eklentilerim) -> Diger Eklentiler",
    `   -> 'Gelistirici Eklentileri' bolumunde '${addinName}' -> Ekle.`,
    "",
    "Kaldirmak icin: Vefeyat-Kaldir.cmd",
    "Sorun ve oneriler: apoyraz@cumhuriyet.edu.tr",
  ].join("\r\n"),
);
await writeFile(new URL("Vefeyat-Kurulum.zip", out), zipSync(installers));

console.log(`Built Vefeyât ${version} [${channel}] (manifest ${manifestVersion}) for ${baseUrl}`);

// Test thesis for trying every command in real Word (scripts/test-doc.mjs).
const testDoc = new URL("docs/vefeyat-deneme.docx", root);
if (existsSync(testDoc)) await cp(testDoc, new URL("vefeyat-deneme.docx", out));
