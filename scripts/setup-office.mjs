import { mkdir, copyFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const publicDir = resolve(root, "public");
const wasmDir = resolve(publicDir, "converter/wasm");
const vendorDir = resolve(publicDir, "converter/vendor/zetajs");
await mkdir(wasmDir, { recursive: true });
await mkdir(vendorDir, { recursive: true });

async function download(url, target) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${response.status} ao baixar ${url}`);
  const data = new Uint8Array(await response.arrayBuffer());
  await writeFile(target, data);
  console.log(`Baixado ${target} (${Math.round(data.byteLength / 1048576)} MB)`);
}

async function prepareConverter() {
  const url = "https://raw.githubusercontent.com/erseco/document-converter/main/docs/converter.js";
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${response.status} ao baixar ${url}`);
  let source = await response.text();
  source = source
    .replace("const filterName = filterMap[format.toLowerCase()] || 'writer_pdf_Export';", "const filterName = e.data.filterName || filterMap[format.toLowerCase()] || 'writer_pdf_Export';")
    .replace("export async function convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin) {", "export async function convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin, filterName) {")
    .replace("requestId: requestId\n        });", "requestId: requestId,\n            filterName: filterName\n        });")
    .replace("await convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin);", "await convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin, data.filterName);");
  await writeFile(resolve(publicDir, "converter/converter.js"), source);
}

await prepareConverter();
await download("https://raw.githubusercontent.com/erseco/document-converter/main/docs/coi-serviceworker.js", resolve(publicDir, "coi-serviceworker.js"));
for (const name of ["soffice.js", "soffice.wasm", "soffice.data", "soffice.data.js.metadata"]) {
  await download(`https://cdn.zetaoffice.net/zetaoffice_latest/${name}`, resolve(wasmDir, name));
}
for (const name of ["zeta.js", "zetaHelper.js"]) {
  await copyFile(resolve(root, "node_modules/zetajs/source", name), resolve(vendorDir, name));
}
