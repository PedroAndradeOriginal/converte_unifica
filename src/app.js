const { PDFDocument, rgb } = window.PDFLib;
const $ = id => document.getElementById(id);
const state = { excel: [], word: [], groups: [], results: [] };
const pending = new Map();
let engineReady = false;
let readyPromise;
const webProtocol = location.protocol === "https:" || location.hostname === "localhost" || location.hostname === "127.0.0.1";

const stem = name => name.replace(/\.[^.]+$/, "");
function keyFor(name) {
  const parts = stem(name).trim().replace(/^_+/, "").replace(/^PPU_II-B_/i, "")
    .replace(/_(EXCEL|WORD)$/i, "").replace(/\s+/g, "").toUpperCase().split("-");
  if (["5135", "5950"].includes(parts[0])) parts.shift();
  if (["GI", "G1"].includes(parts[0])) parts.shift();
  if (parts.length && /^\d+$/.test(parts.at(-1))) parts[parts.length - 1] = String(Number(parts.at(-1)));
  return parts.join("-");
}
function escapeHtml(value) { const div = document.createElement("div"); div.textContent = value; return div.innerHTML; }
function prettySize(bytes) { return bytes < 1048576 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1048576).toFixed(1)} MB`; }
function setStatus(text, type = "info") { const e = $("statusMessage"); e.textContent = text; e.className = `status-message ${type}`; }

window.addEventListener("message", event => {
  if (event.source !== $("officeConverter").contentWindow || !event.data) return;
  const data = event.data;
  if (data.type === "ready") {
    engineReady = true;
    window.dispatchEvent(new Event("office-engine-ready"));
  }
  if (data.type === "error" && !data.requestId) {
    $("engineStatus").textContent = "Falha ao carregar conversor";
  }
  const job = pending.get(data.requestId);
  if (!job) return;
  if (data.type === "result") {
    pending.delete(data.requestId);
    const bytes = new Uint8Array(data.data);
    if (job.trimToFirstPage) keepFirstPdfPage(bytes).then(job.resolve, job.reject);
    else job.resolve(bytes);
  }
  if (data.type === "error") { pending.delete(data.requestId); job.reject(new Error(data.error || "Falha ao exportar documento")); }
});

function waitForEngine() {
  if (!webProtocol) {
    return Promise.reject(new Error("Este sistema precisa ser aberto pelo endereço do GitHub Pages ou por localhost. Ele não funciona abrindo o index.html diretamente pela pasta."));
  }
  if (engineReady) return Promise.resolve();
  if (!readyPromise) readyPromise = new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("O conversor não conseguiu iniciar. Confirme que o sistema está publicado no GitHub Pages e atualize a página com Ctrl+F5.")), 480000);
    window.addEventListener("office-engine-ready", () => { clearTimeout(timer); resolve(); }, { once: true });
  });
  return readyPromise;
}
const XML_NS = {
  main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  officeRel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
};
const parseXml = text => new DOMParser().parseFromString(text, "application/xml");
const xmlText = doc => new XMLSerializer().serializeToString(doc);
const xmlEscape = value => String(value).replace(/[&<>'"]/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&apos;",'"':"&quot;"}[char]));

function cellParts(ref) {
  const match = /^([A-Z]+)(\d+)$/i.exec(ref);
  if (!match) return null;
  let col = 0;
  for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
  return { col: col - 1, row: Number(match[2]) - 1 };
}
function rangeForCell(sheetDoc, ref) {
  const cell = cellParts(ref);
  if (!cell) return null;
  for (const merge of sheetDoc.getElementsByTagNameNS(XML_NS.main, "mergeCell")) {
    const [fromRef, toRef] = merge.getAttribute("ref").split(":");
    const from = cellParts(fromRef), to = cellParts(toRef || fromRef);
    if (from && to && cell.col >= from.col && cell.col <= to.col && cell.row >= from.row && cell.row <= to.row) {
      return { from, to: { col: to.col + 1, row: to.row + 1 } };
    }
  }
  return { from: cell, to: { col: cell.col + 1, row: cell.row + 1 } };
}
async function imageSize(blob) {
  const bitmap = await createImageBitmap(blob);
  const size = { width: bitmap.width, height: bitmap.height };
  bitmap.close();
  return size;
}
async function keepFirstPdfPage(bytes) {
  const source = await PDFDocument.load(bytes, { ignoreEncryption: true });
  if (source.getPageCount() <= 1) return bytes;
  const output = await PDFDocument.create();
  const [page] = await output.copyPages(source, [0]);
  output.addPage(page);
  page.drawRectangle({
    x: 51.72,
    y: 82.94,
    width: 487.30,
    height: 0.48,
    color: rgb(0, 0, 0),
    borderWidth: 0,
  });
  return output.save();
}
function excelDateText(serial) {
  const date = new Date(Date.UTC(1899, 11, 30) + Number(serial) * 86400000);
  if (!Number.isFinite(date.getTime())) return null;
  return `${String(date.getUTCDate()).padStart(2, "0")}/${String(date.getUTCMonth() + 1).padStart(2, "0")}/${date.getUTCFullYear()}`;
}
function pictureAnchor(id, relId, range, size) {
  const cols = range.to.col - range.from.col, rows = range.to.row - range.from.row;
  const cellRatio = Math.max(.1, (cols * 138) / (rows * 38));
  const imageRatio = size.width / size.height;
  let fromCol = range.from.col, fromRow = range.from.row, toCol = range.to.col, toRow = range.to.row;
  let fromColOff = 0, fromRowOff = 0, toColOff = 0, toRowOff = 0;
  if (imageRatio > cellRatio) {
    const usedRows = rows * cellRatio / imageRatio;
    const pad = (rows - usedRows) / 2;
    fromRow += Math.floor(pad); fromRowOff = Math.round((pad % 1) * 38 * 9525);
    const end = range.from.row + pad + usedRows;
    toRow = Math.floor(end); toRowOff = Math.round((end % 1) * 38 * 9525);
  } else {
    const usedCols = cols * imageRatio / cellRatio;
    const pad = (cols - usedCols) / 2;
    fromCol += Math.floor(pad); fromColOff = Math.round((pad % 1) * 138 * 9525);
    const end = range.from.col + pad + usedCols;
    toCol = Math.floor(end); toColOff = Math.round((end % 1) * 138 * 9525);
  }
  const marker = (name, col, colOff, row, rowOff) => `<xdr:${name}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${name}>`;
  return `<xdr:twoCellAnchor editAs="oneCell">${marker("from",fromCol,fromColOff,fromRow,fromRowOff)}${marker("to",toCol,toColOff,toRow,toRowOff)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Imagem em célula ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
}
async function normalizeExcelImages(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const richFile = zip.file("xl/richData/rdrichvalue.xml");
  const richValueRelFile = zip.file("xl/richData/richValueRel.xml");
  const richRelsFile = zip.file("xl/richData/_rels/richValueRel.xml.rels");
  if (!richFile || !richValueRelFile || !richRelsFile) return { buffer, trimToFirstPage: false };

  const stylesFile = zip.file("xl/styles.xml");
  const dateStyles = new Set();
  if (stylesFile) {
    const stylesDoc = parseXml(await stylesFile.async("text"));
    const customDateFormats = new Set([...stylesDoc.getElementsByTagNameNS(XML_NS.main, "numFmt")]
      .filter(node => /[dmy]/i.test(node.getAttribute("formatCode") || ""))
      .map(node => Number(node.getAttribute("numFmtId"))));
    const cellXfs = stylesDoc.getElementsByTagNameNS(XML_NS.main, "cellXfs")[0];
    [...(cellXfs?.children || [])].forEach((xf, index) => {
      const id = Number(xf.getAttribute("numFmtId"));
      if ((id >= 14 && id <= 22) || customDateFormats.has(id)) dateStyles.add(index);
    });
  }

  const richDoc = parseXml(await richFile.async("text"));
  const richValueRelDoc = parseXml(await richValueRelFile.async("text"));
  const richRelsDoc = parseXml(await richRelsFile.async("text"));
  const richValues = [...richDoc.getElementsByTagNameNS("*", "rv")];
  const richValueRefs = [...richValueRelDoc.getElementsByTagNameNS("*", "rel")];
  const richRels = [...richRelsDoc.getElementsByTagNameNS(XML_NS.rel, "Relationship")];
  const images = [];
  for (const rich of richValues) {
    const relationIndex = Number(rich.getElementsByTagNameNS("*", "v")[0]?.textContent);
    const relationId = richValueRefs[relationIndex]?.getAttributeNS(XML_NS.officeRel, "id") || richValueRefs[relationIndex]?.getAttribute("r:id");
    const relation = richRels.find(item => item.getAttribute("Id") === relationId);
    if (!relation) { images.push(null); continue; }
    const target = relation.getAttribute("Target").replace(/^\.\.\//, "xl/");
    const file = zip.file(target);
    images.push(file ? { target: target.replace(/^xl\//, "../"), file, size: await imageSize(await file.async("blob")) } : null);
  }

  const sheetFiles = Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/.test(path));
  for (const sheetPath of sheetFiles) {
    const sheetDoc = parseXml(await zip.file(sheetPath).async("text"));
    const richCells = [...sheetDoc.getElementsByTagNameNS(XML_NS.main, "c")].filter(cell => cell.hasAttribute("vm"));
    if (!richCells.length) continue;
    for (const cell of sheetDoc.getElementsByTagNameNS(XML_NS.main, "c")) {
      if (!dateStyles.has(Number(cell.getAttribute("s"))) || cell.getElementsByTagNameNS(XML_NS.main, "f").length) continue;
      const value = cell.getElementsByTagNameNS(XML_NS.main, "v")[0];
      const formatted = value && excelDateText(value.textContent);
      if (formatted) { cell.setAttribute("t", "str"); value.textContent = formatted; }
    }

    const number = /sheet(\d+)\.xml$/.exec(sheetPath)?.[1];
    const sheetRelsPath = `xl/worksheets/_rels/sheet${number}.xml.rels`;
    const sheetRelsFile = zip.file(sheetRelsPath);
    const sheetRelsDoc = sheetRelsFile ? parseXml(await sheetRelsFile.async("text")) : parseXml(`<Relationships xmlns="${XML_NS.rel}"/>`);
    let drawingRel = [...sheetRelsDoc.getElementsByTagNameNS(XML_NS.rel, "Relationship")].find(rel => rel.getAttribute("Type")?.endsWith("/drawing"));
    if (!drawingRel) continue;
    const drawingPath = `xl/${drawingRel.getAttribute("Target").replace(/^\.\.\//, "")}`;
    const drawingFile = zip.file(drawingPath);
    if (!drawingFile) continue;
    const drawingDoc = parseXml(await drawingFile.async("text"));
    const drawingRelsPath = drawingPath.replace(/\/([^/]+)$/, "/_rels/$1.rels");
    const drawingRelsFile = zip.file(drawingRelsPath);
    const drawingRelsDoc = drawingRelsFile ? parseXml(await drawingRelsFile.async("text")) : parseXml(`<Relationships xmlns="${XML_NS.rel}"/>`);
    const drawingRoot = drawingDoc.documentElement, relRoot = drawingRelsDoc.documentElement;
    let nextId = Math.max(1, ...[...drawingDoc.getElementsByTagNameNS(XML_NS.xdr, "cNvPr")].map(node => Number(node.getAttribute("id")) || 0)) + 1;
    let nextRel = Math.max(0, ...[...drawingRelsDoc.getElementsByTagNameNS(XML_NS.rel, "Relationship")].map(node => Number(node.getAttribute("Id")?.replace(/\D/g, "")) || 0)) + 1;

    for (const cell of richCells) {
      const image = images[Number(cell.getAttribute("vm")) - 1];
      const range = rangeForCell(sheetDoc, cell.getAttribute("r"));
      if (!image || !range) continue;
      const relId = `rId${nextRel++}`;
      const rel = drawingRelsDoc.createElementNS(XML_NS.rel, "Relationship");
      rel.setAttribute("Id", relId); rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"); rel.setAttribute("Target", image.target);
      relRoot.appendChild(rel);
      const fragment = parseXml(`<root xmlns:xdr="${XML_NS.xdr}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${pictureAnchor(nextId++, relId, range, image.size)}</root>`);
      drawingRoot.appendChild(drawingDoc.importNode(fragment.documentElement.firstElementChild, true));
      cell.removeAttribute("t"); cell.removeAttribute("vm");
      [...cell.getElementsByTagNameNS(XML_NS.main, "v")].forEach(value => value.remove());
    }
    zip.file(sheetPath, xmlText(sheetDoc));
    zip.file(drawingPath, xmlText(drawingDoc));
    zip.file(drawingRelsPath, xmlText(drawingRelsDoc));
  }
  return { buffer: await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" }), trimToFirstPage: true };
}
async function convertToPdf(file) {
  await waitForEngine();
  const requestId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
  const originalBuffer = await file.arrayBuffer();
  const normalized = /\.xlsx$/i.test(file.name) ? await normalizeExcelImages(originalBuffer) : { buffer: originalBuffer, trimToFirstPage: false };
  const buffer = normalized.buffer;
  return new Promise((resolve, reject) => {
    pending.set(requestId, { resolve, reject, trimToFirstPage: normalized.trimToFirstPage });
    const filterName = /\.(xlsx|xls|xlsm)$/i.test(file.name) ? "calc_pdf_Export" : "writer_pdf_Export";
    $("officeConverter").contentWindow.postMessage({ type: "convert", buffer, format: "pdf", filterName, requestId }, "*", [buffer]);
  });
}

function addFiles(kind, files) {
  const allowed = kind === "excel" ? /\.(xlsx|xls|xlsm)$/i : /\.(docx|doc)$/i;
  const existing = new Set(state[kind].map(f => `${f.name}|${f.size}|${f.lastModified}`));
  [...files].filter(f => allowed.test(f.name)).forEach(file => {
    const id = `${file.name}|${file.size}|${file.lastModified}`;
    if (!existing.has(id)) { state[kind].push(file); existing.add(id); }
  });
  analyse();
}
function addMixedFiles(files) {
  addFiles("excel", [...files].filter(file => /\.(xlsx|xls|xlsm)$/i.test(file.name)));
  addFiles("word", [...files].filter(file => /\.(docx|doc)$/i.test(file.name)));
}
function analyse() {
  const map = new Map();
  for (const kind of ["excel", "word"]) for (const file of state[kind]) {
    const key = keyFor(file.name);
    if (!map.has(key)) map.set(key, { key, excel: [], word: [] });
    map.get(key)[kind].push(file);
  }
  state.groups = [...map.values()].map(g => ({ ...g, ready: g.excel.length === 1 && g.word.length === 1 })).sort((a,b) => a.key.localeCompare(b.key));
  renderSelection();
}
function renderSelection() {
  const totalFiles = state.excel.length + state.word.length;
  $("fileCount").textContent = totalFiles ? `${totalFiles} arquivo(s): ${state.excel.length} Excel e ${state.word.length} Word` : "Nenhum arquivo selecionado";
  $("clearButton").hidden = state.excel.length + state.word.length === 0;
  const ready = state.groups.filter(g => g.ready);
  $("pairCount").textContent = `${ready.length} ${ready.length === 1 ? "par" : "pares"}`;
  const list = $("pairList");
  if (!state.groups.length) {
    list.className = "pair-list empty";
    list.innerHTML = '<span class="pair-empty-icon" aria-hidden="true"><i></i><i></i></span><strong>Os pares aparecerão aqui</strong><p>Diferenças como “_” no início, espaços, GI/G1 e 5135/5950 são normalizadas.</p>';
  } else {
    list.className = "pair-list";
    list.innerHTML = state.groups.map((g,i) => {
      const label = g.ready ? "Pronto · Excel → Word" : (g.excel.length > 1 || g.word.length > 1 ? "Duplicado" : "Incompleto");
      return `<article class="pair-row"><span class="pair-index">${String(i+1).padStart(2,"0")}</span><span class="pair-copy"><strong>${escapeHtml(g.key)}</strong><span>1. Excel: ${escapeHtml(g.excel[0]?.name || "ausente")}</span><span>2. Word: ${escapeHtml(g.word[0]?.name || "ausente")}</span></span><span class="pair-state ${g.ready ? "" : "warning"}">${label}</span></article>`;
    }).join("");
  }
  $("processButton").disabled = !ready.length;
  if (!state.excel.length && !state.word.length) setStatus("Adicione os documentos originais para começar.");
  else if (ready.length) setStatus(`${ready.length} par(es) pronto(s). A exportação mantém a configuração de impressão e coloca o Excel primeiro.`, "success");
  else setStatus("Nenhum par completo encontrado. Confira os nomes e os tipos dos arquivos.", "warning");
}

async function mergePair(group, current, total) {
  setStatus(`Carregando motor Office e exportando ${current} de ${total}: Excel`, "info");
  const excelPdf = await convertToPdf(group.excel[0]);
  setStatus(`Exportando ${current} de ${total}: Word`, "info");
  const wordPdf = await convertToPdf(group.word[0]);
  const output = await PDFDocument.create();
  for (const bytes of [excelPdf, wordPdf]) {
    const source = await PDFDocument.load(bytes, { ignoreEncryption: true });
    (await output.copyPages(source, source.getPageIndices())).forEach(page => output.addPage(page));
  }
  return output.save();
}
async function processAll() {
  const ready = state.groups.filter(g => g.ready);
  $("processButton").disabled = true; $("processButton").classList.add("is-loading"); $("processLabel").textContent = "Convertendo e unificando";
  state.results.forEach(r => URL.revokeObjectURL(r.url)); state.results = [];
  try {
    for (let i=0; i<ready.length; i++) {
      const bytes = await mergePair(ready[i], i+1, ready.length);
      const name = `${stem(ready[i].excel[0].name).replace(/^_+/, "")}.pdf`;
      const blob = new Blob([bytes], {type:"application/pdf"});
      state.results.push({name, blob, size:blob.size, url:URL.createObjectURL(blob)});
    }
    renderResults();
    setStatus(`${state.results.length} PDF(s) exportado(s) e unido(s). Nenhum arquivo original foi alterado.`, "success");
  } catch(error) { console.error(error); setStatus(`Não foi possível concluir: ${error.message}`, "error"); }
  finally { $("processButton").disabled = false; $("processButton").classList.remove("is-loading"); $("processLabel").textContent = "Converter e unificar documentos"; }
}
function renderResults() {
  $("resultsSection").hidden = !state.results.length;
  $("resultSummary").textContent = `${state.results.length} arquivo(s) gerado(s).`;
  $("resultList").innerHTML = state.results.map((r,i) => `<article class="result-row"><span class="pdf-badge">PDF</span><span class="result-copy"><strong>${escapeHtml(r.name)}</strong><span>${prettySize(r.size)} · Excel primeiro</span></span><button class="download-button" data-index="${i}">Baixar PDF</button></article>`).join("");
  $("resultList").querySelectorAll("[data-index]").forEach(b => b.onclick = () => download(state.results[Number(b.dataset.index)]));
  $("resultsSection").scrollIntoView({behavior:"smooth",block:"start"});
}
function download(r) { const a=document.createElement("a"); a.href=r.url; a.download=r.name; a.click(); }
async function downloadZip() { const zip=new JSZip(); state.results.forEach(r=>zip.file(r.name,r.blob)); const blob=await zip.generateAsync({type:"blob",compression:"DEFLATE"}); const url=URL.createObjectURL(blob); download({url,name:`PDFS_UNIFICADOS_${state.results.length}_ARQUIVOS.zip`}); setTimeout(()=>URL.revokeObjectURL(url),1000); }
function configureDropZone(kind, zoneId, inputId) {
  const zone=$(zoneId), input=$(inputId); zone.onclick=()=>input.click();
  zone.onkeydown=e=>{if(["Enter"," "].includes(e.key)){e.preventDefault();input.click();}};
  input.onchange=e=>{addFiles(kind,e.target.files);e.target.value="";};
  ["dragenter","dragover"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.classList.add("is-dragging");}));
  ["dragleave","drop"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.classList.remove("is-dragging");}));
  zone.addEventListener("drop",e=>addFiles(kind,e.dataTransfer.files));
}
function configureMixedDropZone() {
  const zone = $("filesDropZone"), input = $("filesInput");
  zone.onclick = () => input.click();
  zone.onkeydown = event => { if (["Enter", " "].includes(event.key)) { event.preventDefault(); input.click(); } };
  input.onchange = event => { addMixedFiles(event.target.files); event.target.value = ""; };
  ["dragenter", "dragover"].forEach(type => zone.addEventListener(type, event => { event.preventDefault(); zone.classList.add("is-dragging"); }));
  ["dragleave", "drop"].forEach(type => zone.addEventListener(type, event => { event.preventDefault(); zone.classList.remove("is-dragging"); }));
  zone.addEventListener("drop", event => addMixedFiles(event.dataTransfer.files));
}
configureMixedDropZone();
$("clearButton").onclick=()=>{state.excel=[];state.word=[];state.groups=[];state.results.forEach(r=>URL.revokeObjectURL(r.url));state.results=[];$("resultsSection").hidden=true;renderSelection();};
$("processButton").onclick=processAll; $("downloadAllButton").onclick=downloadZip; $("privacyButton").onclick=()=>$("privacyDialog").showModal(); $("closeDialog").onclick=()=>$("privacyDialog").close();
renderSelection();

if (!webProtocol) {
  setStatus("Abra o sistema pelo endereço do GitHub Pages. A execução direta pelo arquivo index.html não é compatível com o motor Office.", "error");
} else if (window.crossOriginIsolated) {
  $("officeConverter").src = "converter/index.html";
}
