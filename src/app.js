const { jsPDF } = window.jspdf;
const { PDFDocument } = window.PDFLib;

const $ = (id) => document.getElementById(id);
const state = { files: [], groups: [], results: [] };
const accepted = new Set(["xlsx", "xlsm", "docx", "pdf"]);

function ext(name) { return name.split(".").pop().toLowerCase(); }
function stem(name) { return name.replace(/\.[^.]+$/, ""); }
function keyFor(name) {
  let value = stem(name).trim().replace(/^_+/, "").replace(/^PPU_II-B_/i, "");
  value = value.replace(/_(EXCEL|WORD)$/i, "").replace(/\s+/g, "").toUpperCase();
  const parts = value.split("-");
  if (["5135", "5950"].includes(parts[0])) parts.shift();
  if (["GI", "G1"].includes(parts[0])) parts.shift();
  if (parts.length && /^\d+$/.test(parts.at(-1))) parts[parts.length - 1] = String(Number(parts.at(-1)));
  return parts.join("-");
}
function prettySize(bytes) { return bytes < 1048576 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1048576).toFixed(1)} MB`; }
function setStatus(text, type = "info") { const el = $("statusMessage"); el.textContent = text; el.className = `status-message ${type}`; }
function pdfOrder(file) {
  const name = stem(file.name).toUpperCase();
  if (/(^|[_\-\s])EXCEL($|[_\-\s])/.test(name)) return 0;
  if (/(^|[_\-\s])WORD($|[_\-\s])/.test(name)) return 1;
  return 2;
}

function analyse() {
  const map = new Map();
  state.files.forEach(file => {
    const key = keyFor(file.name);
    if (!map.has(key)) map.set(key, { key, excel: [], word: [], pdf: [] });
    const type = ext(file.name);
    if (["xlsx", "xlsm"].includes(type)) map.get(key).excel.push(file);
    else if (type === "docx") map.get(key).word.push(file);
    else map.get(key).pdf.push(file);
  });
  state.groups = [...map.values()].map(group => {
    const ready = (group.excel.length === 1 && group.word.length === 1) || group.pdf.length >= 2;
    return { ...group, ready };
  }).sort((a, b) => a.key.localeCompare(b.key));
  renderSelection();
}

function renderSelection() {
  const excel = state.files.filter(f => ["xlsx", "xlsm"].includes(ext(f.name))).length;
  const word = state.files.filter(f => ext(f.name) === "docx").length;
  $("selectionSummary").hidden = state.files.length === 0;
  $("totalFiles").textContent = state.files.length;
  $("excelFiles").textContent = excel;
  $("wordFiles").textContent = word;
  const ready = state.groups.filter(g => g.ready);
  $("pairCount").textContent = `${ready.length} ${ready.length === 1 ? "par" : "pares"}`;
  const list = $("pairList");
  if (!state.groups.length) {
    list.className = "pair-list empty";
    list.innerHTML = '<span class="pair-empty-icon" aria-hidden="true"><i></i><i></i></span><strong>Os pares aparecerão aqui</strong><p>Pequenas diferenças no nome, como “_” no início, espaços, GI/G1 e 5135/5950, são normalizadas.</p>';
  } else {
    list.className = "pair-list";
    list.innerHTML = state.groups.map((g, i) => {
      const names = [...g.excel, ...g.word, ...g.pdf].map(f => f.name).join(" + ");
      const label = g.ready ? "Pronto" : "Incompleto";
      return `<article class="pair-row"><span class="pair-index">${String(i + 1).padStart(2, "0")}</span><span class="pair-copy"><strong>${escapeHtml(g.key)}</strong><span title="${escapeHtml(names)}">${escapeHtml(names)}</span></span><span class="pair-state ${g.ready ? "" : "warning"}">${label}</span></article>`;
    }).join("");
  }
  $("processButton").disabled = ready.length === 0;
  if (!state.files.length) setStatus("Adicione os documentos para começar.");
  else if (ready.length) setStatus(`${ready.length} correlação(ões) pronta(s). Ordem aplicada: Excel primeiro, Word depois.`, "success");
  else setStatus("Nenhum par completo foi encontrado. Confira os nomes dos arquivos.", "warning");
}

function escapeHtml(value) { const div = document.createElement("div"); div.textContent = value; return div.innerHTML; }
function addFiles(fileList) {
  const incoming = [...fileList].filter(file => accepted.has(ext(file.name)));
  const existing = new Set(state.files.map(f => `${f.name}|${f.size}|${f.lastModified}`));
  incoming.forEach(file => { const id = `${file.name}|${file.size}|${file.lastModified}`; if (!existing.has(id)) { state.files.push(file); existing.add(id); } });
  analyse();
}

async function excelToPdf(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellText: true });
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  workbook.SheetNames.forEach((sheetName, index) => {
    if (index) doc.addPage("a4", "landscape");
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, dateNF: "dd/mm/yyyy", defval: "" });
    doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.text(sheetName, 10, 10);
    if (rows.length) doc.autoTable({ startY: 14, head: [rows[0]], body: rows.slice(1), theme: "grid", styles: { fontSize: 6, cellPadding: 1.2, overflow: "linebreak" }, headStyles: { fillColor: [38, 50, 56], textColor: 255 }, alternateRowStyles: { fillColor: [247, 249, 248] }, margin: 8, rowPageBreak: "avoid" });
  });
  return doc.output("arraybuffer");
}

async function wordToPdf(file) {
  const render = $("wordRender");
  const result = await mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() }, { convertImage: mammoth.images.imgElement(image => image.read("base64").then(data => ({ src: `data:${image.contentType};base64,${data}` }))) });
  render.innerHTML = result.value || "<p>Documento sem conteúdo legível.</p>";
  const doc = new jsPDF({ unit: "pt", format: "a4", orientation: "portrait" });
  await new Promise((resolve, reject) => doc.html(render, { callback: () => resolve(), x: 28, y: 28, width: 539, windowWidth: 794, autoPaging: "text", html2canvas: { scale: .75, useCORS: true, backgroundColor: "#ffffff" } }).catch?.(reject));
  render.innerHTML = "";
  return doc.output("arraybuffer");
}

async function toPdf(file) {
  const type = ext(file.name);
  if (type === "pdf") return file.arrayBuffer();
  if (["xlsx", "xlsm"].includes(type)) return excelToPdf(file);
  if (type === "docx") return wordToPdf(file);
  throw new Error(`Formato não suportado: ${file.name}`);
}

async function mergePdfs(buffers) {
  const out = await PDFDocument.create();
  for (const buffer of buffers) {
    const source = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const pages = await out.copyPages(source, source.getPageIndices());
    pages.forEach(page => out.addPage(page));
  }
  return out.save();
}

async function processAll() {
  const ready = state.groups.filter(g => g.ready);
  $("processButton").disabled = true; $("processButton").classList.add("is-loading"); $("processLabel").textContent = "Processando documentos";
  state.results.forEach(r => URL.revokeObjectURL(r.url)); state.results = [];
  try {
    for (let i = 0; i < ready.length; i++) {
      const group = ready[i]; setStatus(`Processando ${i + 1} de ${ready.length}: ${group.key}`, "info");
      // Regra fixa de composição: o conteúdo do Excel sempre abre o PDF.
      // Para PDFs já convertidos, os sufixos _Excel e _Word definem a ordem.
      const sources = group.pdf.length >= 2
        ? [...group.pdf].sort((a, b) => pdfOrder(a) - pdfOrder(b) || a.name.localeCompare(b.name))
        : [group.excel[0], group.word[0]];
      const buffers = [];
      for (const source of sources) buffers.push(await toPdf(source));
      const bytes = await mergePdfs(buffers);
      const base = group.excel[0] ? stem(group.excel[0].name) : stem(sources[0].name).replace(/_(Excel|Word)$/i, "");
      const blob = new Blob([bytes], { type: "application/pdf" });
      state.results.push({ name: `${base}.pdf`, blob, url: URL.createObjectURL(blob), size: blob.size });
    }
    renderResults(); setStatus(`${state.results.length} PDF(s) unificado(s) com sucesso.`, "success");
  } catch (error) {
    console.error(error); setStatus(`Não foi possível concluir: ${error.message}`, "error");
  } finally {
    $("processButton").disabled = false; $("processButton").classList.remove("is-loading"); $("processLabel").textContent = "Converter e unificar PDFs";
  }
}

function renderResults() {
  $("resultsSection").hidden = !state.results.length;
  $("resultSummary").textContent = `${state.results.length} arquivo(s) gerado(s), pronto(s) para download.`;
  $("resultList").innerHTML = state.results.map((r, i) => `<article class="result-row"><span class="pdf-badge">PDF</span><span class="result-copy"><strong>${escapeHtml(r.name)}</strong><span>${prettySize(r.size)}</span></span><button class="download-button" data-index="${i}">Baixar PDF</button></article>`).join("");
  $("resultList").querySelectorAll("[data-index]").forEach(button => button.addEventListener("click", () => downloadResult(state.results[Number(button.dataset.index)])));
  $("resultsSection").scrollIntoView({ behavior: "smooth", block: "start" });
}
function downloadResult(result) { const link = document.createElement("a"); link.href = result.url; link.download = result.name; link.click(); }
async function downloadZip() { const zip = new JSZip(); state.results.forEach(r => zip.file(r.name, r.blob)); const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } }); const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `PDFS_UNIFICADOS_${state.results.length}_ARQUIVOS.zip`; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }

const dropZone = $("dropZone");
dropZone.addEventListener("click", () => $("fileInput").click());
dropZone.addEventListener("keydown", e => { if (["Enter", " "].includes(e.key)) { e.preventDefault(); $("fileInput").click(); } });
$("fileInput").addEventListener("change", e => { addFiles(e.target.files); e.target.value = ""; });
["dragenter", "dragover"].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.add("is-dragging"); }));
["dragleave", "drop"].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.remove("is-dragging"); }));
dropZone.addEventListener("drop", e => addFiles(e.dataTransfer.files));
$("clearButton").addEventListener("click", () => { state.files = []; state.groups = []; state.results.forEach(r => URL.revokeObjectURL(r.url)); state.results = []; $("resultsSection").hidden = true; renderSelection(); });
$("processButton").addEventListener("click", processAll);
$("downloadAllButton").addEventListener("click", downloadZip);
$("privacyButton").addEventListener("click", () => $("privacyDialog").showModal());
$("closeDialog").addEventListener("click", () => $("privacyDialog").close());
renderSelection();
