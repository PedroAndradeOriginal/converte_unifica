const { PDFDocument } = window.PDFLib;
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
  if (data.type === "result") { pending.delete(data.requestId); job.resolve(new Uint8Array(data.data)); }
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
async function convertToPdf(file) {
  await waitForEngine();
  const requestId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
  const buffer = await file.arrayBuffer();
  return new Promise((resolve, reject) => {
    pending.set(requestId, { resolve, reject });
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
