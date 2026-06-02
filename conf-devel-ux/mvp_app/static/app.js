const modeButtons = document.querySelectorAll("[data-mode]");
const uploadPanel = document.getElementById("uploadPanel");
const folderPanel = document.getElementById("folderPanel");
const zipPanel = document.getElementById("zipPanel");
const processUploadButton = document.getElementById("processUploadButton");
const processFolderButton = document.getElementById("processFolderButton");
const processZipButton = document.getElementById("processZipButton");
const statusPill = document.getElementById("statusPill");
const excelLink = document.getElementById("excelLink");
const summaryOrigin = document.getElementById("summaryOrigin");
const resultsSection = document.getElementById("resultsSection");
const kpis = document.getElementById("kpis");
const divergenceTable = document.getElementById("divergenceTable");
const conciliatedTable = document.getElementById("conciliatedTable");
const divergenceCount = document.getElementById("divergenceCount");
const conciliatedCount = document.getElementById("conciliatedCount");
const paymentReportTable = document.getElementById("paymentReportTable");
const transferReportTable = document.getElementById("transferReportTable");
const paymentReportCount = document.getElementById("paymentReportCount");
const transferReportCount = document.getElementById("transferReportCount");
const systemVersion = document.getElementById("systemVersion");
const lastConferenceDate = document.getElementById("lastConferenceDate");
const sectionNav = document.getElementById("sectionNav");

function setStatus(text, tone = "default") {
  statusPill.textContent = text;
  statusPill.classList.toggle("processing", tone === "processing");
  if (tone === "error") { statusPill.style.background = "rgba(159,57,57,.12)"; statusPill.style.color = "#9f3939"; return; }
  statusPill.style.background = "rgba(42,106,67,.1)"; statusPill.style.color = "#2a6a43";
}
function setBusy(busy) {
  [processUploadButton, processFolderButton, processZipButton].forEach(b => b.disabled = busy);
  if (busy) setStatus("Processando... aguarde a leitura dos PDFs e da planilha.", "processing");
}
function formatMoney(v){ if(v===null||v===undefined||v==="")return "-"; return Number(v).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2}); }
function activateMode(mode){
  modeButtons.forEach(b=>b.classList.toggle("active",b.dataset.mode===mode));
  uploadPanel.classList.toggle("active",mode==="upload"); folderPanel.classList.toggle("active",mode==="folder"); zipPanel.classList.toggle("active",mode==="zip");
  processUploadButton.classList.toggle("hidden",mode!=="upload"); processFolderButton.classList.toggle("hidden",mode!=="folder"); processZipButton.classList.toggle("hidden",mode!=="zip");
}
modeButtons.forEach(b=>b.addEventListener("click",()=>activateMode(b.dataset.mode)));
[["movimentoFiles","movimentoNames"],["folderInput","folderNames"],["zipInput","zipName"]].forEach(([i,l])=>{
  const input=document.getElementById(i), label=document.getElementById(l);
  input.addEventListener("change",()=>{const names=Array.from(input.files||[]).map(f=>f.name); label.textContent=names.length?names.slice(0,12).join(", ")+(names.length>12?` e mais ${names.length-12}`:""):"Nenhum arquivo selecionado";});
});
function renderKpis(s){
  const items=[
    ['PDFs',s.pdf_count,'pdf'],
    ['Ofícios Lidos',s.oficio_count,'doc'],
    ['Relatórios',s.report_total_count,'chart'],
    ['Conciliados',s.conciliated_count,'check'],
    ['Divergências',s.divergence_count,'warn']
  ];
  kpis.innerHTML=items.map(([label,value,kind])=>`<div class="kpi kpi-${kind}"><span class="kpi-icon">${iconFor(kind)}</span><span class="kpi-label">${label}</span><span class="kpi-value">${value}</span></div>`).join("");
}
function iconFor(kind){
  return {pdf:'PDF',doc:'▤',chart:'▥',check:'✓',warn:'!'}[kind] || '';
}
function renderSectionNav(payload){
  if(!sectionNav) return;
  const cards = [
    ['Divergências', 'secDivergencias', 'warn'],
    ['Itens Conciliados', 'secConciliados', 'list'],
    ['Relatório Bancário', 'secRelatorioBancario', 'doc'],
    ['Transferências', 'secTransferencias', 'swap']
  ];
  sectionNav.innerHTML = cards.map(([label,target,kind])=>`<button type="button" class="section-card section-${kind}" data-target="${target}"><span class="section-card-icon">${navIcon(kind)}</span><strong>${label}</strong></button>`).join("");
  sectionNav.querySelectorAll("[data-target]").forEach(btn=>btn.addEventListener("click",()=>document.getElementById(btn.dataset.target)?.scrollIntoView({behavior:"smooth",block:"start"})));
}
function navIcon(kind){
  return {warn:'!',list:'☷',doc:'▤',swap:'⇄'}[kind] || '';
}
function isTotalRow(x){
  return String(x.account_debit || x.numero_oficio || "").toUpperCase().includes("TOTAL");
}
function safe(v){ return (v===null||v===undefined||v==="") ? "-" : String(v); }
function renderPaymentRows(rows){
  paymentReportCount.textContent = `${rows.length} linhas`;
  if(!rows.length){ paymentReportTable.innerHTML='<tr><td colspan="9">Nenhum relatório bancário identificado.</td></tr>'; return; }
  paymentReportTable.innerHTML = rows.map(x=>`<tr class="${isTotalRow(x)?'total-row':''}"><td>${safe(x.date)}</td><td>${safe(x.bank)}</td><td>${safe(x.account_debit)}</td><td>${safe(x.account_description)}</td><td>${safe(x.cnpj_cpf)}</td><td>${safe(x.favorecido)}</td><td>${safe(x.c_custo)}</td><td>${safe(x.conta_financeira)}</td><td>${formatMoney(x.valor)}</td></tr>`).join("");
}
function renderTransferRows(rows){
  transferReportCount.textContent = `${rows.length} linhas`;
  if(!rows.length){ transferReportTable.innerHTML='<tr><td colspan="8">Nenhum relatório de transferência identificado.</td></tr>'; return; }
  transferReportTable.innerHTML = rows.map(x=>`<tr class="${isTotalRow(x)?'total-row':''}"><td>${safe(x.date)}</td><td>${safe(x.bank)}</td><td>${safe(x.account_debit)}</td><td>${safe(x.account_description)}</td><td>${safe(x.c_custo)}</td><td>${safe(x.conta_financeira)}</td><td>${safe(x.conta_corrente_destino)}</td><td>${formatMoney(x.valor)}</td></tr>`).join("");
}
function renderResults(payload){
  renderKpis(payload.summary);
  if (systemVersion && payload.summary?.system_version) systemVersion.textContent = `VERSÃO: ${payload.summary.system_version}`.toUpperCase();
  if (lastConferenceDate) lastConferenceDate.textContent = `ÚLTIMA CONFERÊNCIA: ${payload.summary?.last_conference_date || "ainda não processada"}`.toUpperCase();
  divergenceCount.textContent=`${payload.divergences.length} itens`; document.querySelector('#secDivergencias .panel-head h3').textContent = `Divergências (${payload.divergences.length})`;
  conciliatedCount.textContent=`${payload.conciliated.length} itens`; document.querySelector('#secConciliados .panel-head h3').textContent = `Itens conciliados (${payload.conciliated.length})`;
  divergenceTable.innerHTML=payload.divergences.length?payload.divergences.map(x=>`<tr><td><span class="status-tag danger">${x.status||"-"}</span></td><td>${x.numero_oficio||"-"}</td><td>${x.date||"-"}</td><td>${x.bank||"-"}</td><td>${x.account_base||"-"}</td><td>${x.account_description||"-"}</td><td>${formatMoney(x.value_oficio)}</td><td>${formatMoney(x.value_report)}</td><td>${x.reason||"-"}</td></tr>`).join(""):'<tr><td colspan="9">Nenhuma divergência encontrada.</td></tr>';
  if(payload.conciliated.length){
    const rows = payload.conciliated.map(x=>`<tr><td>${x.numero_oficio||"-"}</td><td>${x.date||"-"}</td><td>${x.bank||"-"}</td><td>${x.account_base||"-"}</td><td>${x.account_description||"-"}</td><td>${formatMoney(x.value_oficio)}</td><td>${formatMoney(x.value_report)}</td></tr>`).join("");
    const totalOficio = payload.summary.conciliated_value_oficio_total ?? payload.conciliated.reduce((s,x)=>s+Number(x.value_oficio||0),0);
    const totalRelatorio = payload.summary.conciliated_value_report_total ?? payload.conciliated.reduce((s,x)=>s+Number(x.value_report||0),0);
    conciliatedTable.innerHTML = rows + `<tr class="total-row"><td colspan="5"><strong>TOTAL GERAL</strong></td><td><strong>${formatMoney(totalOficio)}</strong></td><td><strong>${formatMoney(totalRelatorio)}</strong></td></tr>`;
  } else {
    conciliatedTable.innerHTML='<tr><td colspan="7">Nenhum item conciliado.</td></tr>';
  }
  renderSectionNav(payload);
  renderPaymentRows(payload.payment_report_rows || []);
  renderTransferRows(payload.transfer_report_rows || []);
  if(payload.excel_url){ excelLink.href=payload.excel_url; excelLink.classList.remove("hidden"); }
  summaryOrigin.textContent = payload.folder_upload ? "Origem: pasta enviada pelo navegador." : payload.zip_file ? `Origem: ZIP ${payload.zip_file}.` : "Origem: upload do movimento bancário.";
  if(payload.ignored_files && payload.ignored_files.length) summaryOrigin.textContent += ` Ignorados: ${payload.ignored_files.join(", ")}.`;
  resultsSection.classList.add("visible");
}
async function sendFiles(url, files, fieldName){
  if(!files.length){ setStatus("Selecione os arquivos antes de processar.","error"); return; }
  const fd=new FormData(); Array.from(files).forEach(f=>fd.append(fieldName,f,f.webkitRelativePath||f.name));
  setBusy(true);
  try{
    const r=await fetch(url,{method:"POST",body:fd});
    const text=await r.text();
    let p;
    try { p=JSON.parse(text); } catch (_) { throw new Error("O servidor retornou uma página HTML em vez de JSON. Verifique o log do Render."); }
    if(!r.ok) throw new Error(p.detail||"Falha ao processar.");
    renderResults(p); setStatus("Processamento concluído.");
  }
  catch(e){ setStatus(e.message,"error"); }
  finally{ setBusy(false); }
}
processUploadButton.addEventListener("click",()=>sendFiles("/api/process-movimento",document.getElementById("movimentoFiles").files,"movimento_files"));
processFolderButton.addEventListener("click",()=>sendFiles("/api/process-folder",document.getElementById("folderInput").files,"folder_files"));
processZipButton.addEventListener("click",()=>sendFiles("/api/process-zip",document.getElementById("zipInput").files,"zip_file"));

async function loadPersistedState(){
  try{
    const r = await fetch("/api/state", {cache:"no-store"});
    if(!r.ok) return;
    const state = await r.json();
    if(state.last_conference_date && lastConferenceDate){
      lastConferenceDate.textContent = `ÚLTIMA CONFERÊNCIA: ${state.last_conference_date}`.toUpperCase();
    }
  } catch (_) {}
}
loadPersistedState();
