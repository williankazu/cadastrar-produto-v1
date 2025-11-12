
/* ===========================
   ESTADO & UTIL
=========================== */
const $ = (id)=>document.getElementById(id);
const STORAGE_KEY = "produtos_markup_margem_pv_v12_fix_center_fit";
const LOGO_KEY = "logo_dataurl_v1";
const EAN_PREFIX_KEY = "ean_prefix_last";

let produtos = [];
let editIndex = null;
let metodoAtual = "markup"; // "markup" | "margem" | "pv"
let sortKey = "descricao";
let sortDir = 1;
let filtroBusca = "";
let filtroCategoria = "";
let logoDataUrl = null;

/* ========= CLIPBOARD ========= */
async function copyToClipboard(texto){
  try{
    await navigator.clipboard.writeText(texto);
    toast("Copiado!");
  }catch(e){
    const ta = document.createElement("textarea");
    ta.value = texto; document.body.appendChild(ta);
    ta.select(); document.execCommand("copy");
    ta.remove(); toast("Copiado!");
  }
}
function toast(msg){
  const n = document.createElement("div");
  n.textContent = msg;
  n.style.position="fixed"; n.style.right="16px"; n.style.bottom="16px";
  n.style.background="rgba(0,0,0,.8)"; n.style.color="#fff"; n.style.padding="10px 14px";
  n.style.borderRadius="10px"; n.style.zIndex=9999; n.style.fontSize="14px";
  document.body.appendChild(n);
  setTimeout(()=>n.remove(), 1400);
}

/* ===========================
   INIT
=========================== */
(function init(){
  const saved = localStorage.getItem(STORAGE_KEY);
  if(saved){ try{ produtos = JSON.parse(saved)||[]; }catch(e){} }

  logoDataUrl = localStorage.getItem(LOGO_KEY) || null;
  if(logoDataUrl){ $("logoPreview").src = logoDataUrl; }

  const lastPrefix = localStorage.getItem(EAN_PREFIX_KEY);
  if(lastPrefix && $("eanPrefix")) $("eanPrefix").value = lastPrefix;

  marcarMetodo(metodoAtual);
  bindEvents();
  renderTabela();
  atualizarPreview();
})();

function bindEvents(){
  // método tags
  $("tagMethodMarkup").addEventListener("click", ()=>{ metodoAtual="markup"; marcarMetodo(metodoAtual); atualizarPreview(); });
  $("tagMethodMargem").addEventListener("click", ()=>{ metodoAtual="margem"; marcarMetodo(metodoAtual); atualizarPreview(); });
  $("tagMethodPV").addEventListener("click", ()=>{ metodoAtual="pv"; marcarMetodo(metodoAtual); atualizarPreview(); });

  // entradas que recalculam preview
  ["preco","icms","pis","cofins","st","markup","margem","precoVendaInput","quantidade"].forEach(id=>{
    $(id).addEventListener("input", atualizarPreview);
  });

  ["icms","pis","cofins","st"].forEach(id=>{
    $(id).addEventListener("input", somarTributos);
  });

  // EAN (entrada manual)
  $("ean").addEventListener("input", ()=>{
    const code = $("ean").value.replace(/\D/g,"").slice(0,13);
    $("ean").value = code;
    desenharEAN(code);
  });

  $("eanPrefix").addEventListener("change", (e)=>{
    localStorage.setItem(EAN_PREFIX_KEY, e.target.value);
  });

  $("btnGerarEAN").addEventListener("click", (e)=>{ e.preventDefault(); gerarEANeDesenhar(); });

  // COPY do formulário
  $("btnCopyDescricaoForm").addEventListener("click", (e)=>{ e.preventDefault(); const t=$("descricao").value.trim(); if(t){ copyToClipboard(t); } });
  $("btnCopyEANForm").addEventListener("click", (e)=>{ e.preventDefault(); const c=$("ean").value.trim(); if(c){ copyToClipboard(c); } });

  // CRUD
  $("btnAdicionar").addEventListener("click", adicionarProduto);
  $("btnAtualizar").addEventListener("click", atualizarProduto);
  $("btnLimpar").addEventListener("click", ()=>limparCampos(true));
  $("btnZerar").addEventListener("click", zerarTudo);

  // Export/Import/Relatório
  $("btnExportCSV").addEventListener("click", exportarCSV);
  $("btnExportXLSX").addEventListener("click", exportarXLSX);
  $("btnExportPDF").addEventListener("click", exportarPDF);
  $("fileImport").addEventListener("change", importarArquivo);

  // Etiquetas PDF padrão
  $("btnLabelsPDF").addEventListener("click", gerarEtiquetasPDF);
  // A4 (20 itens)
  $("btnA4Simple").addEventListener("click", imprimirA4_20itens);
  // Copiar em massa
  $("btnCopyMass").addEventListener("click", copiarSelecionadosEmMassa);
  // Excluir selecionados
  $("btnDeleteSelected").addEventListener("click", excluirSelecionados);
  // Select all
  $("selAll").addEventListener("change", toggleSelectAll);
  // Imprimir Personalizado (modal)
  $("btnPrintCustom").addEventListener("click", abrirModalImpressao);
  $("modalPrintClose").addEventListener("click", fecharModalImpressao);
  $("modalPrintCancel").addEventListener("click", fecharModalImpressao);
  $("modalPrintConfirm").addEventListener("click", confirmarImpressaoPersonalizada);

  // Filtros e ordenação
  $("filtroBusca").addEventListener("input", (e)=>{ filtroBusca = e.target.value.toLowerCase(); renderTabela(); });
  $("filtroCategoria").addEventListener("change", (e)=>{ filtroCategoria = e.target.value; renderTabela(); });
  $("btnLimparFiltros").addEventListener("click", ()=>{
    filtroBusca=""; filtroCategoria="";
    $("filtroBusca").value=""; $("filtroCategoria").value="";
    renderTabela();
  });
  document.querySelectorAll("#tabela thead th[data-sort]").forEach(th=>{
    th.addEventListener("click", ()=>{
      const key = th.getAttribute("data-sort");
      if(sortKey===key){ sortDir *= -1; } else { sortKey = key; sortDir = 1; }
      renderTabela();
    });
  });

  // Logo upload
  $("logoInput").addEventListener("change", async (e)=>{
    const file = e.target.files?.[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = ()=> {
      logoDataUrl = reader.result;
      $("logoPreview").src = logoDataUrl;
      localStorage.setItem(LOGO_KEY, logoDataUrl);
    };
    reader.readAsDataURL(file);
  });
}

/* ===========================
   EAN-13 (GS1 Brasil 789/790)
=========================== */
function gerarEAN12BR(prefix = "789"){
  const base9 = Array.from({length: 9}, () => Math.floor(Math.random()*10)).join("");
  return String(prefix).slice(0,3) + base9; // 3+9 = 12
}
function calcularDigitoEAN13(d12){
  const n = d12.split("").map(x=>+x);
  let sum = 0;
  for(let i=0;i<12;i++){
    const v = n[11 - i];
    sum += (i % 2 === 0) ? v * 3 : v;
  }
  return ((10 - (sum % 10)) % 10).toString();
}
function gerarEAN13BR(prefix = "789"){
  const base12 = gerarEAN12BR(prefix);
  return base12 + calcularDigitoEAN13(base12);
}
function validarEAN13(code){
  if(!/^\d{13}$/.test(code)) return false;
  const d12 = code.slice(0,12);
  const dv  = calcularDigitoEAN13(d12);
  return code.endsWith(dv);
}
function desenharEAN(code){
  const svg = $("eanPreview");
  if(code && /^\d{12,13}$/.test(code)){
    let c = code.length===12 ? code + calcularDigitoEAN13(code) : code;
    try{ JsBarcode(svg, c, {format:"ean13", displayValue:true, fontSize:12, width:2, height:50}); }
    catch(e){ svg.innerHTML=""; }
  } else { svg.innerHTML = ""; }
}
function gerarEANeDesenhar(){
  const prefixSel = $("eanPrefix");
  const prefix = prefixSel ? prefixSel.value : "789";
  localStorage.setItem(EAN_PREFIX_KEY, prefix);
  const code = gerarEAN13BR(prefix);
  $("ean").value = code;
  desenharEAN(code);
}

/* ===========================
   CÁLCULOS
=========================== */
function somarTributos(){
  const t = ["icms","pis","cofins","st"].map(id=>parseFloat($(id).value||"0")||0).reduce((a,b)=>a+b,0);
  $("tributos").value = t.toFixed(2);
  atualizarPreview();
}
function custoEfetivo(custo, tribPct){
  const t = Number(tribPct)||0;
  return custo * (1 + t/100);
}
function precoPorMarkup(custoEf, markup){
  const m = Number(markup)||0; return custoEf * (1 + m/100);
}
function precoPorMargem(custoEf, margem){
  const mg = Number(margem)||0;
  if(mg>=100) return Infinity;
  return custoEf / (1 - mg/100);
}
function markupFromPV(custoEf, pv){
  if(custoEf<=0) return 0;
  return (pv/custoEf - 1) * 100;
}
function margemFromPV(custoEf, pv){
  if(pv<=0) return 0;
  return ((pv - custoEf)/pv) * 100;
}
function formatBR(n){ return isFinite(n) ? n.toLocaleString("pt-BR",{style:"currency",currency:"BRL"}) : "—"; }
function escapeHTML(s){
  return (s||"").replace(/[&<>"']/g,(c)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]));
}
function salvar(){ localStorage.setItem(STORAGE_KEY, JSON.stringify(produtos)); }

/* Preview dinâmico + bloqueio */
function atualizarPreview(){
  const custo = parseFloat($("preco").value||"0");
  const trib = ["icms","pis","cofins","st"].map(id=>parseFloat($(id).value||"0")||0).reduce((a,b)=>a+b,0);
  $("tributos").value = trib.toFixed(2);
  $("prevTributos").textContent = (trib||0).toFixed(2).replace(".",",") + "%";

  const mk = parseFloat($("markup").value||"0");
  const mg = parseFloat($("margem").value||"0");
  const pvManual = parseFloat($("precoVendaInput").value||"0");

  const ce = custoEfetivo(custo, trib);
  $("prevCustoEfetivo").textContent = formatBR(ce);

  let pv = 0;
  if(metodoAtual==="markup"){ pv = precoPorMarkup(ce, mk); }
  else if(metodoAtual==="margem"){ pv = precoPorMargem(ce, mg); }
  else { pv = pvManual>0 ? pvManual : 0; }

  if(pv>0){
    const mkAuto = markupFromPV(ce, pv);
    const mgAuto = margemFromPV(ce, pv);
    if(metodoAtual==="pv"){
      $("markup").value = isFinite(mkAuto) ? mkAuto.toFixed(2) : "";
      $("margem").value = isFinite(mgAuto) ? mgAuto.toFixed(2) : "";
    }
  }
  $("precoCalculado").textContent = formatBR(pv);
}

/* ===========================
   CRUD (com bloqueio PV >= CE)
=========================== */
function marcarMetodo(metodo){
  metodoAtual = metodo;
  ["tagMethodMarkup","tagMethodMargem","tagMethodPV"].forEach(id=>$(id).classList.remove("is-active"));
  if(metodo==="markup") $("tagMethodMarkup").classList.add("is-active");
  if(metodo==="margem") $("tagMethodMargem").classList.add("is-active");
  if(metodo==="pv") $("tagMethodPV").classList.add("is-active");
}

function coletarForm(){
  const descricao = $("descricao").value.trim();
  const categoria = $("categoria").value.trim();
  const fornecedor = $("fornecedor").value.trim();
  const preco = parseFloat($("preco").value||"0");
  const icms = parseFloat($("icms").value||"0")||0;
  const pis = parseFloat($("pis").value||"0")||0;
  const cofins = parseFloat($("cofins").value||"0")||0;
  const st = parseFloat($("st").value||"0")||0;
  const tributos = +(icms+pis+cofins+st).toFixed(2);
  const unidade = $("unidade").value;
  const quantidade = Math.max(1, parseInt($("quantidade").value||"1"));

  let markup = parseFloat($("markup").value||"0");
  let margem = parseFloat($("margem").value||"0");
  let precoVendaInput = parseFloat($("precoVendaInput").value||"0");
  const produtoUrl = $("produtoUrl").value.trim();

  let ean = $("ean").value.replace(/\D/g,"");
  if(ean && ean.length===12) ean += calcularDigitoEAN13(ean);
  if(ean && !validarEAN13(ean)){ alert("EAN-13 inválido."); return null; }

  if(!descricao){ alert("Informe a descrição."); return null; }
  if(!(preco>=0)){ alert("Informe o preço de custo."); return null; }
  if(margem>=100){ alert("Margem não pode ser ≥ 100%."); return null; }

  const ce = custoEfetivo(preco, tributos);
  let metodo = metodoAtual;
  let precoVenda = 0;

  if(metodo==="markup"){
    precoVenda = precoPorMarkup(ce, markup);
  }else if(metodo==="margem"){
    precoVenda = precoPorMargem(ce, margem);
  }else{
    if(!(precoVendaInput>0)){ alert("Informe um Preço de venda válido."); return null; }
    precoVenda = precoVendaInput;
    markup = markupFromPV(ce, precoVenda);
    margem = margemFromPV(ce, precoVenda);
  }

  if(!isFinite(precoVenda) || precoVenda<=0){
    alert("Preço de venda inválido.");
    return null;
  }
  if(precoVenda < ce){
    alert("Preço de venda não pode ser menor que o CUSTO EFETIVO (custo + tributos). Ajuste os valores.");
    return null;
  }

  const subtotal = precoVenda * quantidade;
  return { descricao, categoria, fornecedor, preco, icms, pis, cofins, st, tributos, unidade, quantidade, markup, margem, metodo, ean, produtoUrl, precoVenda, subtotal };
}

function adicionarProduto(e){ e?.preventDefault?.(); const item = coletarForm(); if(!item) return;
  produtos.push(item); salvar(); renderTabela(); limparCampos(true);
}
function editarProduto(idx){
  const p = produtos[idx]; if(!p) return;
  editIndex = idx;
  $("descricao").value = p.descricao;
  $("categoria").value = p.categoria||"";
  $("fornecedor").value = p.fornecedor||"";
  $("preco").value = p.preco;
  $("icms").value = p.icms||0;
  $("pis").value = p.pis||0;
  $("cofins").value = p.cofins||0;
  $("st").value = p.st||0;
  $("tributos").value = (p.tributos||0).toFixed(2);
  $("unidade").value = p.unidade;
  $("quantidade").value = p.quantidade;
  $("markup").value = (p.markup||0).toFixed(2);
  $("margem").value = (p.margem||0).toFixed(2);
  $("precoVendaInput").value = (p.precoVenda||0).toFixed(2);
  $("produtoUrl").value = p.produtoUrl||"";
  marcarMetodo(p.metodo||"markup");
  $("ean").value = p.ean||"";
  desenharEAN($("ean").value);
  $("btnAtualizar").disabled=false; $("btnAdicionar").disabled=true;
  atualizarPreview(); window.scrollTo({top:0, behavior:"smooth"});
}
function atualizarProduto(e){ e?.preventDefault?.(); if(editIndex===null) return;
  const item = coletarForm(); if(!item) return;
  produtos[editIndex] = item; salvar(); renderTabela(); limparCampos(true);
}
function excluirProduto(idx){
  if(!confirm("Remover este item?")) return;
  produtos.splice(idx,1); salvar(); renderTabela();
}
function limparCampos(reset=false){
  ["descricao","categoria","fornecedor","preco","icms","pis","cofins","st","tributos","quantidade","markup","margem","precoVendaInput","ean","produtoUrl"].forEach(id=>$(id).value = id==="quantidade"?"1":"");
  $("eanPreview").innerHTML="";
  if(reset){ editIndex=null; $("btnAtualizar").disabled=true; $("btnAdicionar").disabled=false; }
  atualizarPreview();
}
function zerarTudo(){
  if(!confirm("Apagar TODOS os itens?")) return;
  produtos = []; salvar(); renderTabela();
}

/* ===========================
   SELEÇÃO & EXCLUSÃO EM MASSA
=========================== */
function getSelectedIndices(){
  return Array.from(document.querySelectorAll(".selRow:checked")).map(c=>+c.dataset.idx);
}
function toggleSelectAll(e){
  const check = e.target.checked;
  document.querySelectorAll(".selRow").forEach(cb=>{ cb.checked = check; });
}
function excluirSelecionados(){
  const idxs = getSelectedIndices();
  if(!idxs.length){ alert("Nenhum item selecionado."); return; }
  if(!confirm(`Excluir ${idxs.length} item(ns) selecionado(s)?`)) return;
  idxs.sort((a,b)=>b-a).forEach(i=>produtos.splice(i,1));
  salvar(); renderTabela();
  $("selAll").checked = false;
}

/* ===========================
   RENDER + FILTRO + SORT + TOTAIS
=========================== */
function renderTabela(){
  let list = produtos.slice().filter(p=>{
    const matchBusca = !filtroBusca ||
      (p.descricao||"").toLowerCase().includes(filtroBusca) ||
      (p.categoria||"").toLowerCase().includes(filtroBusca) ||
      (p.fornecedor||"").toLowerCase().includes(filtroBusca);
    const matchCat = !filtroCategoria || (p.categoria||"")===filtroCategoria;
    return matchBusca && matchCat;
  });

  list.sort((a,b)=>{
    const va = (a[sortKey]??""); const vb = (b[sortKey]??"");
    if(typeof va==="number" && typeof vb==="number") return (va - vb)*sortDir;
    return String(va).localeCompare(String(vb),"pt-BR")*sortDir;
  });

  let totalCusto=0, totalVenda=0, itens=0;
  const tbody = $("tbody");
  tbody.innerHTML = "";

  list.forEach((p, iInList)=>{
    const idx = produtos.indexOf(p);
    const ce = custoEfetivo(p.preco, p.tributos||0);
    totalCusto += ce * (p.quantidade||1);
    totalVenda += (p.precoVenda||0) * (p.quantidade||1);
    itens += p.quantidade||1;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="has-text-centered">
        <input type="checkbox" class="selRow" data-idx="${idx}">
      </td>
      <td>${escapeHTML(p.descricao)}</td>
      <td>${escapeHTML(p.categoria||"")}</td>
      <td>${escapeHTML(p.fornecedor||"")}</td>
      <td>${escapeHTML(p.unidade)}</td>
      <td class="has-text-right">${p.quantidade}</td>
      <td class="has-text-right">${formatBR(p.preco)}</td>
      <td class="has-text-right">${(p.tributos||0).toFixed(2)} %</td>
      <td class="has-text-centered">
        <span class="tag ${p.metodo==='markup'?'is-info':(p.metodo==='margem'?'is-link':'is-primary')} is-light">
          ${p.metodo==='markup'?'Markup':(p.metodo==='margem'?'Margem':'Preço de venda')}
        </span>
      </td>
      <td class="has-text-right">${(p.markup||0).toFixed(2)} %</td>
      <td class="has-text-right">${(p.margem||0).toFixed(2)} %</td>
      <td class="has-text-right">${formatBR(p.precoVenda)}</td>
      <td class="has-text-right">${formatBR((p.precoVenda||0) * (p.quantidade||1))}</td>
      <td>${p.ean ? `<small>${p.ean}</small><br><svg id="eanRow_${idx}" class="barcode-canvas"></svg>` : "<span class='has-text-grey'>—</span>"}</td>
      <td>${p.produtoUrl ? `<a href="${p.produtoUrl}" target="_blank">Abrir</a>` : "<span class='has-text-grey'>—</span>"}</td>
      <td class="has-text-centered">
        <div class="buttons are-small">
          <button class="button is-link is-light" title="Editar" onclick="editarProduto(${idx})"><i class="fa-solid fa-pen"></i></button>
          <button class="button is-danger is-light" title="Excluir" onclick="excluirProduto(${idx})"><i class="fa-solid fa-trash"></i></button>
          <button class="button is-light" title="Copiar descrição" onclick="copiarDescricaoLinha(${idx})"><i class="fa-regular fa-copy"></i> DESC</button>
          <button class="button is-light" title="Copiar EAN-13" onclick="copiarEANLinha(${idx})"><i class="fa-regular fa-copy"></i> EAN</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    if(p.ean){ try{ JsBarcode(`#eanRow_${idx}`, p.ean, {format:"ean13", displayValue:false, width:2, height:40}); }catch(e){} }
  });

  $("totalSem").textContent = formatBR(totalCusto);
  $("totalCom").textContent = formatBR(totalVenda);
  $("totalItens").textContent = (itens||0).toLocaleString("pt-BR");

  const cats = Array.from(new Set(produtos.map(p=>p.categoria||"").filter(Boolean))).sort((a,b)=>a.localeCompare(b,"pt-BR"));
  const sel = $("filtroCategoria");
  const cur = sel.value;
  sel.innerHTML = `<option value="">Todas</option>` + cats.map(c=>`<option ${c===cur?'selected':''}>${escapeHTML(c)}</option>`).join("");

  salvar();
}

/* ===========================
   EXPORT / IMPORT / RELATÓRIO
=========================== */
function exportarCSV(){
  const rows = [["descricao","categoria","fornecedor","unidade","quantidade","preco_custo","icms_percent","pis_percent","cofins_percent","st_percent","tributos_percent","metodo","markup_percent","margem_percent","preco_venda","subtotal","ean13","url"]];
  produtos.forEach(p=>{
    rows.push([p.descricao,p.categoria||"",p.fornecedor||"",p.unidade,p.quantidade,p.preco,p.icms||0,p.pis||0,p.cofins||0,p.st||0,p.tributos||0,p.metodo,p.markup||0,p.margem||0,p.precoVenda||0,(p.precoVenda||0)*(p.quantidade||1),p.ean||"",p.produtoUrl||""]);
  });
  const csv = rows.map(r=>r.map(v=>String(v).includes(",")?`"${String(v).replace(/"/g,'""')}"`:v).join(",")).join("\n");
  baixarArquivo("produtos.csv","text/csv;charset=utf-8;",csv);
}
function exportarXLSX(){
  const data = produtos.map(p=>({
    descricao:p.descricao, categoria:p.categoria||"", fornecedor:p.fornecedor||"", unidade:p.unidade, quantidade:p.quantidade,
    preco_custo:p.preco, icms_percent:p.icms||0, pis_percent:p.pis||0, cofins_percent:p.cofins||0, st_percent:p.st||0,
    tributos_percent:p.tributos||0, metodo:p.metodo, markup_percent:p.markup||0, margem_percent:p.margem||0,
    preco_venda:p.precoVenda||0, subtotal:(p.precoVenda||0)*(p.quantidade||1), ean13:p.ean||"", url:p.produtoUrl||""
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Produtos");
  XLSX.writeFile(wb, "produtos.xlsx");
}
function exportarPDF(){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:"landscape"});
  doc.setFontSize(14); doc.text("Relatório de Produtos", 14, 16);
  const body = produtos.map((p,i)=>[
    i+1, p.descricao, p.categoria||"", p.fornecedor||"", p.unidade, p.quantidade,
    toBR(p.preco), (p.icms||0)+"%", (p.pis||0)+"%", (p.cofins||0)+"%", (p.st||0)+"%", (p.tributos||0)+"%",
    p.metodo, (p.markup||0).toFixed(2)+"%", (p.margem||0).toFixed(2)+"%",
    toBR(p.precoVenda||0), toBR((p.precoVenda||0)*(p.quantidade||1)), p.ean||"", p.produtoUrl||""
  ]);
  doc.autoTable({
    startY:22,
    head:[["#","Descrição","Categoria","Fornecedor","Unid.","Qtde","Custo","ICMS","PIS","COFINS","ST","Trib.","Mét.","Markup","Margem","PV","Subtotal","EAN-13","URL"]],
    body, styles:{fontSize:8}, headStyles:{fillColor:[102,126,234]}
  });
  const totalCusto = produtos.reduce((a,p)=> a + custoEfetivo(p.preco, p.tributos||0)*(p.quantidade||1), 0);
  const totalVenda = produtos.reduce((a,p)=> a + (p.precoVenda||0)*(p.quantidade||1), 0);
  doc.setFontSize(12);
  doc.text(`Total CUSTO (c/ tributos): ${toBR(totalCusto)}`, 14, doc.lastAutoTable.finalY + 10);
  doc.text(`Total VENDA: ${toBR(totalVenda)}`, 14, doc.lastAutoTable.finalY + 18);
  doc.save("produtos.pdf");
}
function toBR(n){ return (n||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"}); }
function baixarArquivo(nome, mime, conteudo){
  const blob = new Blob([conteudo], {type:mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href=url; a.download=nome; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
}
function importarArquivo(evt){
  const file = evt.target.files[0]; if(!file) return;
  const ext = (file.name.split(".").pop()||"").toLowerCase();
  const reader = new FileReader();
  reader.onload = function(e){
    if(ext==="csv"){ importarCSV(e.target.result); }
    else{ const data = new Uint8Array(e.target.result); const wb = XLSX.read(data,{type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]]; const json = XLSX.utils.sheet_to_json(ws,{defval:""}); importarJSON(json); }
    $("fileImport").value="";
  };
  if(ext==="csv") reader.readAsText(file,"utf-8"); else reader.readAsArrayBuffer(file);
}
function parseCSVRow(row){
  const out=[]; let cur=""; let inQ=false;
  for(let i=0;i<row.length;i++){
    const ch=row[i];
    if(ch==='"'){ if(inQ && row[i+1]==='"'){ cur+='"'; i++; } else inQ=!inQ; }
    else if(ch===',' && !inQ){ out.push(cur); cur=""; }
    else cur+=ch;
  } out.push(cur); return out.map(s=>s.trim());
}
function importarCSV(text){
  const linhas = text.split(/\r?\n/).filter(l=>l.trim().length); if(linhas.length<=1){ alert("CSV vazio ou inválido."); return; }
  const header = parseCSVRow(linhas[0]); const idx = Object.fromEntries(header.map((h,i)=>[h,i]));
  for(let i=1;i<linhas.length;i++){ const cols = parseCSVRow(linhas[i]); addFromCols(cols, idx); }
  salvar(); renderTabela();
}
function addFromCols(cols, idx){
  const get = (k)=> cols[idx[k]] ?? "";
  const descricao = get("descricao"); if(!descricao) return;
  const categoria = get("categoria")||"";
  const fornecedor = get("fornecedor")||"";
  const unidade = get("unidade")||"un";
  const quantidade = parseInt(get("quantidade")||"1")||1;
  const preco = parseFloat(get("preco_custo")||"0")||0;
  const icms = parseFloat(get("icms_percent")||"0")||0;
  const pis = parseFloat(get("pis_percent")||"0")||0;
  const cofins = parseFloat(get("cofins_percent")||"0")||0;
  const st = parseFloat(get("st_percent")||"0")||0;
  const tributos = +(icms+pis+cofins+st).toFixed(2);
  const metodo = (get("metodo")||"markup").toLowerCase();
  let markup = parseFloat(get("markup_percent")||"0")||0;
  let margem = parseFloat(get("margem_percent")||"0")||0;
  let precoVenda = parseFloat(get("preco_venda")||"0")||0;
  let ean = (get("ean13")||"").replace(/\D/g,"");
  const produtoUrl = get("url")||"";
  if(ean && ean.length===12) ean += calcularDigitoEAN13(ean);
  if(ean && !validarEAN13(ean)) ean="";
  const ce = custoEfetivo(preco, tributos);
  let metodoFinal = (metodo==="margem"?"margem":(metodo==="pv"?"pv":"markup"));
  if(metodoFinal==="markup"){ precoVenda = precoPorMarkup(ce, markup); margem = margemFromPV(ce, precoVenda); }
  else if(metodoFinal==="margem"){ precoVenda = precoPorMargem(ce, margem); markup = markupFromPV(ce, precoVenda); }
  else{ if(!(precoVenda>0)) return; markup = markupFromPV(ce, precoVenda); margem = margemFromPV(ce, precoVenda); }
  if(precoVenda<ce) return; // respeita bloqueio
  const subtotal = precoVenda*quantidade;
  produtos.push({descricao,categoria,fornecedor,unidade,quantidade,preco,icms,pis,cofins,st,tributos,metodo:metodoFinal,markup,margem,precoVenda,subtotal,ean,produtoUrl});
}
function importarJSON(json){
  json.forEach(o=>{
    const r = Object.fromEntries(Object.entries(o).map(([k,v])=>[String(k).toLowerCase(), v]));
    const descricao = (r.descricao||r.nome||"").toString().trim(); if(!descricao) return;
    const categoria = (r.categoria||"").toString();
    const fornecedor = (r.fornecedor||"").toString();
    const unidade = (r.unidade||"un").toString();
    const quantidade = parseInt(r.quantidade||1)||1;
    const preco = parseFloat(r.preco_custo||r.custo||r.preco||0)||0;
    const icms = parseFloat(r.icms_percent||r.icms||0)||0;
    const pis = parseFloat(r.pis_percent||r.pis||0)||0;
    const cofins = parseFloat(r.cofins_percent||r.cofins||0)||0;
    const st = parseFloat(r.st_percent||r.st||0)||0;
    const tributos = +(icms+pis+cofins+st).toFixed(2);
    const metodoStr = (r.metodo||"markup").toString().toLowerCase();
    const metodo = (metodoStr==="margem"?"margem":(metodoStr==="pv"?"pv":"markup"));
    let markup = parseFloat(r.markup_percent||r.markup||0)||0;
    let margem = parseFloat(r.margem_percent||r.margem||0)||0;
    let precoVenda = parseFloat(r.preco_venda||0)||0;
    let ean = (r.ean13||r.ean||"").toString().replace(/\D/g,"");
    const produtoUrl = (r.url||"").toString();
    if(ean && ean.length===12) ean += calcularDigitoEAN13(ean);
    if(ean && !validarEAN13(ean)) ean="";
    const ce = custoEfetivo(preco, tributos);
    if(metodo==="markup"){ precoVenda = precoPorMarkup(ce, markup); margem = margemFromPV(ce, precoVenda); }
    else if(metodo==="margem"){ precoVenda = precoPorMargem(ce, margem); markup = markupFromPV(ce, precoVenda); }
    else { if(!(precoVenda>0)) return; markup = markupFromPV(ce, precoVenda); margem = margemFromPV(ce, precoVenda); }
    if(precoVenda<ce) return;
    const subtotal = precoVenda*quantidade;
    produtos.push({descricao,categoria,fornecedor,unidade,quantidade,preco,icms,pis,cofins,st,tributos,metodo,markup,margem,precoVenda,subtotal,ean,produtoUrl});
  });
  salvar(); renderTabela();
}

/* ===========================
   ETIQUETAS PDF padrão
=========================== */
function dataHojeBR(){
  const d = new Date();
  return d.toLocaleDateString("pt-BR");
}
function gerarBarcodeDataURL(code, pxWidth=280, pxHeight=90){
  const canvas = document.createElement("canvas");
  try {
    JsBarcode(canvas, code, {format:"ean13", displayValue:false, width:2, height:Math.max(40, Math.floor(pxHeight/2))});
    return canvas.toDataURL("image/png");
  } catch(e){
    return null;
  }
}
function gerarQRDataURL(text, size=100){
  const qr = new QRious({ value: text || "", size });
  return qr.toDataURL("image/png");
}
function itensSelecionadosOuTodos(){
  const selIdx = getSelectedIndices();
  return selIdx.length ? selIdx.map(i=>produtos[i]) : produtos.slice();
}
function gerarEtiquetasPDF(){
  const items = itensSelecionadosOuTodos();
  if(!items.length){ alert("Nenhum item para etiquetas."); return; }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:"portrait", unit:"mm", format:"a4"});

  const margin = 10, gap = 6, cols = 3;
  const pageW = 210, pageH = 297;
  const labelW = (pageW - margin*2 - gap*(cols-1)) / cols;
  const labelH = 38;
  const pad = 3; // padding interno

  let x = margin, y = margin, col=0;

  items.forEach((p)=>{
    if(y + labelH > pageH - margin){
      doc.addPage();
      x = margin; y = margin; col=0;
    }

    doc.setDrawColor(200); doc.setLineWidth(0.2);
    doc.rect(x, y, labelW, labelH);

    const cx = x + labelW/2;
    const innerW = labelW - 2*pad;
    let cursorY = y + pad;

    if(logoDataUrl){
      const maxLW = Math.min(18, innerW);
      const maxLH = 8;
      doc.addImage(logoDataUrl, "PNG", cx - (maxLW/2), cursorY, maxLW, maxLH, undefined, "FAST");
      cursorY += maxLH + 1.5;
    }

    doc.setFontSize(8); doc.setTextColor(80);
    doc.text(`Data: ${dataHojeBR()}`, cx, cursorY, {align:"center"});
    cursorY += 4;

    doc.setFontSize(10); doc.setTextColor(0);
    const desc = (p.descricao||"").toString();
    const descLines = doc.splitTextToSize(desc, innerW);
    const linesPrinted = Math.min(2, descLines.length);
    doc.text(descLines.slice(0,linesPrinted), cx, cursorY, {align:"center"});
    cursorY += linesPrinted*4 + 1;

    doc.setFontSize(8); doc.setTextColor(60);
    doc.text(`Unid: ${p.unidade}`, cx, cursorY, {align:"center"});
    cursorY += 5;

    doc.setFontSize(12); doc.setTextColor(0);
    doc.text(`Preço: ${toBR(p.precoVenda||0)}`, cx, cursorY, {align:"center"});
    cursorY += 6;

    // Espaço restante dentro da etiqueta
    let remainingH = y + labelH - pad - cursorY;

    // Desenha barcode dentro do espaço restante
    if(p.ean && validarEAN13(p.ean) && remainingH > 6){
      const barData = gerarBarcodeDataURL(p.ean, 280, 90);
      if(barData){
        const barW = innerW;                   // ocupa a largura útil
        const barH = Math.max(8, Math.min(12, remainingH - 2)); // nunca passa da borda
        doc.addImage(barData, "PNG", cx - (barW/2), cursorY, barW, barH, undefined, "FAST");
        cursorY += barH + 1.5;
        remainingH = y + labelH - pad - cursorY;
      }
    }

    // QR dentro do espaço restante (se houver URL)
    if(p.produtoUrl && remainingH > 8){
      const qrSize = Math.max(10, Math.min(14, remainingH)); // cabe no resto
      doc.addImage(gerarQRDataURL(p.produtoUrl, 128), "PNG", cx - (qrSize/2), cursorY, qrSize, qrSize, undefined, "FAST");
      cursorY += qrSize;
    }

    col++;
    if(col>=cols){ col=0; x = margin; y += labelH + gap; }
    else { x += labelW + gap; }
  });

  doc.save("etiquetas.pdf");
}

/* ===========================
   IMPRESSÃO PERSONALIZADA (Centralizada + FIT garantido)
=========================== */
function abrirModalImpressao(){
  const items = itensSelecionadosOuTodos();
  if(!items.length){ alert("Nenhum item selecionado ou cadastrado."); return; }

  const container = document.getElementById("qtyList");
  container.innerHTML = `
    <div class="table-container">
      <table class="table is-fullwidth is-striped is-hoverable is-size-7">
        <thead>
          <tr>
            <th>#</th><th>Descrição</th><th>Unid.</th><th>PV</th><th>EAN</th><th class="has-text-right">Qtd Etiquetas</th>
          </tr>
        </thead>
        <tbody>
          ${items.map((p,i)=>`
            <tr>
              <td>${i+1}</td>
              <td>${escapeHTML(p.descricao)}</td>
              <td>${escapeHTML(p.unidade)}</td>
              <td>${toBR(p.precoVenda||0)}</td>
              <td>${p.ean||"—"}</td>
              <td class="has-text-right">
                <input type="number" min="0" step="1" value="1" class="input is-small qty-input" data-idx="${produtos.indexOf(p)}">
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
  document.getElementById("modalPrint").classList.add("is-active");
}
function fecharModalImpressao(){
  document.getElementById("modalPrint").classList.remove("is-active");
}
function confirmarImpressaoPersonalizada(){
  const opts = {
    logo: document.getElementById("optLogo").checked,
    data: document.getElementById("optData").checked,
    desc: document.getElementById("optDesc").checked,
    unid: document.getElementById("optUnid").checked,
    preco: document.getElementById("optPreco").checked,
    ean: document.getElementById("optEAN").checked,
    qr: document.getElementById("optQR").checked,
    compacto: document.getElementById("optCompact").checked
  };
  const qtyInputs = Array.from(document.querySelectorAll("#qtyList .qty-input"));
  const itensComQtd = qtyInputs.map(input=>{
    const idx = +input.dataset.idx;
    const qtd = Math.max(0, parseInt(input.value||"0"));
    return { produto: produtos[idx], qtd };
  }).filter(x=>x.qtd>0);

  if(!itensComQtd.length){ alert("Defina ao menos 1 etiqueta a imprimir."); return; }

  gerarEtiquetasPersonalizadasPDF(opts, itensComQtd);
  fecharModalImpressao();
}

function gerarEtiquetasPersonalizadasPDF(opts, itensComQtd){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:"portrait", unit:"mm", format:"a4"});

  const margin = 8;
  const pageW = 210, pageH = 297;
  const cols = opts.compacto ? 4 : 3;
  const gap = opts.compacto ? 5 : 6;
  const labelW = (pageW - margin*2 - gap*(cols-1)) / cols;
  const labelH = opts.compacto ? 30 : 38;
  const fontScale = opts.compacto ? 0.9 : 1.0;
  const pad = 3; // *** padding interno para não encostar na borda ***

  let x = margin, y = margin, col=0;

  const expandidos = [];
  itensComQtd.forEach(({produto, qtd})=>{
    for(let i=0;i<qtd;i++) expandidos.push(produto);
  });

  expandidos.forEach((p)=>{
    if(y + labelH > pageH - margin){
      doc.addPage();
      x = margin; y = margin; col=0;
    }

    // Borda da etiqueta
    doc.setDrawColor(210); doc.setLineWidth(0.2);
    doc.rect(x, y, labelW, labelH);

    // Geometria interna
    const innerW = labelW - 2*pad;
    const cx = x + labelW/2;
    let cursorY = y + pad;

    // LOGO
    if(opts.logo && typeof logoDataUrl === "string" && logoDataUrl.length){
      const maxLW = Math.min(18 * fontScale, innerW);
      const maxLH = 7  * fontScale;
      doc.addImage(logoDataUrl, "PNG", cx - (maxLW/2), cursorY, maxLW, maxLH, undefined, "FAST");
      cursorY += maxLH + 1.5;
    }

    // Data
    if(opts.data){
      doc.setFontSize(7*fontScale); doc.setTextColor(80);
      doc.text(`Data: ${dataHojeBR()}`, cx, cursorY, {align:"center"});
      cursorY += 3.8;
    }

    // Descrição (máx 2 linhas) — centralizada e com largura útil
    if(opts.desc){
      doc.setFontSize(9*fontScale); doc.setTextColor(0);
      const desc = (p.descricao||"").toString();
      const descLines = doc.splitTextToSize(desc, Math.max(10, innerW));
      const linesPrinted = Math.min(2, descLines.length);
      doc.text(descLines.slice(0,linesPrinted), cx, cursorY, {align:"center"});
      cursorY += linesPrinted * (opts.compacto?3.5:4) + 0.5;
    }

    // Unidade
    if(opts.unid){
      doc.setFontSize(7.5*fontScale); doc.setTextColor(60);
      doc.text(`Unid: ${p.unidade}`, cx, cursorY, {align:"center"});
      cursorY += opts.compacto ? 4.3 : 5;
    }

    // Preço
    if(opts.preco){
      doc.setFontSize(11*fontScale); doc.setTextColor(0);
      doc.text(`Preço: ${toBR(p.precoVenda||0)}`, cx, cursorY, {align:"center"});
      cursorY += opts.compacto ? 5.5 : 6;
    }

    // Espaço restante dentro da etiqueta
    let remainingH = (y + labelH - pad) - cursorY;

    // Código de barras (usa toda a largura útil e ajusta altura ao espaço disponível)
    if(opts.ean && p.ean && validarEAN13(p.ean) && remainingH > 6){
      const barData = gerarBarcodeDataURL(p.ean, 240, 80);
      if(barData){
        const barW = innerW; // largura útil total
        // altura adequada: no mínimo 8mm e no máximo 12mm, respeitando o espaço
        const barH = Math.max(8, Math.min(12, remainingH - (opts.qr ? 12 : 2))); 
        doc.addImage(barData, "PNG", cx - (barW/2), cursorY, barW, barH, undefined, "FAST");
        cursorY += barH + 1.5;
        remainingH = (y + labelH - pad) - cursorY;
      }
    }

    // QR centralizado, somente se houver espaço restante
    if(opts.qr && p.produtoUrl && remainingH > 8){
      const size = Math.max(10, Math.min(14, remainingH)); // ajusta ao espaço
      const qrData = gerarQRDataURL(p.produtoUrl, 128);
      doc.addImage(qrData, "PNG", cx - (size/2), cursorY, size, size, undefined, "FAST");
      cursorY += size;
    }

    // próxima etiqueta
    col++;
    if(col>=cols){ col=0; x = margin; y += labelH + (opts.compacto?4:6); }
    else { x += labelW + gap; }
  });

  doc.save("etiquetas-personalizadas.pdf");
}

/* ===========================
   A4 SIMPLES (MÁX 20 ITENS)
=========================== */
function imprimirA4_20itens(){
  let items = itensSelecionadosOuTodos();
  if(!items.length){ alert("Nenhum item selecionado ou cadastrado."); return; }
  if(items.length > 20){ items = items.slice(0,20); }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:"portrait", unit:"mm", format:"a4"});

  const pageW = 210, pageH = 297, margin = 10;
  const cols = 2, gapX = 8, rowH = 28; // 10 linhas x 2 colunas = 20 itens
  const cellW = (pageW - margin*2 - gapX*(cols-1)) / cols;

  doc.setFontSize(14);
  doc.text("Lista de Preços (A4 • até 20 itens)", margin, 12);
  doc.setFontSize(9);
  doc.setTextColor(90);
  doc.text(`Gerado em ${dataHojeBR()}`, pageW - margin, 12, {align:"right"});
  doc.setTextColor(0);

  let x = margin, y = 18, col = 0;

  items.forEach((p)=>{
    if(y + rowH > pageH - margin){
      doc.addPage();
      doc.setFontSize(14);
      doc.text("Lista de Preços (A4 • até 20 itens)", margin, 12);
      doc.setFontSize(9); doc.setTextColor(90);
      doc.text(`Gerado em ${dataHojeBR()}`, pageW - margin, 12, {align:"right"}); doc.setTextColor(0);
      x = margin; y = 18; col = 0;
    }

    doc.setDrawColor(220); doc.setLineWidth(0.2);
    doc.rect(x, y, cellW, rowH);

    // Descrição
    doc.setFontSize(10);
    const desc = (p.descricao||"").toString();
    const lines = doc.splitTextToSize(desc, cellW - 4);
    doc.text(lines, x + 2, y + 6);

    // UNIDADE
    doc.setFontSize(9);
    doc.setTextColor(80);
    doc.text(`Unid: ${p.unidade}`, x + 2, y + 11);
    doc.setTextColor(0);

    // PV
    doc.setFontSize(11);
    doc.text(`PV: ${toBR(p.precoVenda||0)}`, x + 2, y + 18);

    // EAN texto
    const eanTxt = (p.ean && validarEAN13(p.ean)) ? p.ean : "—";
    doc.setFontSize(9);
    doc.text(`EAN: ${eanTxt}`, x + 2, y + 23);

    // Código de barras pequeno
    if(p.ean && validarEAN13(p.ean)){
      const barData = gerarBarcodeDataURL(p.ean, 220, 70);
      if(barData){
        const imgW = Math.min(cellW - 60, 60);
        const imgH = 10;
        doc.addImage(barData, "PNG", x + cellW - imgW - 2, y + 14, imgW, imgH, undefined, "FAST");
      }
    }

    // próxima célula
    col++;
    if(col >= cols){ col = 0; x = margin; y += rowH + 4; }
    else { x += cellW + gapX; }
  });

  doc.save("lista-a4-20itens.pdf");
}

/* ===========================
   COPY EM MASSA & COPIES
=========================== */
function copiarSelecionadosEmMassa(){
  const items = itensSelecionadosOuTodos();
  if(!items.length){ toast("Nenhum item."); return; }

  const linhas = items.map((p,i)=>{
    const eanStr = (p.ean && validarEAN13(p.ean)) ? p.ean : "—";
    return `${i+1}. ${p.descricao}  |  EAN: ${eanStr}`;
  });
  copyToClipboard(linhas.join("\n"));
}
window.editarProduto = editarProduto;
window.excluirProduto = excluirProduto;
window.copiarDescricaoLinha = function(idx){
  const p = produtos[idx]; if(!p) return;
  if(p.descricao) copyToClipboard(p.descricao);
};
window.copiarEANLinha = function(idx){
  const p = produtos[idx]; if(!p) return;
  if(p.ean) copyToClipboard(p.ean);
};
