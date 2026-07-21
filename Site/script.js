// Dashboard API — motor de renderização (SPA por hash, uma única fonte de dados: data.json)
let DATA = null;
let currentTab = "vendas";     // vendas | devolucoes | cortes  (páginas geral/supervisor)
let currentVendTab = "geral";  // geral | faturado | afaturar   (página de vendedor)

const fmtBRL = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtInt = (v) => (v || 0).toLocaleString("pt-BR");
const fmtDate = (iso) => { const [y,m,d] = iso.split("-"); return `${d}/${m}/${y}`; };
const pct = (num, den) => den > 0 ? Math.round((num / den) * 100) : 0;
const clampPct = (p) => Math.max(0, Math.min(100, p));
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

function statusColor(p) {
  if (p >= 100) return "var(--good)";
  if (p >= 70) return "var(--blue)";
  if (p >= 40) return "var(--warn)";
  return "var(--critical)";
}

(function init() {
  if (!window.DASHBOARD_DATA) {
    document.getElementById("view").innerHTML = `<p class="empty-state">data.js não encontrado. Rode gerar-dados.ps1 e mantenha data.js na mesma pasta do index.html.</p>`;
    return;
  }
  DATA = window.DASHBOARD_DATA;
  document.getElementById("periodoLabel").textContent = DATA.periodo.label;
  document.getElementById("ultimaAtualizacao").textContent =
    `Atualizado em ${DATA.periodo.ultimaAtualizacao} · ${DATA.periodo.diasUteisRestantes} dias úteis restantes`;
  buildNav();
  window.addEventListener("hashchange", render);
  if (!location.hash) location.hash = "#/total";
  render();
})();

// ---------------------------------------------------------------------------
// Navegação
// ---------------------------------------------------------------------------
function buildNav() {
  const nav = document.getElementById("nav");
  const items = [
    { hash: "#/total", label: "Geral API Total", cor: "#1f5fa8" },
    { hash: "#/as", label: "Geral AS", cor: "#4a3aa7" },
    { hash: "#/varejo", label: "Geral Varejo", cor: "#eb6834" },
  ];
  let html = items.map(it => navItemHtml(it.hash, it.label, it.cor)).join("");
  html += `<div class="nav-section-label">Supervisores</div>`;
  html += DATA.supervisores.map(s => navItemHtml(`#/sup/${s.codigo}`, s.nome, s.cor)).join("");
  nav.innerHTML = html;
  nav.querySelectorAll(".nav-item").forEach(el => {
    el.addEventListener("click", () => { location.hash = el.dataset.hash; document.getElementById("sidebar").classList.remove("open"); });
  });
  document.getElementById("menuToggle").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
  document.getElementById("btnExportPdf").addEventListener("click", () => window.print());
  document.getElementById("btnExportExcel").addEventListener("click", exportCurrentViewToExcel);
}
function navItemHtml(hash, label, cor) {
  return `<div class="nav-item" data-hash="${hash}"><span class="nav-dot" style="background:${cor}"></span>${esc(label)}</div>`;
}
function setActiveNav(hash) {
  document.querySelectorAll(".nav-item").forEach(el => el.classList.toggle("active", el.dataset.hash === hash));
}

// ---------------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------------
function render() {
  const hash = location.hash || "#/total";
  const parts = hash.replace(/^#\//, "").split("/");
  currentTab = "vendas";
  const view = document.getElementById("view");
  const title = document.getElementById("pageTitle");

  if (parts[0] === "total") {
    setActiveNav("#/total"); title.textContent = "Geral API Total";
    renderGeralPage(view, DATA.totalGeral, { escopo: "total" });
  } else if (parts[0] === "as") {
    setActiveNav("#/as"); title.textContent = "Geral AS";
    renderGeralPage(view, DATA.geralAS, { escopo: "as" });
  } else if (parts[0] === "varejo") {
    setActiveNav("#/varejo"); title.textContent = "Geral Varejo";
    renderGeralPage(view, DATA.geralVarejo, { escopo: "varejo" });
  } else if (parts[0] === "sup") {
    const sup = DATA.supervisores.find(s => s.codigo === parts[1]);
    if (!sup) { view.innerHTML = `<p class="empty-state">Supervisor não encontrado.</p>`; return; }
    setActiveNav(`#/sup/${sup.codigo}`);
    if (parts[2] === "vendedor" && parts[3]) {
      const vend = sup.vendedores.find(v => v.codigo === parts[3]);
      if (!vend) { view.innerHTML = `<p class="empty-state">Vendedor não encontrado.</p>`; return; }
      title.textContent = `${vend.nome} · ${sup.nome}`;
      renderVendedorPage(view, sup, vend);
    } else {
      title.textContent = `Equipe ${sup.nome}`;
      renderSupervisorPage(view, sup);
    }
  } else {
    view.innerHTML = `<p class="empty-state">Página não encontrada.</p>`;
  }
  window.scrollTo(0, 0);
}

// ---------------------------------------------------------------------------
// Componentes reutilizáveis
// ---------------------------------------------------------------------------
function kpiCard(label, value, sub, opts = {}) {
  const cls = opts.status === "good" ? "good" : opts.status === "warn" ? "warn" : "";
  let bar = "";
  if (opts.progress !== undefined) {
    const p = clampPct(opts.progress);
    bar = `<div class="kpi-bar-track"><div class="kpi-bar-fill" style="width:${p}%;background:${statusColor(opts.progress)}"></div></div>`;
  }
  return `<div class="kpi ${cls}"><div class="kpi-label">${esc(label)}</div><div class="kpi-value">${value}</div>
    ${sub ? `<div class="kpi-sub">${sub}</div>` : ""}${bar}</div>`;
}

function kpiGrid(agg) {
  const progFat = pct(agg.faturadoLiquido, agg.metaFaturamento);
  const progGeral = pct(agg.faturadoLiquido + agg.aFaturar, agg.metaFaturamento);
  const progPos = pct(agg.positivacaoRealizado, agg.metaPositivacao);
  const devSub = agg.devolucao ? ` · devolução de ${fmtBRL(agg.devolucao)} já descontada` : "";
  return `
    <div class="kpi-group-label">Faturamento</div>
    <div class="kpi-grid">
      ${kpiCard("Meta de Faturamento", fmtBRL(agg.metaFaturamento), "Meta do mês (7 fornecedores)")}
      ${kpiCard("Faturado (líquido)", fmtBRL(agg.faturadoLiquido), `${progFat}% da meta${devSub}`, { progress: progFat })}
      ${kpiCard("A Faturar", fmtBRL(agg.aFaturar), "Pedidos ainda não faturados")}
      ${kpiCard("Projetado (Fat. + A Faturar)", fmtBRL(agg.faturadoLiquido + agg.aFaturar), `${progGeral}% da meta se tudo for faturado`, { progress: progGeral })}
    </div>
    <div class="kpi-group-label">Positivação (clientes)</div>
    <div class="kpi-grid">
      ${kpiCard("Clientes positivados", `${fmtInt(agg.positivacaoRealizado)} / ${fmtInt(agg.metaPositivacao)}`, `${progPos}% da meta · ${fmtInt(agg.positivacaoFaturado)} faturados, ${fmtInt(agg.positivacaoAFaturar)} a faturar`, { progress: progPos })}
      ${kpiCard("Devolveram no mês", fmtInt(agg.clientesComDevolucao), "Não contam como positivados, mesmo com pedido no mês")}
    </div>
  `;
}

function fornecedorBars(agg) {
  const maxScale = Math.max(1, ...DATA.fornecedores.map(f => {
    const b = agg.porFornecedor[f.id] || {};
    return Math.max(b.metaFaturamento || 0, (b.faturadoLiquido || 0) + (b.aFaturar || 0));
  }));
  const rows = DATA.fornecedores.map(f => {
    const b = agg.porFornecedor[f.id] || { metaFaturamento: 0, faturadoLiquido: 0, aFaturar: 0, positivacaoRealizado: 0, metaPositivacao: 0 };
    const wFat = clampPct((b.faturadoLiquido / maxScale) * 100);
    const wProj = clampPct(((b.faturadoLiquido + b.aFaturar) / maxScale) * 100);
    const wMeta = clampPct((b.metaFaturamento / maxScale) * 100);
    return `<div class="forn-bar-row">
      <div class="forn-bar-name"><span class="forn-bar-dot" style="background:${f.cor}"></span>${esc(f.nome)}</div>
      <div class="forn-bar-track" title="${esc(f.nome)}">
        <div class="forn-bar-fill" style="width:${wProj}%;background:${f.cor};opacity:0.35"></div>
        <div class="forn-bar-fill" style="width:${wFat}%;background:${f.cor}"></div>
        <div class="forn-bar-meta-mark" style="left:${wMeta}%"></div>
      </div>
      <div class="forn-bar-values">${fmtBRL(b.faturadoLiquido)} <span style="opacity:.55">/ ${fmtBRL(b.metaFaturamento)}</span></div>
    </div>`;
  }).join("");
  return `<div class="forn-bars">${rows}</div>
    <div class="legend-row">
      <div class="legend-item"><span class="legend-mark"></span> Meta</div>
      <div class="legend-item"><span class="legend-dot" style="background:#1f5fa8;opacity:.35"></span> Faturado + A Faturar</div>
      <div class="legend-item"><span class="legend-dot" style="background:#1f5fa8"></span> Faturado líquido</div>
    </div>`;
}

function fornecedorTable(agg) {
  const rows = DATA.fornecedores.map(f => {
    const b = agg.porFornecedor[f.id] || { metaFaturamento: 0, faturadoLiquido: 0, aFaturar: 0, positivacaoRealizado: 0, metaPositivacao: 0, devolucao: 0, bonificacao: 0 };
    const p = pct(b.faturadoLiquido, b.metaFaturamento);
    return `<tr>
      <td><span class="forn-bar-dot" style="background:${f.cor};display:inline-block;margin-right:6px"></span>${esc(f.nome)}</td>
      <td class="num">${fmtBRL2(b.metaFaturamento)}</td>
      <td class="num">${fmtBRL2(b.faturadoLiquido)}</td>
      <td class="num">${fmtBRL2(b.aFaturar)}</td>
      <td class="num">${p}%</td>
      <td class="num">${fmtInt(b.positivacaoRealizado)} / ${fmtInt(b.metaPositivacao)}</td>
      <td class="num">${b.devolucao ? fmtBRL2(b.devolucao) : "—"}</td>
    </tr>`;
  }).join("");
  return `<div class="table-scroll"><table class="forn-summary-table"><thead><tr>
    <th>Fornecedor</th><th class="num">Meta</th><th class="num">Faturado líq.</th><th class="num">A Faturar</th>
    <th class="num">% Meta</th><th class="num">Positivação</th><th class="num">Devolução</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

// ---------------------------------------------------------------------------
// Página Geral (Total / AS / Varejo)
// ---------------------------------------------------------------------------
function renderGeralPage(view, agg, opts) {
  view.innerHTML = `
    ${kpiGrid(agg)}
    <div class="section-title"><span class="bar"></span>Meta x Realizado por Fornecedor</div>
    <div class="card">${fornecedorBars(agg)}</div>
    <div class="card">${fornecedorTable(agg)}</div>
    ${opts.escopo === "total" ? rankingSupervisoresCard() : ""}
  `;
}
function rankingSupervisoresCard() {
  const rows = [...DATA.supervisores].sort((a,b) => (b.geral.faturadoLiquido) - (a.geral.faturadoLiquido)).map(s => {
    const p = pct(s.geral.faturadoLiquido, s.geral.metaFaturamento);
    return `<tr>
      <td><span class="forn-bar-dot" style="background:${s.cor};display:inline-block;margin-right:6px"></span>
        <a href="#/sup/${s.codigo}" style="color:var(--ink);text-decoration:none;font-weight:600">${esc(s.nome)}</a></td>
      <td class="num">${fmtBRL2(s.geral.metaFaturamento)}</td>
      <td class="num">${fmtBRL2(s.geral.faturadoLiquido)}</td>
      <td class="num">${fmtBRL2(s.geral.aFaturar)}</td>
      <td class="num" style="color:${statusColor(p)};font-weight:700">${p}%</td>
    </tr>`;
  }).join("");
  return `<div class="section-title"><span class="bar"></span>Por Supervisor</div>
    <div class="card"><div class="table-scroll"><table><thead><tr>
      <th>Supervisor</th><th class="num">Meta</th><th class="num">Faturado líq.</th><th class="num">A Faturar</th><th class="num">% Meta</th>
    </tr></thead><tbody>${rows}</tbody></table></div></div>`;
}

// ---------------------------------------------------------------------------
// Página Supervisor (tabs: Vendas / Devoluções / Cortes)
// ---------------------------------------------------------------------------
function renderSupervisorPage(view, sup) {
  view.innerHTML = `
    <div class="tabs" id="pageTabs">
      <button class="tab-btn ${currentTab==='vendas'?'active':''}" data-tab="vendas">Vendas</button>
      <button class="tab-btn ${currentTab==='devolucoes'?'active':''}" data-tab="devolucoes">Devoluções</button>
      <button class="tab-btn ${currentTab==='cortes'?'active':''}" data-tab="cortes">Cortes</button>
    </div>
    <div id="tabBody"></div>
  `;
  document.querySelectorAll("#pageTabs .tab-btn").forEach(b => b.addEventListener("click", () => {
    currentTab = b.dataset.tab; renderSupervisorPage(view, sup);
  }));
  const body = document.getElementById("tabBody");
  if (currentTab === "vendas") renderSupervisorVendas(body, sup);
  else if (currentTab === "devolucoes") renderDevolucoesTab(body, sup.codigo, sup.nome);
  else renderCortesTab(body, sup.codigo, sup.nome);
}

function renderSupervisorVendas(body, sup) {
  const vendedores = [...sup.vendedores].sort((a,b) => (b.geral.faturadoLiquido+b.geral.aFaturar) - (a.geral.faturadoLiquido+a.geral.aFaturar));
  body.innerHTML = `
    ${kpiGrid(sup.geral)}
    <div class="section-title"><span class="bar"></span>Meta x Realizado por Fornecedor — equipe ${esc(sup.nome)}</div>
    <div class="card">${fornecedorBars(sup.geral)}</div>
    <div class="card">${fornecedorTable(sup.geral)}</div>

    <div class="section-title"><span class="bar"></span>Vendedores da equipe (${vendedores.length})</div>
    <div class="vendedor-list">
      ${vendedores.map(v => {
        const p = pct(v.geral.faturadoLiquido, v.geral.metaFaturamento);
        return `<div class="vendedor-card" data-href="#/sup/${sup.codigo}/vendedor/${v.codigo}">
          <div class="vname">${esc(v.nome)}</div>
          <div class="vmeta"><span>Faturado</span><span>${fmtBRL(v.geral.faturadoLiquido)}</span></div>
          <div class="vmeta"><span>Meta</span><span>${fmtBRL(v.geral.metaFaturamento)}</span></div>
          <div class="vmeta"><span>A Faturar</span><span>${fmtBRL(v.geral.aFaturar)}</span></div>
          <div class="vbar"><div style="width:${clampPct(p)}%;background:${statusColor(p)}"></div></div>
        </div>`;
      }).join("")}
    </div>
  `;
  body.querySelectorAll(".vendedor-card").forEach(el => el.addEventListener("click", () => location.hash = el.dataset.href));
}

// ---------------------------------------------------------------------------
// Página Vendedor (tabs: Geral / Faturado / A Faturar)
// ---------------------------------------------------------------------------
function renderVendedorPage(view, sup, vend) {
  const pedidos = vend.pedidos || [];
  const faturados = pedidos.filter(p => p.status === "FATURADO");
  const aFaturar = pedidos.filter(p => p.status !== "FATURADO");
  const listaAtual = currentVendTab === "faturado" ? faturados : currentVendTab === "afaturar" ? aFaturar : pedidos;

  view.innerHTML = `
    <div class="breadcrumb"><a href="#/sup/${sup.codigo}">Equipe ${esc(sup.nome)}</a> / ${esc(vend.nome)}</div>
    ${kpiGrid(vend.geral)}
    <div class="section-title"><span class="bar"></span>Faturamento por Fornecedor</div>
    <div class="card">${fornecedorBars(vend.geral)}</div>
    <div class="card">${fornecedorTable(vend.geral)}</div>

    <div class="section-title"><span class="bar"></span>Pedidos</div>
    <div class="tabs" id="vendTabs">
      <button class="tab-btn ${currentVendTab==='geral'?'active':''}" data-tab="geral">Geral (${pedidos.length})</button>
      <button class="tab-btn ${currentVendTab==='faturado'?'active':''}" data-tab="faturado">Faturado (${faturados.length})</button>
      <button class="tab-btn ${currentVendTab==='afaturar'?'active':''}" data-tab="afaturar">A Faturar (${aFaturar.length})</button>
    </div>
    <div class="card">${pedidosTable(listaAtual)}</div>
  `;
  document.querySelectorAll("#vendTabs .tab-btn").forEach(b => b.addEventListener("click", () => {
    currentVendTab = b.dataset.tab; renderVendedorPage(view, sup, vend);
  }));
}
function pedidosTable(lista) {
  if (!lista.length) return `<p class="empty-state">Nenhum pedido nesta aba.</p>`;
  const rows = lista.map(p => `<tr>
    <td>${fmtDate(p.data)}</td>
    <td>${esc(p.numeroPedido)}</td>
    <td>${esc(p.codCliente)}</td>
    <td>${esc(p.razaoSocial)}</td>
    <td><span class="status-pill ${p.status==='FATURADO'?'status-faturado':'status-afaturar'}">${p.status==='FATURADO'?'Faturado':'A Faturar'}</span></td>
    <td class="num">${fmtBRL2(p.valor)}</td>
  </tr>`).join("");
  return `<div class="table-scroll"><table id="exportTable"><thead><tr>
    <th>Data digitação</th><th>Nº Pedido</th><th>Cód. Cliente</th><th>Razão Social</th><th>Status</th><th class="num">Valor</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

// ---------------------------------------------------------------------------
// Devoluções / Cortes (por supervisor, reutilizado também no Geral Total)
// ---------------------------------------------------------------------------
function renderDevolucoesTab(body, codigo, nome) {
  const bloco = DATA.devolucoes.porSupervisor.find(s => s.codigo === codigo);
  const itens = bloco ? bloco.itens : [];
  const porFornecedor = {};
  itens.forEach(i => { const f = i.fornecedor || "Outros"; porFornecedor[f] = (porFornecedor[f]||0) + i.valor; });
  const fornRows = Object.entries(porFornecedor).sort((a,b)=>b[1]-a[1]).map(([f,v]) => `<tr><td>${esc(f)}</td><td class="num">${fmtBRL2(v)}</td></tr>`).join("");

  body.innerHTML = `
    <div class="kpi-grid">
      ${kpiCard("Devolução total — " + esc(nome), fmtBRL2(bloco ? bloco.valorTotal : 0), `${itens.length} itens no mês`)}
    </div>
    <div class="section-title"><span class="bar"></span>Por Fornecedor</div>
    <div class="card"><div class="table-scroll"><table><thead><tr><th>Fornecedor</th><th class="num">Valor devolvido</th></tr></thead>
      <tbody>${fornRows || '<tr><td colspan="2" class="empty-state">Sem devoluções</td></tr>'}</tbody></table></div></div>
    <div class="section-title"><span class="bar"></span>Detalhamento (produto, pedido, cliente)</div>
    <div class="card">${devolucoesTable(itens)}</div>
  `;
}
function devolucoesTable(itens) {
  if (!itens.length) return `<p class="empty-state">Nenhuma devolução no mês.</p>`;
  const rows = itens.map(i => `<tr>
    <td>${fmtDate(i.data)}</td><td>${esc(i.numeroPedido)}</td><td>${esc(i.codCliente)}</td>
    <td>${esc(i.razaoSocial)}</td><td>${esc(i.produto)}</td><td>${esc(i.fornecedor||"—")}</td>
    <td>${esc(i.vendedor)}</td><td class="num">${fmtBRL2(i.valor)}</td>
  </tr>`).join("");
  return `<div class="table-scroll"><table id="exportTable"><thead><tr>
    <th>Data</th><th>Nº Pedido</th><th>Cód. Cliente</th><th>Razão Social</th><th>Produto</th><th>Fornecedor</th><th>Vendedor</th><th class="num">Valor</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

function renderCortesTab(body, codigo, nome) {
  const bloco = DATA.cortes.porSupervisor.find(s => s.codigo === codigo);
  const itens = bloco ? bloco.itens : [];
  body.innerHTML = `
    <div class="kpi-grid">
      ${kpiCard("Vendido em itens de corte — " + esc(nome), fmtBRL2(bloco ? bloco.valorTotal : 0), `${itens.length} pedidos no mês`)}
    </div>
    <div class="section-title"><span class="bar"></span>Alocação por Produto (risco de corte)</div>
    <div class="card">${cortesProdutosTable()}</div>
    <div class="section-title"><span class="bar"></span>Detalhamento (produto, pedido, valor) — ${esc(nome)}</div>
    <div class="card">${cortesTable(itens)}</div>
  `;
}
function cortesProdutosTable() {
  const produtos = DATA.cortes.produtos.slice(0, 40);
  if (!produtos.length) return `<p class="empty-state">Sem dados de corte.</p>`;
  const rows = produtos.map(p => {
    const consumoPct = p.alocado > 0 ? Math.round((p.vendidoUnid / p.alocado) * 100) : 0;
    return `<tr><td>${esc(p.produto)}</td><td>${esc(p.categoria)}</td><td>${esc(p.fornecedor||"—")}</td>
      <td class="num">${fmtBRL2(p.valor)}</td>
      <td class="num">${fmtInt(p.vendidoUnid)} / ${fmtInt(p.alocado)}</td>
      <td class="num" style="color:${statusColor(100-consumoPct)};font-weight:700">${consumoPct}%</td></tr>`;
  }).join("");
  return `<div class="table-scroll"><table><thead><tr>
    <th>Produto</th><th>Categoria</th><th>Fornecedor</th><th class="num">Valor vendido no mês</th>
    <th class="num">Unid. vendidas / Alocação</th><th class="num">% consumido</th>
  </tr></thead><tbody>${rows}</tbody></table></div>
  <p class="hint" style="margin-top:8px">Mostrando os 40 produtos de maior valor vendido dentre os ${DATA.cortes.produtos.length} itens de alocação controlada. % consumido acima de 100% indica que a venda já superou a alocação de referência — risco de corte no restante do mês.</p>`;
}
function cortesTable(itens) {
  if (!itens.length) return `<p class="empty-state">Nenhum pedido com produto de corte no mês.</p>`;
  const rows = itens.map(i => `<tr>
    <td>${fmtDate(i.data)}</td><td>${esc(i.numeroPedido)}</td><td>${esc(i.produto)}</td>
    <td>${esc(i.fornecedor||"—")}</td><td>${esc(i.vendedor)}</td>
    <td><span class="status-pill ${i.status==='FATURADO'?'status-faturado':'status-afaturar'}">${i.status==='FATURADO'?'Faturado':'A Faturar'}</span></td>
    <td class="num">${fmtBRL2(i.valor)}</td>
  </tr>`).join("");
  return `<div class="table-scroll"><table id="exportTable"><thead><tr>
    <th>Data</th><th>Nº Pedido</th><th>Produto</th><th>Fornecedor</th><th>Vendedor</th><th>Status</th><th class="num">Valor</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

// ---------------------------------------------------------------------------
// Exportar Excel (CSV — abre nativamente no Excel, sem dependências externas)
// ---------------------------------------------------------------------------
function exportCurrentViewToExcel() {
  const table = document.getElementById("exportTable") || document.querySelector(".view table");
  if (!table) { alert("Não há tabela para exportar nesta tela."); return; }
  const rows = [...table.querySelectorAll("tr")].map(tr =>
    [...tr.children].map(td => `"${td.textContent.trim().replace(/"/g, '""')}"`).join(";")
  );
  const csv = "﻿" + rows.join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `dashboard-api-${(document.getElementById("pageTitle").textContent||"export").toLowerCase().replace(/\s+/g,"-")}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}
