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
function statusPillHtml(status) {
  const ok = status === "FATURADO";
  return `<span class="status-pill ${ok?'status-faturado':'status-afaturar'}">${ok?'Faturado':'A Faturar'}</span>`;
}

// ---------------------------------------------------------------------------
// Busca genérica (pedido winthor/clube, vendedor, RCA, CNPJ, cód. cliente) — item 7
// ---------------------------------------------------------------------------
function searchBoxHtml(id) {
  return `<div class="search-box"><input type="text" id="${id}" class="search-input"
    placeholder="Buscar por pedido Winthor, pedido Clube, nota fiscal, vendedor, cód. RCA, CNPJ ou cód. cliente..."></div>`;
}
function wireSearchBox(inputId, containerEl) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("input", () => {
    const term = input.value.trim().toLowerCase();
    containerEl.querySelectorAll("[data-search]").forEach(row => {
      const match = !term || row.dataset.search.includes(term);
      row.style.display = match ? "" : "none";
      // se a linha de detalhe (accordion) estiver aberta e a linha-mãe some, esconde junto
      if (!match && row.nextElementSibling && row.nextElementSibling.classList.contains("grupo-detail")) {
        row.nextElementSibling.style.display = "none";
        row.classList.remove("open");
      }
    });
  });
}

// ---------------------------------------------------------------------------
// Tabela agrupada cliente+pedido, com detalhe de produtos expansível — Devoluções/Cortes
// ---------------------------------------------------------------------------
function grupoTable(grupos, opts = {}) {
  if (!grupos.length) return `<p class="empty-state">Nenhum registro no mês.</p>`;
  const showSup = !!opts.showSupervisor;
  const showStatus = !!opts.showStatus;
  const showCorte = !!opts.showCorte;
  const rows = grupos.map(g => {
    const searchText = [g.numeroPedidoWinthor, g.numeroPedidoClube, g.numeroNotaFiscal, g.codigoRCA, g.vendedor, g.cnpj, g.codCliente, g.razaoSocial]
      .join(" ").toLowerCase();
    const detailCols = 9 + (showSup ? 1 : 0);
    return `
      <tr class="grupo-row" data-search="${esc(searchText)}">
        <td class="grupo-toggle">▸</td>
        <td>${fmtDate(g.data)}</td>
        <td>${esc(g.numeroPedidoWinthor)}</td>
        <td>${esc(g.numeroPedidoClube)}</td>
        <td>${esc(g.numeroNotaFiscal)}</td>
        <td>${esc(g.codigoRCA)}</td>
        <td>${esc(g.codCliente)}</td>
        <td>${esc(g.cnpj)}</td>
        <td>${esc(g.razaoSocial)}</td>
        ${showSup ? `<td>${esc(g.supervisor)}</td>` : ""}
        <td>${esc(g.vendedor)}</td>
        <td class="num">${fmtBRL2(g.valorTotal)}</td>
      </tr>
      <tr class="grupo-detail" style="display:none">
        <td colspan="${detailCols + 2}">
          <table class="produtos-subtable"><thead><tr>
            <th>Cód. Produto</th><th>Produto</th><th>Fornecedor</th>
            ${showStatus ? "<th>Status</th>" : ""}
            ${showCorte ? `<th>Embalagem</th><th class="num">Qt. Corte</th><th class="num">Preço Unit.</th>` : ""}
            <th class="num">Valor</th>
          </tr></thead><tbody>
            ${g.produtos.map(p => `<tr>
              <td>${esc(p.codProduto)}</td><td>${esc(p.produto)}</td><td>${esc(p.fornecedor||"—")}</td>
              ${showStatus ? `<td>${statusPillHtml(p.status)}</td>` : ""}
              ${showCorte ? `<td>${esc(p.embalagem)}</td><td class="num">${fmtInt(p.qtCorte)}</td><td class="num">${fmtBRL2(p.precoUnit)}</td>` : ""}
              <td class="num">${fmtBRL2(p.valor)}</td>
            </tr>`).join("")}
          </tbody></table>
        </td>
      </tr>`;
  }).join("");
  const idAttr = opts.exportable === false ? "" : ` id="exportTable"`;
  return `<div class="table-scroll"><table${idAttr}><thead><tr>
    <th class="grupo-toggle"></th><th>Data</th><th>Pedido Winthor</th><th>Pedido Clube</th><th>Nota Fiscal</th><th>Cód. RCA</th><th>Cód. Cliente</th><th>CNPJ</th>
    <th>Razão Social</th>${showSup ? "<th>Supervisor</th>" : ""}<th>Vendedor</th><th class="num">Valor</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}
function wireGrupoTable(container) {
  container.querySelectorAll(".grupo-row").forEach(row => {
    row.addEventListener("click", () => {
      const detail = row.nextElementSibling;
      const open = detail.style.display !== "none";
      detail.style.display = open ? "none" : "";
      row.classList.toggle("open", !open);
    });
  });
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
  let html = `<div class="nav-section-label">Geral</div>`;
  html += navItemHtml("#/total", "Geral API Total", "#1f5fa8");
  html += navItemHtml("#/devolucoes/total", "Devoluções Geral", "#d03b3b");
  html += navItemHtml("#/cortes/total", "Cortes Geral", "#8a5a3b");
  html += `<div class="nav-section-label">Supervisores</div>`;
  html += DATA.supervisores.map(s => navItemHtml(`#/sup/${s.codigo}`, s.nome, s.cor)).join("");
  nav.innerHTML = html;
  nav.querySelectorAll(".nav-item").forEach(el => {
    el.addEventListener("click", () => { location.hash = el.dataset.hash; document.getElementById("sidebar").classList.remove("open"); });
  });
  document.getElementById("menuToggle").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
  document.getElementById("btnBack").addEventListener("click", () => history.back());
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
  } else if (parts[0] === "devolucoes" && parts[1] === "total") {
    setActiveNav("#/devolucoes/total"); title.textContent = "Devoluções Geral";
    renderDevolucoesGeralPage(view, "Devoluções Geral", DATA.devolucoes.geral);
  } else if (parts[0] === "cortes" && parts[1] === "total") {
    setActiveNav("#/cortes/total"); title.textContent = "Cortes Geral";
    renderCortesGeralPage(view, "Cortes Geral", DATA.cortes.geral);
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
      ${kpiCard("Fat + A Fat", fmtBRL(agg.faturadoLiquido + agg.aFaturar), `${progGeral}% da meta se tudo for faturado`, { progress: progGeral })}
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
    const b = agg.porFornecedor[f.id] || { metaFaturamento: 0, faturadoLiquido: 0, aFaturar: 0, positivacaoRealizado: 0, metaPositivacao: 0, positivacaoFaturado: 0, positivacaoAFaturar: 0, devolucao: 0, bonificacao: 0 };
    const p = pct(b.faturadoLiquido, b.metaFaturamento);
    const pProj = pct(b.faturadoLiquido + b.aFaturar, b.metaFaturamento);
    return `<tr>
      <td><span class="forn-bar-dot" style="background:${f.cor};display:inline-block;margin-right:6px"></span>${esc(f.nome)}</td>
      <td class="num">${fmtBRL2(b.metaFaturamento)}</td>
      <td class="num">${fmtBRL2(b.faturadoLiquido)}</td>
      <td class="num">${fmtBRL2(b.aFaturar)}</td>
      <td class="num">${fmtBRL2(b.faturadoLiquido + b.aFaturar)}<br><span style="font-size:10.5px;color:var(--ink-muted)">${pProj}% da meta</span></td>
      <td class="num">${p}%</td>
      <td class="num">${fmtInt(b.positivacaoRealizado)} / ${fmtInt(b.metaPositivacao)}<br><span style="font-size:10.5px;color:var(--ink-muted)">${fmtInt(b.positivacaoFaturado)}F · ${fmtInt(b.positivacaoAFaturar)}AF</span></td>
      <td class="num">${b.devolucao ? fmtBRL2(b.devolucao) : "—"}</td>
    </tr>`;
  }).join("");
  return `<div class="table-scroll"><table class="forn-summary-table"><thead><tr>
    <th>Fornecedor</th><th class="num">Meta</th><th class="num">Faturado líq.</th><th class="num">A Faturar</th>
    <th class="num">Fat + A Fat</th><th class="num">% Meta</th><th class="num">Positivação</th><th class="num">Devolução</th>
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
    <div class="card" id="pedidosCard">${pedidosTable(listaAtual, "pedidosSearch")}</div>

    <div class="section-title"><span class="bar"></span>Devoluções e Cortes deste vendedor</div>
    <div id="vendDevCortesSlot"></div>
  `;
  document.querySelectorAll("#vendTabs .tab-btn").forEach(b => b.addEventListener("click", () => {
    currentVendTab = b.dataset.tab; renderVendedorPage(view, sup, vend);
  }));
  const pedidosCard = document.getElementById("pedidosCard");
  wireSearchBox("pedidosSearch", pedidosCard);
  renderVendedorDevCortes(document.getElementById("vendDevCortesSlot"), sup.codigo, vend.codigo);
}
function renderVendedorDevCortes(container, codigoSup, codigoVend) {
  const devBloco = DATA.devolucoes.porSupervisor.find(s => s.codigo === codigoSup);
  const corteBloco = DATA.cortes.porSupervisor.find(s => s.codigo === codigoSup);
  const devGrupos = devBloco ? devBloco.grupos.filter(g => g.codigoRCA === codigoVend) : [];
  const corteGrupos = corteBloco ? corteBloco.grupos.filter(g => g.codigoRCA === codigoVend) : [];
  const devTotal = devGrupos.reduce((a,g) => a + g.valorTotal, 0);
  const corteTotal = corteGrupos.reduce((a,g) => a + g.valorTotal, 0);
  container.innerHTML = `
    <div class="kpi-grid">
      ${kpiCard("Devolução no mês", fmtBRL2(devTotal), `${devGrupos.length} pedido(s) com devolução`)}
      ${kpiCard("Vendido em itens de corte", fmtBRL2(corteTotal), `${corteGrupos.length} pedido(s) com item de corte`)}
    </div>
    <div class="tabs" id="devCorteTabs">
      <button class="tab-btn active" data-tab="devolucoes">Devoluções (${devGrupos.length})</button>
      <button class="tab-btn" data-tab="cortes">Cortes (${corteGrupos.length})</button>
    </div>
    <div class="card" id="devCorteCard"></div>
  `;
  const showTab = (tab) => {
    container.querySelectorAll("#devCorteTabs .tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
    const card = document.getElementById("devCorteCard");
    const grupos = tab === "cortes" ? corteGrupos : devGrupos;
    const opts = tab === "cortes" ? { showCorte: true, exportable: false } : { exportable: false };
    card.innerHTML = grupos.length
      ? `<p class="hint">Clique numa linha para ver os produtos daquele pedido.</p>${grupoTable(grupos, opts)}`
      : `<p class="empty-state">Nenhum registro no mês.</p>`;
    wireGrupoTable(card);
  };
  container.querySelectorAll("#devCorteTabs .tab-btn").forEach(b => b.addEventListener("click", () => showTab(b.dataset.tab)));
  showTab("devolucoes");
}
function pedidosTable(lista, containerId) {
  if (!lista.length) return `<p class="empty-state">Nenhum pedido nesta aba.</p>`;
  const rows = lista.map(p => {
    const searchText = [p.numeroPedidoWinthor, p.numeroPedidoClube, p.numeroNotaFiscal, p.codigoRCA, p.cnpj, p.codCliente, p.razaoSocial].join(" ").toLowerCase();
    return `<tr data-search="${esc(searchText)}">
    <td>${fmtDate(p.data)}</td>
    <td>${esc(p.numeroPedidoWinthor)}</td>
    <td>${esc(p.numeroPedidoClube)}</td>
    <td>${esc(p.numeroNotaFiscal)}</td>
    <td>${esc(p.codigoRCA)}</td>
    <td>${esc(p.codCliente)}</td>
    <td>${esc(p.cnpj)}</td>
    <td>${esc(p.razaoSocial)}</td>
    <td>${statusPillHtml(p.status)}</td>
    <td class="num">${fmtBRL2(p.valor)}</td>
  </tr>`;
  }).join("");
  return `${searchBoxHtml(containerId)}<div class="table-scroll"><table id="exportTable"><thead><tr>
    <th>Data digitação</th><th>Pedido Winthor</th><th>Pedido Clube</th><th>Nota Fiscal</th><th>Cód. RCA</th><th>Cód. Cliente</th><th>CNPJ</th><th>Razão Social</th><th>Status</th><th class="num">Valor</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

// ---------------------------------------------------------------------------
// Devoluções / Cortes (por supervisor, reutilizado também no Geral Total/AS/Varejo)
// ---------------------------------------------------------------------------
function fornecedorValorTable(grupos, label) {
  const porFornecedor = {};
  grupos.forEach(g => g.produtos.forEach(p => {
    const f = p.fornecedor || "Outros";
    porFornecedor[f] = (porFornecedor[f]||0) + p.valor;
  }));
  const rows = Object.entries(porFornecedor).sort((a,b)=>b[1]-a[1]).map(([f,v]) => `<tr><td>${esc(f)}</td><td class="num">${fmtBRL2(v)}</td></tr>`).join("");
  return `<div class="table-scroll"><table><thead><tr><th>Fornecedor</th><th class="num">${esc(label)}</th></tr></thead>
    <tbody>${rows || `<tr><td colspan="2" class="empty-state">Sem itens</td></tr>`}</tbody></table></div>`;
}
function vendedorValorTable(grupos, label) {
  const porVendedor = {};
  grupos.forEach(g => {
    const key = g.vendedor || "—";
    if (!porVendedor[key]) porVendedor[key] = { valor: 0, pedidos: 0 };
    porVendedor[key].valor += g.valorTotal;
    porVendedor[key].pedidos++;
  });
  const rows = Object.entries(porVendedor).sort((a,b)=>b[1].valor-a[1].valor)
    .map(([v,d]) => `<tr><td>${esc(v)}</td><td class="num">${fmtInt(d.pedidos)}</td><td class="num">${fmtBRL2(d.valor)}</td></tr>`).join("");
  return `<div class="table-scroll"><table><thead><tr><th>Vendedor</th><th class="num">Pedidos</th><th class="num">${esc(label)}</th></tr></thead>
    <tbody>${rows || `<tr><td colspan="3" class="empty-state">Sem itens</td></tr>`}</tbody></table></div>`;
}
function renderDevolucoesTab(body, codigo, nome) {
  const bloco = DATA.devolucoes.porSupervisor.find(s => s.codigo === codigo);
  renderGrupoESearch(body, "Devolução total — " + nome, bloco, { showSupervisor: false, showVendedor: true, fornecedorLabel: "Valor devolvido", detailTitle: "Detalhamento (cliente, CNPJ, produto, RCA)" });
}
function renderDevolucoesGeralPage(view, titulo, bloco) {
  renderGrupoESearch(view, titulo, bloco, { showSupervisor: true, fornecedorLabel: "Valor devolvido", detailTitle: "Detalhamento (cliente, CNPJ, produto, RCA, supervisor)" });
}
function renderCortesGeralPage(view, titulo, bloco) {
  view.innerHTML = `
    <div class="kpi-grid">${kpiCard(titulo, fmtBRL2(bloco ? bloco.valorTotal : 0), `${bloco ? bloco.grupos.length : 0} pedido(s) no mês`)}</div>
    <div class="section-title"><span class="bar"></span>Produtos mais cortados no mês</div>
    <div class="card">${cortesProdutosTable()}</div>
    <div id="cortesGrupoSlot"></div>
  `;
  renderGrupoESearch(document.getElementById("cortesGrupoSlot"), null, bloco, { showSupervisor: true, showCorte: true, fornecedorLabel: "Valor vendido", detailTitle: "Detalhamento (cliente, CNPJ, produto, RCA, supervisor)", skipKpi: true });
}
function renderCortesTab(body, codigo, nome) {
  const bloco = DATA.cortes.porSupervisor.find(s => s.codigo === codigo);
  body.innerHTML = `
    <div class="kpi-grid">${kpiCard("Vendido em itens de corte — " + esc(nome), fmtBRL2(bloco ? bloco.valorTotal : 0), `${bloco ? bloco.grupos.length : 0} pedido(s) no mês`)}</div>
    <div class="section-title"><span class="bar"></span>Produtos mais cortados no mês</div>
    <div class="card">${cortesProdutosTable()}</div>
    <div id="cortesGrupoSlot"></div>
  `;
  renderGrupoESearch(document.getElementById("cortesGrupoSlot"), null, bloco, { showSupervisor: false, showVendedor: true, showCorte: true, fornecedorLabel: "Valor vendido", detailTitle: `Detalhamento (produto, pedido, valor) — ${nome}`, skipKpi: true });
}

// Renderiza KPI (opcional) + tabela "Por Fornecedor" + busca + tabela agrupada com accordion,
// usado por Devoluções e Cortes (Geral/AS/Varejo/por Supervisor) — evita duplicar a mesma
// composição 4 vezes.
function renderGrupoESearch(container, kpiLabel, bloco, opts) {
  const grupos = bloco ? bloco.grupos : [];
  const searchId = "search-" + Math.random().toString(36).slice(2, 9);
  container.innerHTML = `
    ${opts.skipKpi ? "" : `<div class="kpi-grid">${kpiCard(kpiLabel, fmtBRL2(bloco ? bloco.valorTotal : 0), `${grupos.length} pedido(s) no mês`)}</div>`}
    ${opts.showVendedor ? `
      <div class="section-title"><span class="bar"></span>Por Fornecedor e por Vendedor</div>
      <div class="grid-2">
        <div class="card">${fornecedorValorTable(grupos, opts.fornecedorLabel)}</div>
        <div class="card">${vendedorValorTable(grupos, opts.fornecedorLabel)}</div>
      </div>
    ` : `
      <div class="section-title"><span class="bar"></span>Por Fornecedor</div>
      <div class="card">${fornecedorValorTable(grupos, opts.fornecedorLabel)}</div>
    `}
    <div class="section-title"><span class="bar"></span>${opts.detailTitle}</div>
    <p class="hint">Clique numa linha para ver os produtos daquele pedido.</p>
    <div class="card" id="${searchId}-card">${searchBoxHtml(searchId)}${grupoTable(grupos, opts)}</div>
  `;
  const card = document.getElementById(`${searchId}-card`);
  wireGrupoTable(card);
  wireSearchBox(searchId, card);
}
function cortesProdutosTable() {
  const produtos = DATA.cortes.produtos.slice(0, 40);
  if (!produtos.length) return `<p class="empty-state">Sem dados de corte.</p>`;
  const rows = produtos.map(p => `<tr><td>${esc(p.produto)}</td><td>${esc(p.fornecedor||"—")}</td>
      <td class="num">${fmtInt(p.qtCorte)}</td>
      <td class="num">${fmtBRL2(p.valor)}</td></tr>`).join("");
  return `<div class="table-scroll"><table><thead><tr>
    <th>Produto</th><th>Fornecedor</th><th class="num">Qt. Cortada no mês</th><th class="num">Valor cortado no mês</th>
  </tr></thead><tbody>${rows}</tbody></table></div>
  <p class="hint" style="margin-top:8px">Top 40 produtos com mais corte por valor, dentre ${DATA.cortes.produtos.length} itens cortados no mês.</p>`;
}

// ---------------------------------------------------------------------------
// Exportar Excel (CSV — abre nativamente no Excel, sem dependências externas)
// ---------------------------------------------------------------------------
function exportCurrentViewToExcel() {
  const table = document.getElementById("exportTable") || document.querySelector(".view table");
  if (!table) { alert("Não há tabela para exportar nesta tela."); return; }
  // linhas de detalhe do accordion (produtos de um pedido) ficam de fora do CSV - só a
  // tabela-resumo (cliente/pedido/valor) é exportada, senão o texto aninhado vira uma bagunça
  const rows = [...table.querySelectorAll("tr")].filter(tr => !tr.classList.contains("grupo-detail")).map(tr =>
    [...tr.children].filter(td => !td.classList.contains("grupo-toggle")).map(td => `"${td.textContent.trim().replace(/"/g, '""')}"`).join(";")
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

