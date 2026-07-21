// Ritmo Diário — quanto cada vendedor precisa positivar e faturar por dia útil restante
// para bater a meta do mês. Reaproveita o mesmo data.json do Dashboard API (Site/).
let DATA = null;

const fmtBRL = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtInt = (v) => (v || 0).toLocaleString("pt-BR");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const clampPct = (p) => Math.max(0, Math.min(100, p));

(function init() {
  if (!window.DASHBOARD_DATA) {
    document.getElementById("view").innerHTML = `<p class="empty-state">Site/data.js não encontrado. Rode gerar-dados.ps1 primeiro.</p>`;
    return;
  }
  DATA = window.DASHBOARD_DATA;
  document.getElementById("periodoLabel").textContent = DATA.periodo.label;
  document.getElementById("diasInfo").innerHTML =
    `${DATA.periodo.diasUteisPassados}/${DATA.periodo.diasUteisTotais} dias úteis usados<br><strong>${DATA.periodo.diasUteisRestantes} dias úteis restantes</strong>`;
  buildNav();
  window.addEventListener("hashchange", render);
  if (!location.hash) location.hash = "#/total";
  render();
})();

function buildNav() {
  const nav = document.getElementById("nav");
  const items = [
    { hash: "#/total", label: "Geral API Total", cor: "#1f5fa8" },
    { hash: "#/as", label: "Geral AS", cor: "#4a3aa7" },
    { hash: "#/varejo", label: "Geral Varejo", cor: "#eb6834" },
  ];
  let html = items.map(it => `<div class="nav-item" data-hash="${it.hash}"><span class="nav-dot" style="background:${it.cor}"></span>${esc(it.label)}</div>`).join("");
  html += `<div class="nav-section-label">Supervisores</div>`;
  html += DATA.supervisores.map(s => `<div class="nav-item" data-hash="#/sup/${s.codigo}"><span class="nav-dot" style="background:${s.cor}"></span>${esc(s.nome)}</div>`).join("");
  nav.innerHTML = html;
  nav.querySelectorAll(".nav-item").forEach(el => el.addEventListener("click", () => {
    location.hash = el.dataset.hash; document.getElementById("sidebar").classList.remove("open");
  }));
  document.getElementById("menuToggle").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
  document.getElementById("btnExportPdf").addEventListener("click", () => window.print());
  document.getElementById("btnExportExcel").addEventListener("click", exportToExcel);
}
function setActiveNav(hash) {
  document.querySelectorAll(".nav-item").forEach(el => el.classList.toggle("active", el.dataset.hash === hash));
}

function render() {
  const hash = location.hash || "#/total";
  const parts = hash.replace(/^#\//, "").split("/");
  const view = document.getElementById("view");
  const title = document.getElementById("pageTitle");

  if (parts[0] === "total") { setActiveNav("#/total"); title.textContent = "Ritmo — Geral API Total"; renderGeral(view, DATA.totalGeral); }
  else if (parts[0] === "as") { setActiveNav("#/as"); title.textContent = "Ritmo — Geral AS"; renderGeral(view, DATA.geralAS); }
  else if (parts[0] === "varejo") { setActiveNav("#/varejo"); title.textContent = "Ritmo — Geral Varejo"; renderGeral(view, DATA.geralVarejo); }
  else if (parts[0] === "sup") {
    const sup = DATA.supervisores.find(s => s.codigo === parts[1]);
    if (!sup) { view.innerHTML = `<p class="empty-state">Supervisor não encontrado.</p>`; return; }
    setActiveNav(`#/sup/${sup.codigo}`);
    if (parts[2] === "vendedor" && parts[3]) {
      const vend = sup.vendedores.find(v => v.codigo === parts[3]);
      if (!vend) { view.innerHTML = `<p class="empty-state">Vendedor não encontrado.</p>`; return; }
      title.textContent = `Ritmo — ${vend.nome} · ${sup.nome}`;
      renderVendedor(view, sup, vend);
    } else {
      title.textContent = `Ritmo — Equipe ${sup.nome}`;
      renderSupervisor(view, sup);
    }
  } else { view.innerHTML = `<p class="empty-state">Página não encontrada.</p>`; }
  window.scrollTo(0, 0);
}

// ---------------------------------------------------------------------------
// Cálculo do ritmo
// ---------------------------------------------------------------------------
function paceFaturamento(agg) {
  const proj = agg.faturadoLiquido + agg.aFaturar;
  const falta = Math.max(0, agg.metaFaturamento - proj);
  const dRest = DATA.periodo.diasUteisRestantes;
  const dPass = DATA.periodo.diasUteisPassados;
  const necessarioDia = dRest > 0 ? falta / dRest : falta;
  const atualDia = dPass > 0 ? proj / dPass : proj;
  return { falta, necessarioDia, atualDia, bateu: falta <= 0 };
}
function pacePositivacao(agg) {
  const falta = Math.max(0, agg.metaPositivacao - agg.positivacaoRealizado);
  const dRest = DATA.periodo.diasUteisRestantes;
  const dPass = DATA.periodo.diasUteisPassados;
  const necessarioDia = dRest > 0 ? falta / dRest : falta;
  const atualDia = dPass > 0 ? agg.positivacaoRealizado / dPass : agg.positivacaoRealizado;
  return { falta, necessarioDia, atualDia, bateu: falta <= 0 };
}
function paceBadge(pf) {
  if (pf.bateu) return `<span class="pace-badge ok">Meta batida</span>`;
  if (pf.atualDia >= pf.necessarioDia) return `<span class="pace-badge ok">No ritmo</span>`;
  if (pf.atualDia >= pf.necessarioDia * 0.7) return `<span class="pace-badge atencao">Atenção</span>`;
  return `<span class="pace-badge risco">Abaixo do ritmo</span>`;
}

// ---------------------------------------------------------------------------
// Componentes
// ---------------------------------------------------------------------------
function headerCards(agg) {
  const pf = paceFaturamento(agg);
  const pp = pacePositivacao(agg);
  return `
    <div class="kpi-group-label">Faturamento</div>
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Falta faturar (vendas+pendente)</div><div class="kpi-value">${fmtBRL(pf.falta)}</div>
        <div class="kpi-sub">Meta ${fmtBRL(agg.metaFaturamento)} · já projetado ${fmtBRL(agg.faturadoLiquido + agg.aFaturar)}</div></div>
      <div class="kpi"><div class="kpi-label">Ritmo necessário / dia útil</div><div class="kpi-value">${fmtBRL(pf.necessarioDia)}</div>
        <div class="kpi-sub">Ritmo médio até agora: ${fmtBRL(pf.atualDia)}/dia — ${paceBadge(pf)}</div></div>
    </div>
    <div class="kpi-group-label">Positivação (clientes)</div>
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Falta positivar</div><div class="kpi-value">${fmtInt(pp.falta)}</div>
        <div class="kpi-sub">Meta ${fmtInt(agg.metaPositivacao)} · já positivados ${fmtInt(agg.positivacaoRealizado)} (${fmtInt(agg.positivacaoFaturado)} faturados, ${fmtInt(agg.positivacaoAFaturar)} a faturar)</div></div>
      <div class="kpi"><div class="kpi-label">Clientes/dia necessário</div><div class="kpi-value">${pp.necessarioDia.toFixed(1)}</div>
        <div class="kpi-sub">Ritmo médio até agora: ${pp.atualDia.toFixed(1)}/dia — ${paceBadge(pp)}</div></div>
    </div>
  `;
}

function fornecedorPaceGrid(agg) {
  const cards = DATA.fornecedores.map(f => {
    const b = agg.porFornecedor[f.id] || { metaFaturamento:0, faturadoLiquido:0, aFaturar:0, metaPositivacao:0, positivacaoRealizado:0 };
    const pf = paceFaturamento(b);
    const pp = pacePositivacao(b);
    return `<div class="pace-card">
      <div class="pf-name"><span class="pf-dot" style="background:${f.cor}"></span>${esc(f.nome)}</div>
      <div class="pace-row"><span>Meta</span><b>${fmtBRL(b.metaFaturamento)}</b></div>
      <div class="pace-row"><span>Projetado</span><b>${fmtBRL(b.faturadoLiquido + b.aFaturar)}</b></div>
      <div class="pace-need">${fmtBRL(pf.necessarioDia)}<span style="font-size:11px;font-weight:600;color:var(--ink-muted)"> /dia útil</span></div>
      <div class="pace-row"><span>Clientes/dia p/ positivar</span><b>${pp.necessarioDia.toFixed(1)}</b></div>
      ${paceBadge(pf)}
    </div>`;
  }).join("");
  return `<div class="pace-grid">${cards}</div>`;
}

// ---------------------------------------------------------------------------
// Páginas
// ---------------------------------------------------------------------------
function renderGeral(view, agg) {
  view.innerHTML = `
    ${headerCards(agg)}
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Fornecedor</div>
    <div class="card">${fornecedorPaceGrid(agg)}</div>
  `;
}

function renderSupervisor(view, sup) {
  const vendedores = [...sup.vendedores].sort((a,b) => {
    const pa = paceFaturamento(a.geral), pb = paceFaturamento(b.geral);
    return (pb.necessarioDia - pb.atualDia) - (pa.necessarioDia - pa.atualDia);
  });
  view.innerHTML = `
    ${headerCards(sup.geral)}
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Fornecedor — equipe ${esc(sup.nome)}</div>
    <div class="card">${fornecedorPaceGrid(sup.geral)}</div>
    <div class="section-title"><span class="bar"></span>Vendedores — ritmo individual (${vendedores.length})</div>
    <div class="table-scroll"><table id="exportTable"><thead><tr>
      <th>Vendedor</th><th class="num">Falta faturar</th><th class="num">R$/dia necessário</th>
      <th class="num">R$/dia atual</th><th class="num">Falta positivar</th><th class="num">Clientes/dia necessário</th><th>Situação</th>
    </tr></thead><tbody>
      ${vendedores.map(v => {
        const pf = paceFaturamento(v.geral), pp = pacePositivacao(v.geral);
        return `<tr style="cursor:pointer" onclick="location.hash='#/sup/${sup.codigo}/vendedor/${v.codigo}'">
          <td><strong>${esc(v.nome)}</strong></td>
          <td class="num">${fmtBRL2(pf.falta)}</td>
          <td class="num">${fmtBRL2(pf.necessarioDia)}</td>
          <td class="num">${fmtBRL2(pf.atualDia)}</td>
          <td class="num">${fmtInt(pp.falta)}</td>
          <td class="num">${pp.necessarioDia.toFixed(1)}</td>
          <td>${paceBadge(pf)}</td>
        </tr>`;
      }).join("")}
    </tbody></table></div>
  `;
}

function renderVendedor(view, sup, vend) {
  view.innerHTML = `
    <div class="breadcrumb"><a href="#/sup/${sup.codigo}">Equipe ${esc(sup.nome)}</a> / ${esc(vend.nome)}</div>
    ${headerCards(vend.geral)}
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Fornecedor</div>
    <div class="card">${fornecedorPaceGrid(vend.geral)}</div>
    <div class="card" style="margin-top:14px">
      <div class="table-scroll"><table id="exportTable"><thead><tr>
        <th>Fornecedor</th><th class="num">Meta</th><th class="num">Projetado</th><th class="num">Falta</th>
        <th class="num">R$/dia necessário</th><th class="num">Clientes/dia necessário</th>
      </tr></thead><tbody>
        ${DATA.fornecedores.map(f => {
          const b = vend.geral.porFornecedor[f.id] || { metaFaturamento:0, faturadoLiquido:0, aFaturar:0, metaPositivacao:0, positivacaoRealizado:0 };
          const pf = paceFaturamento(b), pp = pacePositivacao(b);
          return `<tr><td><span class="forn-bar-dot" style="background:${f.cor};display:inline-block;margin-right:6px"></span>${esc(f.nome)}</td>
            <td class="num">${fmtBRL2(b.metaFaturamento)}</td>
            <td class="num">${fmtBRL2(b.faturadoLiquido + b.aFaturar)}</td>
            <td class="num">${fmtBRL2(pf.falta)}</td>
            <td class="num">${fmtBRL2(pf.necessarioDia)}</td>
            <td class="num">${pp.necessarioDia.toFixed(1)}</td>
          </tr>`;
        }).join("")}
      </tbody></table></div>
    </div>
  `;
}

function exportToExcel() {
  const table = document.getElementById("exportTable") || document.querySelector(".view table");
  if (!table) { alert("Não há tabela para exportar nesta tela."); return; }
  const rows = [...table.querySelectorAll("tr")].map(tr =>
    [...tr.children].map(td => `"${td.textContent.trim().replace(/"/g, '""')}"`).join(";")
  );
  const csv = "﻿" + rows.join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `ritmo-diario-${(document.getElementById("pageTitle").textContent||"export").toLowerCase().replace(/\s+/g,"-")}.csv`;
  a.click(); URL.revokeObjectURL(url);
}
