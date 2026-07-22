// Ritmo Diário — quanto cada vendedor precisa faturar e positivar por dia útil restante
// para bater a meta do mês. Reaproveita o mesmo data.json do Dashboard API (Site/).
let DATA = null;

const fmtBRL = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtInt = (v) => Math.round(v || 0).toLocaleString("pt-BR");
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
  let html = `<div class="nav-item" data-hash="#/total"><span class="nav-dot" style="background:#1f5fa8"></span>Geral API Total</div>`;
  html += `<div class="nav-section-label">Supervisores</div>`;
  html += DATA.supervisores.map(s => `<div class="nav-item" data-hash="#/sup/${s.codigo}"><span class="nav-dot" style="background:${s.cor}"></span>${esc(s.nome)}</div>`).join("");
  nav.innerHTML = html;
  nav.querySelectorAll(".nav-item").forEach(el => el.addEventListener("click", () => {
    location.hash = el.dataset.hash; document.getElementById("sidebar").classList.remove("open");
  }));
  document.getElementById("menuToggle").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
  document.getElementById("btnBack").addEventListener("click", () => history.back());
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
// Último dia útil do mês vigente (segunda a sábado) — usado como prazo de referência
// ---------------------------------------------------------------------------
function ultimoDiaUtilLabel() {
  const ano = DATA.periodo.ano, mes = DATA.periodo.mes;
  const d = new Date(ano, mes, 0); // dia 0 do mes seguinte = ultimo dia do mes vigente
  while (d.getDay() === 0) d.setDate(d.getDate() - 1); // domingo nao conta como dia util
  const dias = ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"];
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dias[d.getDay()]}, ${dd}/${mm}`;
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
  return { falta, necessarioDia, atualDia, bateu: falta <= 0, proj };
}
function pacePositivacao(agg) {
  const falta = Math.max(0, agg.metaPositivacao - agg.positivacaoRealizado);
  const dRest = DATA.periodo.diasUteisRestantes;
  const dPass = DATA.periodo.diasUteisPassados;
  // Positivação é sempre número inteiro de clientes: arredonda pra cima o "necessário"
  // (senão o ritmo fica sistematicamente abaixo do que precisa pra bater a meta) e
  // pra mais próximo o "atual" (é só a média já realizada, nao precisa de folga).
  const necessarioDia = Math.ceil(dRest > 0 ? falta / dRest : falta);
  const atualDia = Math.round(dPass > 0 ? agg.positivacaoRealizado / dPass : agg.positivacaoRealizado);
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

// 1) "O que falta" — o número mais importante, sempre no topo.
function heroFalta(agg) {
  const pf = paceFaturamento(agg);
  const pp = pacePositivacao(agg);
  const prazo = ultimoDiaUtilLabel();
  return `
    <div class="hero-falta">
      <div class="hero-falta-titulo">O que falta para bater a meta <span class="hero-falta-prazo">— até ${prazo}, último dia útil do mês</span></div>
      <div class="hero-falta-grid">
        <div class="hero-falta-item ${pf.bateu ? "ok" : ""}">
          <div class="hero-falta-label">Falta faturar</div>
          <div class="hero-falta-valor">${pf.bateu ? "Meta batida ✓" : fmtBRL(pf.falta)}</div>
          <div class="hero-falta-sub">Meta de faturamento do mês: ${fmtBRL(agg.metaFaturamento)}</div>
        </div>
        <div class="hero-falta-item ${pp.bateu ? "ok" : ""}">
          <div class="hero-falta-label">Falta positivar</div>
          <div class="hero-falta-valor">${pp.bateu ? "Meta batida ✓" : `${fmtInt(pp.falta)} clientes`}</div>
          <div class="hero-falta-sub">Meta de positivação do mês: ${fmtInt(agg.metaPositivacao)} clientes</div>
        </div>
      </div>
    </div>
  `;
}

// 2) "Projetado" — onde estamos agora: o que já foi faturado + o que ainda está a faturar,
// comparado com a meta. Positivação usa uma barra simples (realizado vs meta) porque um
// mesmo cliente pode ter pedido faturado E pedido a faturar ao mesmo tempo — empilhar as
// duas contagens somaria esse cliente duas vezes e daria um total errado.
function barraComMarcador(fillsHtml, pctMeta) {
  return `
    <div class="projetado-bar-track">
      ${fillsHtml}
      <div class="projetado-meta-marker" style="left:${pctMeta}%"></div>
    </div>
  `;
}
function projetadoFaturamento(agg) {
  const atual = agg.faturadoLiquido, pendente = agg.aFaturar, meta = agg.metaFaturamento;
  // +8% de folga na escala pra marcador de meta nunca cair exatamente na borda direita
  // (senao o rotulo "Meta", centralizado no marcador, fica cortado/colado no texto do cabecalho)
  const escala = (Math.max(atual + pendente, meta) || 1) * 1.08;
  const pctMeta = clampPct(meta / escala * 100);
  const fills = `
    <div class="projetado-bar-fill faturado" style="flex:${Math.max(atual, 0.0001)} 0 0"></div>
    <div class="projetado-bar-fill afaturar" style="flex:${Math.max(pendente, 0.0001)} 0 0"></div>
  `;
  return `
    <div class="projetado-row">
      <div class="projetado-row-head"><span>Faturamento</span><b>${fmtBRL2(atual + pendente)} <span class="projetado-vs-meta">projetado / meta ${fmtBRL2(meta)}</span></b></div>
      ${barraComMarcador(fills, pctMeta)}
      <div class="projetado-legenda">
        <span><i class="dot faturado"></i>Já faturado: ${fmtBRL2(atual)}</span>
        <span><i class="dot afaturar"></i>Ainda a faturar: ${fmtBRL2(pendente)}</span>
      </div>
    </div>
  `;
}
function projetadoPositivacao(agg) {
  const atual = agg.positivacaoRealizado, meta = agg.metaPositivacao;
  const escala = (Math.max(atual, meta) || 1) * 1.08;
  const pctMeta = clampPct(meta / escala * 100);
  const fills = `<div class="projetado-bar-fill realizado" style="flex:${Math.max(atual, 0.0001)} 0 0"></div>`;
  return `
    <div class="projetado-row">
      <div class="projetado-row-head"><span>Positivação</span><b>${fmtInt(atual)} clientes <span class="projetado-vs-meta">projetado / meta ${fmtInt(meta)}</span></b></div>
      ${barraComMarcador(fills, pctMeta)}
      <div class="projetado-legenda">
        <span><i class="dot realizado"></i>Clientes positivados: ${fmtInt(atual)} (${fmtInt(agg.positivacaoFaturado)} com nota já faturada, ${fmtInt(agg.positivacaoAFaturar)} com pedido ainda a faturar)</span>
      </div>
    </div>
  `;
}
function projetadoSection(agg) {
  return `
    <div class="section-title"><span class="bar"></span>Projetado — como estamos hoje</div>
    <div class="card">
      ${projetadoFaturamento(agg)}
      ${projetadoPositivacao(agg)}
    </div>
  `;
}

// 3) "Ritmo diário necessário" — quanto precisa fazer por dia útil (seg a sáb) pra
// fechar a diferença até o fim do mês.
function ritmoDiarioSection(agg) {
  const pf = paceFaturamento(agg);
  const pp = pacePositivacao(agg);
  return `
    <div class="section-title"><span class="bar"></span>Ritmo diário necessário <span class="section-sub">(segunda a sábado · restam ${DATA.periodo.diasUteisRestantes} dias úteis)</span></div>
    <div class="kpi-grid">
      <div class="kpi ${pf.bateu ? "good" : ""}">
        <div class="kpi-label">Faturamento necessário / dia útil</div>
        <div class="kpi-value">${fmtBRL(pf.necessarioDia)}</div>
        <div class="kpi-sub">Média conquistada até agora: ${fmtBRL(pf.atualDia)}/dia útil — ${paceBadge(pf)}</div>
      </div>
      <div class="kpi ${pp.bateu ? "good" : ""}">
        <div class="kpi-label">Clientes a positivar / dia útil</div>
        <div class="kpi-value">${fmtInt(pp.necessarioDia)}</div>
        <div class="kpi-sub">Média conquistada até agora: ${fmtInt(pp.atualDia)}/dia útil — ${paceBadge(pp)}</div>
      </div>
    </div>
  `;
}

function headerCards(agg) {
  return `
    ${heroFalta(agg)}
    ${projetadoSection(agg)}
    ${ritmoDiarioSection(agg)}
  `;
}

function fornecedorPaceGrid(agg) {
  const cards = DATA.fornecedores.map(f => {
    const b = agg.porFornecedor[f.id] || { metaFaturamento:0, faturadoLiquido:0, aFaturar:0, metaPositivacao:0, positivacaoRealizado:0 };
    const pf = paceFaturamento(b);
    const pp = pacePositivacao(b);
    return `<div class="pace-card">
      <div class="pf-name"><span class="pf-dot" style="background:${f.cor}"></span>${esc(f.nome)}</div>
      <div class="pace-row"><span>Meta de faturamento</span><b>${fmtBRL(b.metaFaturamento)}</b></div>
      <div class="pace-row"><span>Projetado (faturado + a faturar)</span><b>${fmtBRL(b.faturadoLiquido + b.aFaturar)}</b></div>
      <div class="pace-row falta"><span>Falta p/ bater meta</span><b>${pf.bateu ? "Meta batida ✓" : fmtBRL(pf.falta)}</b></div>
      <div class="pace-need">${fmtBRL(pf.necessarioDia)}<span class="pace-need-sub"> /dia útil necessário</span></div>
      <div class="pace-row"><span>Positivação necessária</span><b>${fmtInt(pp.necessarioDia)} clientes/dia útil</b></div>
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
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Indústria</div>
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
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Indústria — equipe ${esc(sup.nome)}</div>
    <div class="card">${fornecedorPaceGrid(sup.geral)}</div>
    <div class="section-title"><span class="bar"></span>Vendedores — ritmo individual (${vendedores.length})</div>
    <div class="table-scroll"><table id="exportTable"><thead><tr>
      <th>Vendedor</th><th class="num">Falta p/ bater meta (R$)</th><th class="num">R$ necessário/dia útil</th>
      <th class="num">R$ média atual/dia útil</th><th class="num">Falta positivar (clientes)</th><th class="num">Clientes necessário/dia útil</th><th>Situação</th>
    </tr></thead><tbody>
      ${vendedores.map(v => {
        const pf = paceFaturamento(v.geral), pp = pacePositivacao(v.geral);
        return `<tr style="cursor:pointer" onclick="location.hash='#/sup/${sup.codigo}/vendedor/${v.codigo}'">
          <td><strong>${esc(v.nome)}</strong></td>
          <td class="num">${pf.bateu ? "—" : fmtBRL2(pf.falta)}</td>
          <td class="num">${fmtBRL2(pf.necessarioDia)}</td>
          <td class="num">${fmtBRL2(pf.atualDia)}</td>
          <td class="num">${pp.bateu ? "—" : fmtInt(pp.falta)}</td>
          <td class="num">${fmtInt(pp.necessarioDia)}</td>
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
    <div class="section-title"><span class="bar"></span>Ritmo necessário por Indústria</div>
    <div class="card">${fornecedorPaceGrid(vend.geral)}</div>
    <div class="card" style="margin-top:14px">
      <div class="table-scroll"><table id="exportTable"><thead><tr>
        <th>Indústria</th><th class="num">Meta</th><th class="num">Projetado</th><th class="num">Falta p/ bater meta</th>
        <th class="num">R$ necessário/dia útil</th><th class="num">Clientes necessário/dia útil</th>
      </tr></thead><tbody>
        ${DATA.fornecedores.map(f => {
          const b = vend.geral.porFornecedor[f.id] || { metaFaturamento:0, faturadoLiquido:0, aFaturar:0, metaPositivacao:0, positivacaoRealizado:0 };
          const pf = paceFaturamento(b), pp = pacePositivacao(b);
          return `<tr><td><span class="forn-bar-dot" style="background:${f.cor};display:inline-block;margin-right:6px"></span>${esc(f.nome)}</td>
            <td class="num">${fmtBRL2(b.metaFaturamento)}</td>
            <td class="num">${fmtBRL2(b.faturadoLiquido + b.aFaturar)}</td>
            <td class="num">${pf.bateu ? "—" : fmtBRL2(pf.falta)}</td>
            <td class="num">${fmtBRL2(pf.necessarioDia)}</td>
            <td class="num">${fmtInt(pp.necessarioDia)}</td>
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
