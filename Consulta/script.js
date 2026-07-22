// Consulta de Pedidos — busca rapida do status operacional de qualquer pedido da empresa
let DATA = null;
const MAX_RESULTADOS = 150;

const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtDate = (iso) => { const [y,m,d] = iso.split("-"); return `${d}/${m}/${y}`; };
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

// Ordem oficial do processo (definida pelo usuário): Liberado -> Montado -> Faturado ->
// Pendente (Em Aprovação) -> Devolvido -> Bloqueado. Usada tanto na régua/legenda quanto
// na ordem dos chips de resumo, para o vendedor sempre ver os status na mesma sequência.
const STATUS_ORDEM = ["LIBERADO", "MONTADO", "FATURADO", "PENDENTE", "DEVOLVIDO", "BLOQUEADO", "MISTO"];
const STATUS_LABEL = {
  LIBERADO: "Liberado", MONTADO: "Montado", FATURADO: "Faturado",
  PENDENTE: "Em Aprovação", DEVOLVIDO: "Devolvido", BLOQUEADO: "Bloqueado", MISTO: "Misto"
};
const PROCESSO_PASSOS = [
  { status: "LIBERADO", label: "Liberado" },
  { status: "MONTADO", label: "Montado" },
  { status: "FATURADO", label: "Faturado" },
  { status: "PENDENTE", label: "Em Aprovação" },
  { status: "DEVOLVIDO", label: "Devolvido" },
  { status: "BLOQUEADO", label: "Bloqueado" },
];

function legendaProcessoHtml() {
  const passos = PROCESSO_PASSOS.map((p, i) => `
    <div class="legenda-passo">
      <div class="legenda-numero status-${p.status}">${i + 1}</div>
      <div class="legenda-label">${esc(p.label)}</div>
    </div>
    ${i < PROCESSO_PASSOS.length - 1 ? `<div class="legenda-seta">→</div>` : ""}
  `).join("");
  return `
    <div class="legenda-processo">
      <div class="legenda-titulo">Ordem do processo — o que vem depois</div>
      <div class="legenda-trilha">${passos}</div>
    </div>
  `;
}

(function init() {
  if (!window.DASHBOARD_DATA) {
    document.getElementById("resultados").innerHTML = `<p class="empty-hint">data.js não encontrado. Rode gerar-dados.ps1 primeiro.</p>`;
    return;
  }
  DATA = window.DASHBOARD_DATA;
  document.getElementById("periodoLabel").textContent = DATA.periodo.label;
  document.getElementById("legendaProcesso").innerHTML = legendaProcessoHtml();
  const input = document.getElementById("buscaPedido");
  input.addEventListener("input", () => render(input.value.trim().toLowerCase()));
  render("");
})();

function statusPrincipal(pedido) {
  const distintos = [...new Set(pedido.produtos.map(p => p.statusBloqueio).filter(Boolean))];
  return distintos.length === 1 ? distintos[0] : "MISTO";
}

function render(termo) {
  const el = document.getElementById("resultados");
  if (!termo) {
    el.innerHTML = `<p class="empty-hint">Digite um número de pedido, CNPJ, código de cliente, código RCA ou nome de supervisor para ver o status.</p>`;
    return;
  }
  const matches = DATA.consultaPedidos.filter(p => {
    const alvo = [p.numeroPedidoWinthor, p.numeroPedidoClube, p.cnpj, p.codCliente, p.codigoRCA, p.supervisor, p.razaoSocial, p.vendedor]
      .join(" ").toLowerCase();
    return alvo.includes(termo);
  });
  const devMatches = (DATA.consultaDevolucoes || []).filter(d => {
    const alvo = [d.cnpj, d.codCliente, d.codigoRCA, d.supervisor, d.razaoSocial, d.vendedor].join(" ").toLowerCase();
    return alvo.includes(termo);
  });

  if (!matches.length && !devMatches.length) {
    el.innerHTML = `<p class="empty-hint">Nenhum pedido encontrado para "${esc(termo)}".</p>`;
    return;
  }
  matches.sort((a, b) => (a.razaoSocial || "").localeCompare(b.razaoSocial || ""));
  const mostrar = matches.slice(0, MAX_RESULTADOS);

  const porStatus = {};
  matches.forEach(p => { const s = statusPrincipal(p); porStatus[s] = (porStatus[s]||0) + 1; });
  const resumoHtml = STATUS_ORDEM.filter(s => porStatus[s]).map(s =>
    `<span class="resumo-chip status-${s}">${porStatus[s]} ${STATUS_LABEL[s]}</span>`
  ).join("");

  const countMsg = matches.length > MAX_RESULTADOS
    ? `${matches.length} pedidos encontrados — mostrando os ${MAX_RESULTADOS} primeiros. Refine a busca para ver os demais.`
    : `${matches.length} pedido(s) encontrado(s)`;

  const pedidosSecaoHtml = matches.length ? `
    <div class="resultados-toolbar">
      <div class="resultados-count">${countMsg}</div>
      <div class="resultados-resumo">${resumoHtml}</div>
    </div>
    <div class="pedido-list">${mostrar.map(pedidoCardHtml).join("")}</div>
  ` : `<p class="empty-hint">Nenhum pedido encontrado para "${esc(termo)}" — veja as devoluções abaixo.</p>`;

  const devSecaoHtml = devMatches.length ? `
    <div class="secao-titulo-consulta">Devoluções no período
      <span class="hint-inline">a base não vincula a devolução a um pedido específico — agrupado por cliente</span>
    </div>
    <div class="pedido-list">${devMatches.map(devolucaoCardHtml).join("")}</div>
  ` : "";

  el.innerHTML = `${pedidosSecaoHtml}${devSecaoHtml}`;
  el.querySelectorAll(".pedido-card-head").forEach(h => h.addEventListener("click", () => {
    h.parentElement.classList.toggle("open");
  }));
}

function campo(label, valor) {
  return `<div class="campo"><span class="campo-label">${esc(label)}</span><span class="campo-valor">${esc(valor)}</span></div>`;
}

function pedidoCardHtml(p) {
  const status = statusPrincipal(p);
  const temCorte = p.produtosCortados && p.produtosCortados.length;
  return `<div class="pedido-card">
    <div class="pedido-card-head">
      <span class="status-badge status-${esc(status)}">${esc(STATUS_LABEL[status] || status)}</span>
      ${temCorte ? `<span class="badge-corte" title="Este pedido teve produto(s) cortado(s)">Corte</span>` : ""}
      <div class="pedido-card-titulo">
        <div class="razao-social">${esc(p.razaoSocial)}</div>
        <div class="vendedor-linha">${esc(p.vendedor)} · Equipe ${esc(p.supervisor)}</div>
      </div>
      <div class="pedido-card-valor">${fmtBRL2(p.valorTotal)}</div>
      <div class="pedido-card-toggle">▸</div>
    </div>
    <div class="campo-grid">
      ${campo("Data", fmtDate(p.data))}
      ${campo("Pedido Winthor", p.numeroPedidoWinthor)}
      ${campo("Pedido Clube", p.numeroPedidoClube)}
      ${campo("Cód. RCA", p.codigoRCA)}
      ${campo("Cód. Cliente", p.codCliente)}
      ${campo("CNPJ", p.cnpj)}
    </div>
    <div class="pedido-card-body">
      <table><thead><tr>
        <th>Cód. Produto</th><th>Produto</th><th>Fornecedor</th><th>Status Pedido</th><th>Status</th><th class="num">Valor</th>
      </tr></thead><tbody>
        ${p.produtos.map(pr => `<tr>
          <td>${esc(pr.codProduto)}</td><td>${esc(pr.produto)}</td><td>${esc(pr.fornecedor||"—")}</td>
          <td>${pr.statusPedido === "FATURADO" ? "Faturado" : "A Faturar"}</td>
          <td><span class="status-badge status-${esc(pr.statusBloqueio)}">${esc(STATUS_LABEL[pr.statusBloqueio] || pr.statusBloqueio)}</span></td>
          <td class="num">${fmtBRL2(pr.valor)}</td>
        </tr>`).join("")}
      </tbody></table>
      ${temCorte ? `
        <div class="corte-subtitulo">Produtos com corte neste pedido</div>
        <table class="corte-subtable"><thead><tr>
          <th>Cód. Produto</th><th>Produto</th><th>Fornecedor</th><th>Embalagem</th><th class="num">Qt. Corte</th><th class="num">Preço Unit.</th><th class="num">Valor</th>
        </tr></thead><tbody>
          ${p.produtosCortados.map(pr => `<tr>
            <td>${esc(pr.codProduto)}</td><td>${esc(pr.produto)}</td><td>${esc(pr.fornecedor||"—")}</td>
            <td>${esc(pr.embalagem)}</td><td class="num">${Math.round(pr.qtCorte||0).toLocaleString("pt-BR")}</td>
            <td class="num">${fmtBRL2(pr.precoUnit)}</td><td class="num">${fmtBRL2(pr.valor)}</td>
          </tr>`).join("")}
        </tbody></table>
      ` : ""}
    </div>
  </div>`;
}

function devolucaoCardHtml(d) {
  return `<div class="pedido-card devolucao-card">
    <div class="pedido-card-head">
      <span class="status-badge status-DEVOLVIDO">Devolução</span>
      <div class="pedido-card-titulo">
        <div class="razao-social">${esc(d.razaoSocial)}</div>
        <div class="vendedor-linha">${esc(d.vendedor)} · Equipe ${esc(d.supervisor)}</div>
      </div>
      <div class="pedido-card-valor">${fmtBRL2(d.valorTotal)}</div>
      <div class="pedido-card-toggle">▸</div>
    </div>
    <div class="campo-grid">
      ${campo("Cód. RCA", d.codigoRCA)}
      ${campo("Cód. Cliente", d.codCliente)}
      ${campo("CNPJ", d.cnpj)}
    </div>
    <div class="pedido-card-body">
      <table><thead><tr>
        <th>Data</th><th>Cód. Produto</th><th>Produto</th><th>Fornecedor</th><th class="num">Valor devolvido</th>
      </tr></thead><tbody>
        ${d.produtos.map(pr => `<tr>
          <td>${fmtDate(pr.data)}</td><td>${esc(pr.codProduto)}</td><td>${esc(pr.produto)}</td><td>${esc(pr.fornecedor||"—")}</td>
          <td class="num">${fmtBRL2(pr.valor)}</td>
        </tr>`).join("")}
      </tbody></table>
    </div>
  </div>`;
}
