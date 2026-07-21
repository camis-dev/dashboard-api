// Consulta de Pedidos — busca rapida do status operacional de qualquer pedido da empresa
let DATA = null;
const MAX_RESULTADOS = 150;

const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtDate = (iso) => { const [y,m,d] = iso.split("-"); return `${d}/${m}/${y}`; };
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

const STATUS_ORDEM = ["BLOQUEADO", "PENDENTE", "MONTADO", "LIBERADO", "FATURADO", "DEVOLVIDO", "MISTO"];
const STATUS_LABEL = {
  BLOQUEADO: "Bloqueado", PENDENTE: "Pendente", MONTADO: "Montado",
  LIBERADO: "Liberado", FATURADO: "Faturado", DEVOLVIDO: "Devolvido", MISTO: "Misto"
};

(function init() {
  if (!window.DASHBOARD_DATA) {
    document.getElementById("resultados").innerHTML = `<p class="empty-hint">data.js não encontrado. Rode gerar-dados.ps1 primeiro.</p>`;
    return;
  }
  DATA = window.DASHBOARD_DATA;
  document.getElementById("periodoLabel").textContent = DATA.periodo.label;
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
  if (!matches.length) {
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

  el.innerHTML = `
    <div class="resultados-toolbar">
      <div class="resultados-count">${countMsg}</div>
      <div class="resultados-resumo">${resumoHtml}</div>
    </div>
    <div class="pedido-list">${mostrar.map(pedidoCardHtml).join("")}</div>
  `;
  el.querySelectorAll(".pedido-card-head").forEach(h => h.addEventListener("click", () => {
    h.parentElement.classList.toggle("open");
  }));
}

function campo(label, valor) {
  return `<div class="campo"><span class="campo-label">${esc(label)}</span><span class="campo-valor">${esc(valor)}</span></div>`;
}

function pedidoCardHtml(p) {
  const status = statusPrincipal(p);
  return `<div class="pedido-card">
    <div class="pedido-card-head">
      <span class="status-badge status-${esc(status)}">${esc(STATUS_LABEL[status] || status)}</span>
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
    </div>
  </div>`;
}
