// Consulta de Pedidos — busca rapida do status operacional de qualquer pedido da empresa
let DATA = null;
const MAX_RESULTADOS = 150;

const fmtBRL2 = (v) => (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtDate = (iso) => { const [y,m,d] = iso.split("-"); return `${d}/${m}/${y}`; };
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

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
  const mostrar = matches.slice(0, MAX_RESULTADOS);
  const countMsg = matches.length > MAX_RESULTADOS
    ? `${matches.length} pedidos encontrados — mostrando os ${MAX_RESULTADOS} primeiros. Refine a busca para ver os demais.`
    : `${matches.length} pedido(s) encontrado(s).`;
  el.innerHTML = `<div class="resultados-count">${countMsg}</div>` + mostrar.map(pedidoCardHtml).join("");
  el.querySelectorAll(".pedido-card-head").forEach(h => h.addEventListener("click", () => {
    h.parentElement.classList.toggle("open");
  }));
}

function pedidoCardHtml(p) {
  const status = statusPrincipal(p);
  return `<div class="pedido-card">
    <div class="pedido-card-head">
      <span class="status-badge status-${esc(status)}">${esc(status)}</span>
      <div class="pedido-card-info">
        <div class="linha1">${esc(p.razaoSocial)} <span style="font-weight:400;color:var(--ink-muted)">· ${esc(p.vendedor)} (${esc(p.supervisor)})</span></div>
        <div class="linha2">
          ${fmtDate(p.data)} · Pedido Winthor ${esc(p.numeroPedidoWinthor)} · Pedido Clube ${esc(p.numeroPedidoClube)} ·
          Cód. RCA ${esc(p.codigoRCA)} · Cód. Cliente ${esc(p.codCliente)} · CNPJ ${esc(p.cnpj)}
        </div>
      </div>
      <div class="pedido-card-valor">${fmtBRL2(p.valorTotal)}</div>
      <div class="pedido-card-toggle">▸</div>
    </div>
    <div class="pedido-card-body">
      <table><thead><tr>
        <th>Cód. Produto</th><th>Produto</th><th>Fornecedor</th><th>Status Pedido</th><th>Status Bloqueio</th><th class="num">Valor</th>
      </tr></thead><tbody>
        ${p.produtos.map(pr => `<tr>
          <td>${esc(pr.codProduto)}</td><td>${esc(pr.produto)}</td><td>${esc(pr.fornecedor||"—")}</td>
          <td>${esc(pr.statusPedido)}</td>
          <td><span class="status-badge status-${esc(pr.statusBloqueio)}">${esc(pr.statusBloqueio)}</span></td>
          <td class="num">${fmtBRL2(pr.valor)}</td>
        </tr>`).join("")}
      </tbody></table>
    </div>
  </div>`;
}
