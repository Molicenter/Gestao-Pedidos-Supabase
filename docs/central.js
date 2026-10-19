// ============================================================
// Gestão de Pedidos · CENTRAL (Administrador e Ceasa) — etapa 2
// Separação e Fechamento · Visão Fornecedores · Pedido por Fornecedor ·
// Visão das Lojas (ver como) · Log de Pedidos · Limpar Pedidos
// Mesmas regras do tela_pedidos_generica.py. Usa as funções do app.js.
// ============================================================
const SETORES_CEASA = ["Folhagem", "FLV Normal", "FLV Ofertas"];
const FATOR_ALERTA_CENTRAL = 3;       // FATOR_ALERTA_MEDIA do Streamlit
const ehCentral = () => !!sessao && sessao.perfil !== "loja";
// Ceasa: só FLV/Folhagem · responsável de setor (Administrador com lista de setores): só os setores dele
const setoresVisiveis = () => {
  if (sessao?.perfil === "ceasa") return SETORES.filter((s) => SETORES_CEASA.includes(s.nome));
  if (Array.isArray(sessao?.setores) && sessao.setores.length) return SETORES.filter((s) => sessao.setores.includes(s.nome));
  return SETORES;
};
const adminTotal = () => sessao?.perfil === "admin" && !(Array.isArray(sessao.setores) && sessao.setores.length);
const ehFLV = (nome) => nome === "FLV Normal" || nome === "FLV Ofertas";   // Cód. Iceasa, preço, BOX, alerta 3x
const ABAS = {
  separacao: "Separação e Fechamento",
  porfornecedor: "Pedido por Fornecedor",
  fornecedores: "Visão Fornecedores",
  loja: "Visão das Lojas",
  log: "Log de Pedidos",
  historico: "Histórico",
  catalogo: "Catálogo de Produtos",
};
// FLV Normal/Ofertas abrem na Separação; os demais na Visão Fornecedores (igual ao Streamlit)
const abasDoSetor = (s) => (ehFLV(s.nome) ? ["separacao", "porfornecedor", "fornecedores", "loja", "log", "historico", "catalogo"] : ["fornecedores", "separacao", "loja", "log", "historico", "catalogo"]);

// 🚚 mapa fixo fornecedor → produtos (por Cód. Iceasa), só para o Pedido por Fornecedor do FLV
const MAPA_FORNECEDORES_FLV = {
  "NIDE": [45, 49, 67, 57, 46, 48, 47], "Claudir Mendes": [57, 46, 49], "SANDRO": [75],
  "DENIZE": [45, 49, 67, 57, 46, 48, 47, 41, 52, 69, 56], "JOVANO": [1746, 88, 49, 140, 85],
  "JEFINHO": [85, 140, 256, 267, 88, 57, 45, 49, 1662, 46], "LUCIANO": [61, 41, 49, 56, 45],
  "THIAGO": [61, 91, 67, 74, 49, 52, 45, 56], "CRISTIAN": [40, 949, 42, 83, 68, 538, 78],
  "ROGERIO NARANTE": [538], "FERNANDO NARANTE": [46], "SILVIO MAND SALSA": [76],
  "HORTA": [108, 109], "GLAUCIA MACIEL": [84, 85], "ALEMÃO": [39], "RENAN SS": [72],
  "NEGUIN": [85, 86, 88, 61, 45], "RODRIGO CHANAN": [85, 86, 88, 61, 1662, 140],
  "MARCELO MORANGO": [58], "JOÃO BATISTA": [79, 60, 56, 1662, 69], "GIACOMELLO": [95],
  "PRIMO": [240, 86, 49, 45, 88, 85], "RENATO MANDIOCA": [75], "THIAGO SERRA": [91, 49, 45, 56, 61],
  "TICO": [236, 237, 707, 2730, 42, 581, 78, 546, 80, 83, 540, 949, 40, 110, 68, 109],
  "ALGACIR": [1516, 53], "MAURICIO": [62], "PAULO IGASHIBAHI": [47, 48],
  "GILSOM BATATA": [508, 551], "DORI BATATA": [508, 551], "BANANA SANTOME": [2567, 2569, 2568],
  "MELANCIA CARLIN": [1], "MELANCIA MARCINHO": [673, 1, 3003], "RODRIGO BATATA": [508],
};
const FORNECEDORES_LINHA_FLV = new Set(["BANANA SANTOME", "MELANCIA CARLIN", "MELANCIA MARCINHO", "RODRIGO BATATA"]);

// Iceasa na impressão/BOX: vazio se ausente ou acima de 9000
function iceasaImpressao(v) {
  const n = Math.trunc(Number(String(v ?? "").trim()));
  if (!String(v ?? "").trim() || !Number.isFinite(n)) return "";
  return n > 9000 ? "" : String(n);
}
const precoCelula = (v) => (v === null || v === undefined || v === "" || !Number.isFinite(Number(v)) ? "" : "R$ " + Number(v).toFixed(2).replace(".", ","));
function celulaPreco(s) {
  const t = String(s ?? "").replace("R$", "").replace(/\s/g, "").replace(",", ".");
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
}

// ============================================================
// Estado
// ============================================================
let centralPainel = null;   // pedidos_web_central_painel
let cd = null;              // pedidos_web_central_setor
let cR = null;              // regras do setor aberto
let cSetor = null;          // item de SETORES
let abaAtual = null;
let cVal = {}, cOrig = {};  // "codigo|loja" → valor
let cExt = {}, cExtOrig = {}; // codigo → { p: "R$ 12,50", o: "" }
let cMedia = {};            // "codigo|loja" → média
let cBloq = {};             // codigo → Set(lojas não liberadas)
let cFiltro = "Todos";
let pf = null;              // Pedido por Fornecedor: [{ nome, cods: [], linha }]

const chave = (c, l) => `${c}|${l}`;
const normV = (v) => (cR?.texto ? String(v ?? "").trim() : (qtdInt(v) > 0 ? String(qtdInt(v)) : ""));
function celulasSujas() { return Object.keys(cVal).filter((k) => normV(cVal[k]) !== normV(cOrig[k])); }
function extrasSujos() {
  return Object.keys(cExt).filter((c) => {
    const a = cExt[c], b = cExtOrig[c] || { p: "", o: "" };
    return celulaPreco(a.p) !== celulaPreco(b.p) || String(a.o || "").trim() !== String(b.o || "").trim();
  });
}
function centralSujo() { return (!!cd && (celulasSujas().length > 0 || extrasSujos().length > 0)) || catalogoSujo(); }

// ============================================================
// Rotas da central
// ============================================================
async function rotaCentral(s, aba) {
  if (s && !setoresVisiveis().includes(s)) s = null;
  // saindo de uma digitação (aba Visão das Lojas)
  if (telaAtual === "setor" && !(s && aba === "loja" && setorAberto?.nome === s.nome)) {
    if (!(await podeSairDoSetor())) return;
  }
  // voltando da Visão das Lojas: recarrega a central (a loja pode ter mudado o pedido)
  if (telaAtual === "setor" && !centralSujo()) cd = null;
  // saindo do setor com ajustes não salvos
  const setorSujo = cd?.setor || cat?.setor;
  const saindoDoCatalogo = catalogoSujo() && (aba !== "catalogo" || !s || s.nome !== cat.setor);
  if ((centralSujo() && (!s || s.nome !== setorSujo)) || saindoDoCatalogo) {
    if (!confirm("Há alterações não salvas. Sair sem salvar?")) {
      history.replaceState(null, "", `#/setor/${slug(setorSujo)}/${abaAtual || ""}`);
      if (cSetor) renderAbas(cSetor, abaAtual);
      return;
    }
    cd = null; cat = null;
  }
  $("#pil-loja").hidden = !(s && aba === "loja");
  if (!s) { $("#abas").hidden = true; return abrirCentral(); }
  if (!abasDoSetor(s).includes(aba)) aba = abasDoSetor(s)[0];
  abaAtual = aba;
  renderAbas(s, aba);
  if (aba === "loja") { $("#pil-loja").hidden = false; return abrirSetor(s); }
  return abrirSetorCentral(s, aba);
}
function renderAbas(s, aba) {
  const el = $("#abas");
  el.hidden = false;
  el.innerHTML = `<div class="abas-lista">${abasDoSetor(s).map((a) =>
    `<a href="#/setor/${slug(s.nome)}/${a}" class="${a === aba ? "ativa" : ""}">${ABAS[a]}</a>`).join("")}</div>
    <button class="btn btn-perigo btn-pequeno" type="button" id="btn-limpar">Limpar pedidos</button>`;
  $("#btn-limpar").onclick = () => confirmarLimpar(s);
}
function renderNavCentral() {
  const ativo = (telaAtual === "setor" ? setorAberto?.nome : telaAtual === "central" && cSetor ? cSetor.nome : null);
  // (telaAtual "usuarios" = tela de Usuários, só do Administrador)
  const icPainel = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>';
  $("#nav").innerHTML = `<div class="nav-rotulo">Central</div>
    <button type="button" data-ir="painel" class="${telaAtual === "painel" ? "ativo" : ""}">${icPainel}Painel da central</button>
    <div class="nav-rotulo">Setores</div>
    ${setoresVisiveis().map((s) => {
      const n = lojasProntas(s.nome).length;
      const mini = n === 8 ? '<span class="estado-mini ok">8/8</span>' : n ? `<span class="estado-mini rasc">${n}/8</span>` : "";
      return `<button type="button" data-setor="${slug(s.nome)}" class="${ativo === s.nome ? "ativo" : ""}"><span class="bolinha" style="background:var(--g-${s.g})"></span>${esc(s.titulo)}${mini}</button>`;
    }).join("")}
    ${adminTotal() ? `<div class="nav-rotulo">Administração</div>
    <button type="button" data-ir="usuarios" class="${telaAtual === "usuarios" ? "ativo" : ""}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="9" cy="8" r="4"/><path d="M2 21a7 7 0 0 1 14 0M16 4a4 4 0 0 1 0 8M22 21a7 7 0 0 0-4-6.3"/></svg>Usuários</button>` : ""}`;
}

// lojas que já salvaram (confirmado) ou declararam sem pedido hoje
function lojasProntas(nome) {
  const st = centralPainel?.setores?.[nome] || {};
  return [...new Set([...(st.confirmadas || []), ...(centralPainel?.sem_pedido?.[nome] || [])])];
}

// ============================================================
// Painel da central
// ============================================================
async function abrirCentral() {
  telaAtual = "painel"; cSetor = null;
  $("#barra-total").hidden = true; document.body.classList.remove("com-barra");
  $("#pg-titulo").textContent = "Painel da central";
  $("#pg-sub").textContent = "Situação dos pedidos de todas as lojas, por setor. Clique no setor para abrir a Separação.";
  if (!centralPainel) $("#conteudo").innerHTML = '<div class="vazio">Carregando...</div>';
  try {
    centralPainel = await rpc("pedidos_web_central_painel", { p_token: sessao.token });
  } catch (err) {
    toast(err.message, "erro");
    if (!centralPainel) return;
  }
  renderNavCentral();
  // situação da rotina do servidor (estoque/médias); se o SQL v1.5 ainda não rodou, só não mostra
  rpc("pedidos_web_monitor", { p_token: sessao.token }).then((m) => {
    const el = $("#monitor-erp"); if (!el || !m?.problemas) return;
    el.innerHTML = `<div class="faixa erro monitor-erp"><div><b>⚠️ A rotina do ERP no servidor parou de atualizar.</b>
      <ul>${m.problemas.split("\n").map((l) => `<li>${esc(l)}</li>`).join("")}</ul>
      <p class="dica">Confira a tarefa “Molicenter - Estoque ERP” no Agendador de Tarefas do servidor 192.168.10.69. Até voltar, estoque e médias nas telas são da última atualização que deu certo.</p></div></div>`;
  }).catch(() => {});
  const vis = setoresVisiveis();
  const hoje = vis.map((s) => ({ s, pz: prazoHoje(s) })).filter((x) => x.pz && !x.pz.passou).sort((a, b) => a.pz.min - b.pz.min);
  const rasc = vis.filter((s) => (centralPainel.setores?.[s.nome]?.rascunho || []).length);
  const comPedido = vis.filter((s) => (centralPainel.setores?.[s.nome]?.itens || 0) > 0);
  const med = centralPainel.medias;
  $("#conteudo").innerHTML = `<div class="pagina">
    <div id="monitor-erp"></div>
    <div class="kpis">
      <div class="kpi destaque"><small>Prazos hoje</small><strong class="num">${hoje.length}</strong><span>${hoje.length ? esc(hoje.map((x) => x.s.titulo).join(", ")) : "nenhum setor fecha hoje"}</span></div>
      <div class="kpi verde"><small>Setores com pedido</small><strong class="num">${comPedido.length}</strong><span>de ${vis.length} setores</span></div>
      <div class="kpi laranja"><small>Lojas digitando</small><strong class="num">${rasc.reduce((n, s) => n + centralPainel.setores[s.nome].rascunho.length, 0)}</strong><span>${rasc.length ? "em " + esc(rasc.map((s) => s.titulo).join(", ")) : "ninguém com rascunho aberto"}</span></div>
      <div class="kpi magenta"><small>Próximo prazo</small><strong class="num">${hoje[0] ? hoje[0].s.limites.find((l) => l.d.includes(agoraSP().getDay())).h : "—"}</strong><span>${hoje[0] ? esc(hoje[0].s.titulo) + " · " + esc(hoje[0].pz.txt.toLowerCase()) : "sem prazo pela frente hoje"}</span></div>
    </div>
    <p class="aviso-estoque">${centralPainel.estoque_em ? `Estoque do ERP atualizado em ${esc(centralPainel.estoque_em)}. ` : ""}${med ? `Médias de venda (90 dias) atualizadas em ${esc(med.em)}${med.ok ? "" : " — <b>houve falha em algum setor</b>"}.` : "Médias de venda: ainda sem atualização automática registrada."}</p>
    ${GRUPOS.map((g) => {
      const doGrupo = vis.filter((s) => s.g === g.id);
      if (!doGrupo.length) return "";
      return `<section class="grupo" style="--cor:var(--g-${g.id})">
      <h3 class="grupo-titulo" style="color:var(--g-${g.id}-texto)">${g.nome}</h3>
      <div class="cards">${doGrupo.map((s) => {
        const st = centralPainel.setores?.[s.nome] || {};
        const sem = centralPainel.sem_pedido?.[s.nome] || [];
        const prontas = lojasProntas(s.nome);
        const pz = prazoHoje(s);
        const bolas = LOJAS.map((l, i) => {
          const n = i + 1;
          const cls = sem.includes(n) ? "sem" : (st.confirmadas || []).includes(n) ? "ok" : (st.rascunho || []).includes(n) ? "rasc" : "";
          return `<i class="${cls}" title="${l}">${p2(n)}</i>`;
        }).join("");
        return `<button class="setor" type="button" data-setor="${slug(s.nome)}" style="--cor:var(--g-${s.g});--cor-fundo:var(--g-${s.g}-fundo);--cor-texto:var(--g-${s.g}-texto)">
          <span class="setor-faixa"></span>
          <span class="setor-corpo">
            <span class="setor-nome">${esc(s.titulo)}</span>
            <span class="setor-prazo">${esc(s.prazo)}</span>
            <span class="lojas-mini">${bolas}</span>
            <span class="setor-rodape"><span class="tag ${prontas.length === 8 ? "ok" : prontas.length ? "rasc" : "novo"}">${prontas.length} de 8 lojas${st.itens ? " · " + st.itens + " itens" : ""}</span>${pz ? `<span class="relogio ${pz.passou ? "passou" : ""}">${esc(pz.txt)}</span>` : ""}</span>
          </span>
        </button>`;
      }).join("")}</div></section>`;
    }).join("")}
    <p class="legenda-lojas"><span><i class="ok"></i>salvou</span><span><i class="sem"></i>sem pedido hoje</span><span><i class="rasc"></i>digitando (não salvou)</span><span><i></i>não começou</span></p>
  </div>`;
}

// ============================================================
// Abrir um setor na central (carrega uma vez; as abas compartilham os ajustes)
// ============================================================
async function abrirSetorCentral(s, aba, recarregar = false) {
  telaAtual = "central"; cSetor = s; cR = regras(s.nome);
  $("#pg-titulo").textContent = `${s.titulo} · ${ABAS[aba]}`;
  $("#pg-sub").textContent = {
    separacao: "As 8 lojas lado a lado. Você pode corrigir quantidades de qualquer loja e clicar em Salvar ajustes.",
    fornecedores: "Pedido consolidado por fornecedor, com as lojas e o total.",
    porfornecedor: "Quantidades das lojas e preços da Separação, montados por fornecedor para mandar no WhatsApp.",
    log: "Quando cada loja confirmou o pedido, e quando o Administrador salvou por uma loja.",
    historico: "O pedido final de cada ciclo (guardado ao Limpar pedidos), comparado com o pedido atual e com a média de venda.",
    catalogo: "Produtos do setor, nomes, códigos e em quais lojas cada produto pode ser pedido.",
  }[aba] || "";
  renderNavCentral();
  if (aba === "log") { esconderBarraCentral(); return renderLog(); }
  if (aba === "historico") { esconderBarraCentral(); return renderHistorico(s); }
  if (aba === "catalogo") return abrirCatalogo(s, recarregar);
  if (!cd || cd.setor !== s.nome || recarregar) {
    $("#conteudo").innerHTML = '<div class="vazio">Carregando pedidos das lojas...</div>';
    esconderBarraCentral();
    try {
      const d = await rpc("pedidos_web_central_setor", { p_token: sessao.token, p_setor: s.nome });
      prepararDados(d);
    } catch (err) {
      $("#conteudo").innerHTML = `<div class="cartao"><p class="faixa erro">${esc(err.message)}</p></div>`;
      return;
    }
  }
  if (aba === "separacao") renderSeparacao();
  else if (aba === "fornecedores") renderFornecedores();
  else if (aba === "porfornecedor") renderPorFornecedor();
}
function prepararDados(d) {
  cd = d; cVal = {}; cOrig = {}; cExt = {}; cExtOrig = {}; cMedia = {}; cBloq = {}; cFiltro = "Todos"; pf = null;
  for (const p of d.produtos) cBloq[p.c] = new Set(p.x || []);
  for (const r of d.pedido) {
    const v = cR.texto ? String(r.t || "").trim() : (Number(r.q) > 0 ? String(Math.trunc(Number(r.q))) : "");
    cVal[chave(r.c, r.l)] = v; cOrig[chave(r.c, r.l)] = v;
  }
  for (const e of d.extras) {
    const x = { p: precoCelula(e.p), o: e.o || "" };
    cExt[e.c] = { ...x }; cExtOrig[e.c] = { ...x };
  }
  for (const [c, l, m] of d.medias) cMedia[chave(c, l)] = Number(m) || 0;
  const cmp = (a, b) => String(a).localeCompare(String(b), "pt-BR", { sensitivity: "base" });
  d.produtos.sort((a, b) => (cR.flores ? 0 : cmp(a.f, b.f)) || cmp(a.d, b.d));
  // só produtos liberados em pelo menos uma loja
  d.ativos = d.produtos.filter((p) => cBloq[p.c].size < 8);
}
const valor = (c, l) => cVal[chave(c, l)] ?? "";
const bloqueado = (c, l) => cBloq[c]?.has(l);
function totalLinha(p) {
  if (cR.texto) return LOJAS.reduce((n, _, i) => n + (!bloqueado(p.c, i + 1) && String(valor(p.c, i + 1)).trim() ? 1 : 0), 0);
  return LOJAS.reduce((n, _, i) => n + (bloqueado(p.c, i + 1) ? 0 : qtdInt(valor(p.c, i + 1))), 0);
}
function ordenarGrupos(nomes) {
  const mapa = cd.ordem || {};
  return [...new Set(nomes.filter((x) => String(x || "").trim()))].sort((a, b) => ((mapa[a] ?? 1e6) - (mapa[b] ?? 1e6)) || a.toLowerCase().localeCompare(b.toLowerCase(), "pt-BR"));
}

// ---------- status das 8 lojas (verde só depois de Salvar, ou Sem Pedido) ----------
function statusLojasHtml() {
  const conf = new Set(cd.pedido.filter((r) => r.ok !== false).map((r) => r.l));
  const rasc = new Set(cd.pedido.filter((r) => r.ok === false).map((r) => r.l));
  const sem = new Set(cd.sem_pedido || []);
  return `<div class="cartao"><h3>Status de digitação das lojas <small>${dataHoraBR()}</small></h3>
    <div class="status-lojas">${LOJAS.map((l, i) => {
      const n = i + 1;
      const [cls, txt] = conf.has(n) ? ["ok", "OK"] : sem.has(n) ? ["sem", "Sem pedido"] : rasc.has(n) ? ["rasc", "Digitando"] : ["falta", "Faltando"];
      return `<div class="st-loja ${cls}"><b>${l}</b><small>${txt}</small></div>`;
    }).join("")}</div></div>`;
}
function obsLojasHtml() {
  if (!cR.obs) return "";
  const lista = cd.obs_lojas || [];
  return `<div class="cartao"><h3>Observações das lojas <small>${dataHoraBR()}</small></h3>
    ${lista.length ? `<div class="tabela-wrap"><table class="tabela-simples"><thead><tr><th>Loja</th><th>Observação</th></tr></thead><tbody>
      ${lista.map((o) => `<tr><td class="num">${nomeLoja(o.l)}</td><td class="pre">${esc(o.o)}</td></tr>`).join("")}</tbody></table></div>`
      : '<p class="dica">Nenhuma loja registrou observação.</p>'}</div>`;
}

// ============================================================
// Separação e Fechamento
// ============================================================
function linhasSeparacao() {
  const q = norm($("#c-busca")?.value);
  return cd.ativos.filter((p) => (cFiltro === "Todos" || grupoFornecedor(p.f) === cFiltro)
    && (!q || norm(`${p.d} ${p.erp} ${p.ice ?? ""} ${p.b}`).includes(q)));
}
function alertasSeparacao(linhas) {
  if (!ehFLV(cd.setor) || cR.texto) return [];
  const out = [];
  for (const p of linhas) for (let l = 1; l <= 8; l++) {
    const q = qtdInt(valor(p.c, l)); const m = cMedia[chave(p.c, l)] || 0;
    if (q > 0 && m > 0 && q > FATOR_ALERTA_CENTRAL * m && !bloqueado(p.c, l)) out.push({ p, l, q, m, x: q / m });
  }
  return out.sort((a, b) => b.x - a.x);
}
function renderSeparacao() {
  const grupos = ordenarGrupos(cd.ativos.map((p) => grupoFornecedor(p.f)));
  const flv = ehFLV(cd.setor), mp = norm(cd.setor) === "materia prima";
  $("#conteudo").innerHTML = `<div class="pagina">
    ${statusLojasHtml()}
    <div class="cartao">
      <div class="barra-filtros">
        ${grupos.length > 1 && !cR.flores ? `<div class="chips" id="c-chips"><button type="button" class="sel" data-f="Todos">Todos</button>${grupos.map((g) => `<button type="button" data-f="${esc(g)}">${esc(g)}</button>`).join("")}</div>` : ""}
        <input class="busca" id="c-busca" type="search" placeholder="Buscar produto ou código" autocomplete="off">
        <span class="metrica" id="c-metrica"></span>
      </div>
      <details class="faixa erro alerta-curva" id="c-alertas" hidden></details>
      <div class="tabela-wrap grade-wrap"><table class="grade" id="c-grade"></table></div>
    </div>
    ${obsLojasHtml()}
  </div>`;
  $("#c-chips")?.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-f]"); if (!b) return;
    cFiltro = b.dataset.f;
    $$("#c-chips button").forEach((x) => x.classList.toggle("sel", x === b));
    desenharGradeSeparacao();
  });
  $("#c-busca").addEventListener("input", desenharGradeSeparacao);
  ligarEdicao($("#c-grade"));
  desenharGradeSeparacao();
  barraCentral("Separação e Fechamento");
}
function colunasSeparacao(temAlerta) {
  const flv = ehFLV(cd.setor), mp = norm(cd.setor) === "materia prima";
  const c = [];
  c.push(cR.flores ? { k: "barra", t: "Cód. Barra" } : { k: "forn", t: "Fornecedor" });
  if (cR.erp) c.push({ k: "erp", t: "Cód. ERP" });
  if (flv) c.push({ k: "ice", t: "Cód. Iceasa" });
  c.push({ k: "desc", t: "Descrição" });
  if (cR.foto) c.push({ k: "foto", t: "Foto" });
  if (temAlerta) c.push({ k: "alerta", t: "⚠️" });
  for (let l = 1; l <= 8; l++) c.push({ k: "l" + l, t: "L" + p2(l), loja: l });
  if (!cR.texto) c.push({ k: "total", t: "Total" });
  if (flv) c.push({ k: "preco", t: "R$ Preço" });
  if (flv || mp) c.push({ k: "obs", t: "Observação" });
  return c;
}
function desenharGradeSeparacao() {
  const linhas = linhasSeparacao();
  const alertas = alertasSeparacao(cd.ativos.filter((p) => cFiltro === "Todos" || grupoFornecedor(p.f) === cFiltro));
  const porCod = {};
  for (const a of alertas) (porCod[a.p.c] = porCod[a.p.c] || []).push("L" + p2(a.l));
  const cols = colunasSeparacao(alertas.length > 0);
  const al = $("#c-alertas");
  al.hidden = !alertas.length;
  if (alertas.length) al.innerHTML = `<summary><b>${alertas.length} ${alertas.length === 1 ? "pedido" : "pedidos"} acima de ${FATOR_ALERTA_CENTRAL}x a média</b> — conferir antes de fechar</summary>
    <table><thead><tr><th>Cód. ERP</th><th>Descrição</th><th>Loja</th><th>Pedido</th><th>Média</th><th>Vezes</th></tr></thead><tbody>
    ${alertas.map((a) => `<tr><td class="num">${esc(a.p.erp)}</td><td>${esc(a.p.d)}</td><td>${nomeLoja(a.l)}</td><td class="num">${a.q}</td><td class="num">${fmtNum(a.m, 2)}</td><td class="num"><b>${fmtNum(a.x, 1)}</b></td></tr>`).join("")}</tbody></table>`;
  $("#c-grade").innerHTML = `<thead><tr>${cols.map((c) => `<th class="${c.loja || c.k === "total" ? "num-col" : ""}" ${c.loja ? `title="${nomeLoja(c.loja)}"` : ""}>${c.t}</th>`).join("")}</tr></thead>
    <tbody>${linhas.map((p) => `<tr data-c="${p.c}">${cols.map((c) => celulaGrade(p, c, porCod)).join("")}</tr>`).join("")
      || `<tr><td colspan="${cols.length}" class="vazio">Nenhum produto com esse filtro.</td></tr>`}</tbody>`;
  atualizarMetrica();
}
function celulaGrade(p, c, porCod = {}) {
  switch (c.k) {
    case "forn": return `<td class="txt-suave">${esc(p.f)}</td>`;
    case "barra": return `<td class="num txt-suave">${esc(p.b)}</td>`;
    case "erp": return `<td class="num txt-suave">${esc(p.erp)}</td>`;
    case "ice": return `<td class="num txt-suave">${esc(p.ice ?? "")}</td>`;
    case "desc": return `<td class="desc">${esc(p.d)}</td>`;
    case "foto": {
      // miniatura (clique amplia); se a foto não abrir, vira o link "Ver"
      const url = String(p.foto || "").trim();
      if (!/^https?:\/\//.test(url)) return '<td class="foto-cel"></td>';
      return `<td class="foto-cel">${fotoMini(url, "data-ampliar")}</td>`;
    }
    case "alerta": return `<td class="alerta-cel">${esc((porCod[p.c] || []).join(" "))}</td>`;
    case "total": return `<td class="num total" data-total="${p.c}">${totalLinha(p) || ""}</td>`;
    case "preco": return `<td><input class="ext preco" data-p="${p.c}" value="${esc(cExt[p.c]?.p || "")}" placeholder="R$" autocomplete="off"></td>`;
    case "obs": return `<td><input class="ext obs" data-o="${p.c}" value="${esc(cExt[p.c]?.o || "")}" autocomplete="off"></td>`;
  }
  if (c.loja) {
    const l = c.loja;
    if (bloqueado(p.c, l)) return '<td class="bloq" title="Produto não liberado para esta loja">-</td>';
    const v = valor(p.c, l);
    const m = cMedia[chave(p.c, l)] || 0, q = qtdInt(v);
    const alto = ehFLV(cd.setor) && m > 0 && q > FATOR_ALERTA_CENTRAL * m;
    const mudou = normV(v) !== normV(cOrig[chave(p.c, l)]);
    return `<td class="cel-loja"><input class="${cR.texto ? "texto" : "num"} ${alto ? "alto" : ""} ${mudou ? "mudou" : ""}" data-k="${p.c}|${l}" value="${esc(v)}" ${cR.texto ? "" : 'inputmode="numeric"'} autocomplete="off" aria-label="${esc(p.d)} · ${nomeLoja(l)}"></td>`;
  }
  return "<td></td>";
}
// edição: células de loja, preço e observação (vale para Separação e Fornecedores)
function ligarEdicao(raiz) {
  raiz.addEventListener("input", (e) => {
    const i = e.target;
    if (i.dataset.k) {
      if (!cR.texto) { const limpo = i.value.replace(/[^\d.,]/g, ""); if (limpo !== i.value) i.value = limpo; }
      cVal[i.dataset.k] = i.value;
      const [c] = i.dataset.k.split("|");
      i.classList.toggle("mudou", normV(i.value) !== normV(cOrig[i.dataset.k]));
      const p = cd.produtos.find((x) => String(x.c) === c);
      $$(`[data-total="${c}"]`).forEach((t) => (t.textContent = totalLinha(p) || ""));
    } else if (i.dataset.p) {
      cExt[i.dataset.p] = { ...(cExt[i.dataset.p] || { o: "" }), p: i.value };
      i.classList.toggle("mudou", celulaPreco(i.value) !== celulaPreco(cExtOrig[i.dataset.p]?.p));
    } else if (i.dataset.o) {
      cExt[i.dataset.o] = { ...(cExt[i.dataset.o] || { p: "" }), o: i.value };
      i.classList.toggle("mudou", i.value.trim() !== String(cExtOrig[i.dataset.o]?.o || "").trim());
    } else return;
    atualizarMetrica();
  });
  // Enter desce para a mesma coluna da linha de baixo (como numa planilha)
  raiz.addEventListener("keydown", (e) => {
    const i = e.target.closest("input"); if (!i) return;
    if (e.key !== "Enter" && e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    e.preventDefault();
    const td = i.closest("td"), tr = td.parentElement;
    const idx = [...tr.children].indexOf(td);
    let alvo = e.key === "ArrowUp" ? tr.previousElementSibling : tr.nextElementSibling;
    while (alvo && !alvo.children[idx]?.querySelector("input")) alvo = e.key === "ArrowUp" ? alvo.previousElementSibling : alvo.nextElementSibling;
    const prox = alvo?.children[idx]?.querySelector("input");
    if (prox) { prox.focus(); prox.select(); }
  });
}
function atualizarMetrica() {
  const lista = abaAtual === "fornecedores" ? cd.ativos : linhasSeparacao();
  const n = lista.filter((p) => totalLinha(p) > 0).length;
  const m = $("#c-metrica"); if (m) m.innerHTML = `<b class="num">${n}</b> ${n === 1 ? "item" : "itens"} com pedido`;
  const r = $("#c-resumo"); if (!r) return;
  const nc = celulasSujas().length, ne = extrasSujos().length;
  r.innerHTML = nc + ne
    ? `<strong class="num" style="color:var(--laranja)">${nc + ne} ${nc + ne === 1 ? "alteração não salva" : "alterações não salvas"}</strong><small>Clique em Salvar ajustes para gravar.</small>`
    : `<strong>Tudo salvo</strong><small>Atualizado às ${horaBR()}. Use Atualizar para ver o que as lojas digitaram depois.</small>`;
}

// ---------- barra de ações da central ----------
function barraCentral(origem) {
  const b = $("#barra-total");
  b.hidden = false; document.body.classList.add("com-barra");
  b.innerHTML = `<div class="resumo" id="c-resumo"></div>
    <button class="btn btn-secundario" type="button" id="c-atualizar">Atualizar</button>
    <button class="btn btn-secundario" type="button" id="c-excel">Exportar Excel</button>
    <button class="btn btn-secundario" type="button" id="c-imprimir">Imprimir</button>
    <button class="btn btn-primario" type="button" id="c-salvar">Salvar ajustes</button>`;
  $("#c-salvar").onclick = () => salvarAjustes(origem);
  $("#c-atualizar").onclick = async () => {
    if (centralSujo() && !confirm("Há ajustes não salvos. Atualizar e perder esses ajustes?")) return;
    await abrirSetorCentral(cSetor, abaAtual, true);
    toast("Dados atualizados", "ok");
  };
  $("#c-excel").onclick = () => (abaAtual === "fornecedores" ? excelFornecedores() : excelSeparacao());
  $("#c-imprimir").onclick = () => (abaAtual === "fornecedores" ? imprimirFornecedores() : imprimirSeparacao());
  atualizarMetrica();
}
function esconderBarraCentral() { $("#barra-total").hidden = true; document.body.classList.remove("com-barra"); }

async function salvarAjustes(origem) {
  const cels = celulasSujas().map((k) => { const [c, l] = k.split("|").map(Number); return { c, l, v: normV(cVal[k]) }; });
  const exts = extrasSujos().map((c) => ({ c: Number(c), p: celulaPreco(cExt[c].p), o: String(cExt[c].o || "").trim() }));
  if (!cels.length && !exts.length) { toast("Nada para salvar."); return; }
  const btn = $("#c-salvar");
  btn.disabled = true; btn.textContent = "Gravando...";
  try {
    const r = await rpc("pedidos_web_central_salvar", { p_token: sessao.token, p_setor: cd.setor, p_origem: origem, p_celulas: cels, p_extras: exts });
    centralPainel = null;
    await abrirSetorCentral(cSetor, abaAtual, true);
    toast(`Alterações consolidadas: ${r.celulas} ${r.celulas === 1 ? "célula" : "células"}${r.lojas ? ` em ${r.lojas} ${r.lojas === 1 ? "loja" : "lojas"}` : ""}${r.extras ? `, ${r.extras} preço/observação` : ""}.`, "ok");
  } catch (err) {
    toast(err.message, "erro");
    btn.disabled = false; btn.textContent = "Salvar ajustes";
  }
}

// ============================================================
// Visão Fornecedores (Resumo)
// ============================================================
function gruposFornecedores() {
  return ordenarGrupos(cd.ativos.map((p) => p.f)).map((f) => {
    const prods = cd.ativos.filter((p) => p.f === f).sort((a, b) => a.d.localeCompare(b.d, "pt-BR", { sensitivity: "base" }));
    const lojas = [1, 2, 3, 4, 5, 6, 7, 8].filter((l) => !prods.every((p) => bloqueado(p.c, l)));
    return { f, prods, lojas };
  });
}
function colunasFornecedor(lojas) {
  const flv = ehFLV(cd.setor), mp = norm(cd.setor) === "materia prima";
  const c = [];
  if (cR.erp) c.push({ k: "erp", t: "Cód. ERP" });
  if (flv) c.push({ k: "ice", t: "Cód. Iceasa" });
  if (cR.flores) c.push({ k: "barra", t: "Cód. Barra" });
  c.push({ k: "desc", t: "Produto" });
  if (cR.foto) c.push({ k: "foto", t: "Foto" });
  for (const l of lojas) c.push({ k: "l" + l, t: "L" + p2(l), loja: l });
  if (!cR.texto) c.push({ k: "total", t: "Total" });
  if (mp) c.push({ k: "obs", t: "Observação" });
  return c;
}
function renderFornecedores() {
  const grupos = gruposFornecedores();
  $("#conteudo").innerHTML = `<div class="pagina">
    <div class="barra-filtros"><span class="metrica" id="c-metrica"></span></div>
    ${grupos.map((g) => {
      const cols = colunasFornecedor(g.lojas);
      return `<div class="cartao"><h3>Fornecedor: ${esc(g.f)} <small>${g.prods.length} ${g.prods.length === 1 ? "produto" : "produtos"}</small></h3>
        <div class="tabela-wrap grade-wrap"><table class="grade">
          <thead><tr>${cols.map((c) => `<th class="${c.loja || c.k === "total" ? "num-col" : ""}">${c.t}</th>`).join("")}</tr></thead>
          <tbody>${g.prods.map((p) => `<tr data-c="${p.c}">${cols.map((c) => celulaGrade(p, c)).join("")}</tr>`).join("")}</tbody>
        </table></div></div>`;
    }).join("") || '<div class="cartao"><p class="vazio">Nenhum produto liberado neste setor.</p></div>'}
    ${obsLojasHtml()}
  </div>`;
  $$(".grade", $("#conteudo")).forEach(ligarEdicao);
  barraCentral("Visão Fornecedores (Resumo)");
}

// ============================================================
// Pedido por Fornecedor (FLV Normal/Ofertas) — só para impressão do dia
// ============================================================
function iniciarPF() {
  const porIce = {};
  for (const p of cd.produtos) if (p.ice !== null && p.ice !== undefined) porIce[Number(p.ice)] = p.c;
  pf = Object.entries(MAPA_FORNECEDORES_FLV).map(([nome, ices]) => ({
    nome, cods: ices.map((i) => porIce[i]).filter((c) => c !== undefined), linha: FORNECEDORES_LINHA_FLV.has(nome),
  }));
}
function rotuloProduto(p) {
  const dup = cd.produtos.filter((x) => x.d === p.d).length > 1;
  return dup ? `${p.d} [${p.ice ?? p.c}]` : p.d;
}
// quantidades e preço já gravados (pedidos das lojas e Separação)
const qtdBanco = (c, l) => (bloqueado(c, l) ? 0 : qtdInt(cOrig[chave(c, l)]));
const precoBanco = (c) => celulaPreco(cExtOrig[c]?.p) || 0;
function blocoPF(f) {
  const prods = f.cods.map((c) => cd.produtos.find((p) => p.c === c)).filter(Boolean);
  if (!prods.length) return { html: '<p class="dica">Sem produtos selecionados.</p>', print: "" };
  if (f.linha) {
    const cab = prods.map((p) => `${p.ice ?? p.c} - ${p.d.split(/\s+/).slice(0, 2).join(" ")}`);
    const linhas = [1, 2, 3, 4, 5, 6, 7, 8].map((l) => [p2(l), ...prods.map((p) => qtdBanco(p.c, l))]);
    linhas.push(["TOTAL", ...prods.map((p) => [1, 2, 3, 4, 5, 6, 7, 8].reduce((n, l) => n + qtdBanco(p.c, l), 0))]);
    const t = `<table><thead><tr><th>Loja</th>${cab.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>${linhas.map((r) => `<tr>${r.map((v, i) => `<td class="${i ? "num" : ""}">${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
    return { html: `<div class="tabela-wrap">${t.replace("<table>", '<table class="tabela-simples">')}</div>`, print: t };
  }
  let soma = 0;
  const linhas = prods.map((p) => {
    const tot = [1, 2, 3, 4, 5, 6, 7, 8].reduce((n, l) => n + qtdBanco(p.c, l), 0);
    const pre = precoBanco(p.c), rt = tot * pre; soma += rt;
    return [iceasaImpressao(p.ice), p.d, tot > 0 ? tot : "", pre > 0 ? precoCelula(pre) : "", rt > 0 ? precoCelula(rt) : ""];
  });
  const t = `<table><thead><tr><th>Cód</th><th>Produto</th><th>Total</th><th>R$ Preço</th><th>R$ Total</th></tr></thead><tbody>${linhas.map((r) => `<tr>${r.map((v, i) => `<td class="${i === 1 ? "esq" : "num"}">${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  const tf = precoCelula(soma) || "R$ 0,00";
  return {
    html: `<div class="tabela-wrap">${t.replace("<table>", '<table class="tabela-simples">')}</div><p class="total-final">Total final: ${tf}</p>`,
    print: `${t}<div class="total-final">Total Final: ${tf}</div>`,
  };
}
function renderPorFornecedor() {
  if (!pf) iniciarPF();
  esconderBarraCentral();
  const opcoes = cd.produtos.map((p) => ({ c: p.c, r: rotuloProduto(p) })).sort((a, b) => a.r.toLowerCase().localeCompare(b.r.toLowerCase(), "pt-BR"));
  $("#conteudo").innerHTML = `<div class="pagina">
    <p class="faixa info"><span>As quantidades vêm dos pedidos <b>salvos</b> das lojas e os preços da Separação. Os ajustes de nome e de produtos aqui valem só para a impressão de agora.</span></p>
    <div class="barra-filtros">
      <input class="busca" id="pf-novo" placeholder="Nome do novo fornecedor" autocomplete="off">
      <button class="btn btn-secundario" type="button" id="pf-add">Adicionar fornecedor</button>
      <button class="btn btn-primario" type="button" id="pf-imprimir">Imprimir</button>
    </div>
    <datalist id="pf-opcoes">${opcoes.map((o) => `<option value="${esc(o.r)}"></option>`).join("")}</datalist>
    <div class="pf-grade">${pf.map((f, i) => {
      const b = blocoPF(f);
      return `<div class="cartao pf-card" data-i="${i}">
        <input class="pf-nome" data-campo="nome" value="${esc(f.nome)}" aria-label="Nome do fornecedor">
        <div class="pf-prods">${f.cods.map((c) => { const p = cd.produtos.find((x) => x.c === c); return p ? `<span class="pf-chip">${esc(rotuloProduto(p))}<button type="button" data-tirar="${c}" aria-label="Tirar">×</button></span>` : ""; }).join("")}
          <input class="pf-mais" list="pf-opcoes" placeholder="+ produto" data-campo="mais" autocomplete="off"></div>
        <label class="marcar"><input type="checkbox" data-campo="linha" ${f.linha ? "checked" : ""}> Layout girado (lojas nas linhas)</label>
        ${b.html}
      </div>`;
    }).join("")}</div>
  </div>`;
  const porRotulo = Object.fromEntries(opcoes.map((o) => [o.r, o.c]));
  const grade = $(".pf-grade");
  grade.addEventListener("change", (e) => {
    const card = e.target.closest(".pf-card"); if (!card) return;
    const f = pf[Number(card.dataset.i)];
    const campo = e.target.dataset.campo;
    if (campo === "nome") { f.nome = e.target.value.trim() || f.nome; return; }
    if (campo === "linha") f.linha = e.target.checked;
    if (campo === "mais") { const c = porRotulo[e.target.value]; if (c === undefined) return; if (!f.cods.includes(c)) f.cods.push(c); }
    renderPorFornecedor();
  });
  grade.addEventListener("click", (e) => {
    const b = e.target.closest("[data-tirar]"); if (!b) return;
    const f = pf[Number(b.closest(".pf-card").dataset.i)];
    f.cods = f.cods.filter((c) => String(c) !== b.dataset.tirar);
    renderPorFornecedor();
  });
  $("#pf-add").onclick = () => {
    const n = $("#pf-novo").value.trim();
    if (!n || pf.some((f) => f.nome === n)) return;
    pf.push({ nome: n, cods: [], linha: false }); renderPorFornecedor();
  };
  $("#pf-imprimir").onclick = () => {
    const blocos = pf.map((f) => ({ f, b: blocoPF(f) })).filter((x) => x.b.print);
    imprimirHtml(`<h3>🚚 Pedido por Fornecedor — ${esc(cd.setor)}</h3><p class="emitido">Emitido em ${dataHoraBR()}</p>
      ${blocos.map(({ f, b }) => `<div class="bloco"><h4 class="forn">🛒 ${esc(f.nome)}</h4>${b.print}</div>`).join("")}`, "portrait", "pf");
  };
}

// ============================================================
// Log de Pedidos
// ============================================================
function renderLog() {
  const hoje = agoraSP();
  const iso = (d) => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
  const sete = new Date(hoje); sete.setDate(sete.getDate() - 7);
  $("#conteudo").innerHTML = `<div class="pagina"><div class="cartao">
    <p class="dica" style="margin-bottom:14px">Registra quando cada loja confirmou o pedido (Salvar pedido ou Sem pedido hoje), quando o Administrador salvou por uma loja na Separação ou no Resumo por Fornecedor, e quando alguém usou o Limpar pedidos (pelo app novo). Não mostra cada número digitado durante o salvamento automático.</p>
    <div class="barra-filtros">
      <label class="campo-inline">Loja <select id="log-loja"><option value="">Todas</option>${LOJAS.map((l, i) => `<option value="${i + 1}">${l}</option>`).join("")}</select></label>
      <label class="campo-inline">De <input type="date" id="log-de" value="${iso(sete)}"></label>
      <label class="campo-inline">Até <input type="date" id="log-ate" value="${iso(hoje)}"></label>
    </div>
    <div id="log-tabela"><p class="vazio">Carregando...</p></div></div></div>`;
  ["#log-loja", "#log-de", "#log-ate"].forEach((s) => $(s).addEventListener("change", carregarLog));
  carregarLog();
}
const ROTULO_ACAO = { salvar_pedido: "Pedido salvo (loja)", sem_pedido_hoje: "Sem pedido hoje", salvar_admin: "Salvo pelo Administrador", limpar_pedidos: "Pedidos limpos" };
async function carregarLog() {
  const de = $("#log-de").value, ate = $("#log-ate").value, el = $("#log-tabela");
  if (de && ate && de > ate) { el.innerHTML = '<p class="faixa aviso">A data inicial é depois da data final. Ajuste o período.</p>'; return; }
  try {
    const linhas = await rpc("pedidos_web_log", { p_token: sessao.token, p_setor: cSetor.nome, p_loja: $("#log-loja").value ? Number($("#log-loja").value) : null, p_de: de, p_ate: ate });
    el.innerHTML = linhas.length ? `<div class="tabela-wrap"><table class="tabela-simples"><thead><tr><th>Quando</th><th>Loja</th><th>Ação</th><th>Usuário</th><th>Origem</th><th>Itens</th></tr></thead><tbody>
      ${linhas.map((r) => `<tr><td class="num">${esc(r.quando)}</td><td>${r.loja === 0 ? "Todas as lojas" : nomeLoja(r.loja)}</td><td><span class="tag ${r.acao === "salvar_admin" ? "novo" : r.acao === "sem_pedido_hoje" ? "sem" : r.acao === "limpar_pedidos" ? "alerta" : "ok"}">${esc(ROTULO_ACAO[r.acao] || r.acao)}</span></td><td>${esc(r.usuario)}</td><td>${esc(r.origem)}</td><td class="num">${r.acao === "limpar_pedidos" ? `${r.itens} apagados` : r.itens}</td></tr>`).join("")}</tbody></table></div>
      <p class="dica" style="margin-top:10px">${linhas.length} ${linhas.length === 1 ? "registro" : "registros"} no período.</p>`
      : '<p class="vazio">Nenhum registro encontrado para o período e a loja escolhidos.</p>';
  } catch (err) { el.innerHTML = `<p class="faixa erro">${esc(err.message)}</p>`; }
}

// ============================================================
// Histórico (ciclos fechados pelo Limpar pedidos)
// ============================================================
const FAIXA_MEDIA = { alto: 1.5, baixo: 0.7 };    // pedido ÷ média: acima de 1,5x ou abaixo de 0,7x chama atenção
const fmtX = (r) => r.toLocaleString("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "x";
let histDados = null, histLimite = 40;
async function renderHistorico(s) {
  $("#conteudo").innerHTML = '<div class="vazio">Carregando histórico...</div>';
  try {
    const [h, d] = await Promise.all([
      rpc("pedidos_web_historico_setor", { p_token: sessao.token, p_setor: s.nome, p_ciclos: 10 }),
      // busca o pedido atual de novo (as lojas podem ter digitado), a não ser que haja ajuste não salvo na Separação
      cd && cd.setor === s.nome && centralSujo() ? null : rpc("pedidos_web_central_setor", { p_token: sessao.token, p_setor: s.nome }),
    ]);
    if (d) prepararDados(d);
    histDados = h; histLimite = 40;
  } catch (err) {
    const falta = /Could not find|historico_setor/.test(err.message) || err.message === MENSAGENS["Could not find the function"];
    $("#conteudo").innerHTML = `<div class="cartao"><p class="faixa ${falta ? "aviso" : "erro"}">${falta ? "O histórico ainda não está no banco (falta rodar o SQL da versão 1.5)." : esc(err.message)}</p></div>`;
    return;
  }
  if (abaAtual !== "historico" || cSetor?.nome !== s.nome) return;
  desenharHistorico();
}
// totais do pedido aberto (o que está na Separação agora), por loja
function atualPorLoja() {
  const t = {};
  for (let l = 1; l <= 8; l++) {
    let qtd = 0, itens = 0;
    for (const p of cd.ativos) {
      if (bloqueado(p.c, l)) continue;
      const v = valor(p.c, l);
      if (cR.texto ? String(v).trim() : qtdInt(v) > 0) { itens++; qtd += cR.texto ? 0 : qtdInt(v); }
    }
    t[l] = { itens, qtd };
  }
  return t;
}
function desenharHistorico() {
  const h = histDados, ciclos = h.ciclos || [];
  if (!ciclos.length) {
    $("#conteudo").innerHTML = `<div class="pagina"><div class="cartao"><h3>Histórico</h3>
      <p class="faixa info"><span>Ainda não há ciclo guardado neste setor. A partir de agora, cada vez que alguém usar <b>Limpar pedidos</b>, o pedido final de cada loja (já com os ajustes da Separação) fica guardado aqui, com a média de venda daquele dia.</span></p></div></div>`;
    return;
  }
  const texto = cR.texto, ult = ciclos[0];
  const medida = (o) => (texto ? o.itens : o.qtd);   // setores em texto comparam nº de itens
  const unid = texto ? "itens" : "unidades";
  // ---- 1. pedido atual × último ciclo, por loja ----
  const atual = atualPorLoja();
  const linhaLojas = LOJAS.map((nome, i) => {
    const l = i + 1, a = medida(atual[l]), u = ult.lojas?.[l] ? Number(medida(ult.lojas[l])) : 0;
    const d = u ? (a - u) / u : null;
    return `<tr><td>${nome}</td><td class="num">${fmtNum(a)}</td><td class="num">${fmtNum(u)}</td>
      <td class="num">${!a && !u ? "" : d === null ? '<span class="tag novo">novo</span>' : `<span class="${d > 0.3 ? "delta-mais" : d < -0.3 ? "delta-menos" : ""}">${d > 0 ? "+" : ""}${Math.round(d * 100)}%</span>`}</td></tr>`;
  }).join("");
  // ---- 2. por produto ----
  const ultProd = {};
  for (const [c, l, q, t] of h.ultimo || []) {
    const o = (ultProd[c] ||= { q: 0, n: 0 });
    o.q += Number(q) || 0; if (texto ? String(t).trim() : Number(q) > 0) o.n++;
  }
  const prods = cd.produtos.map((p) => {
    let qa = 0, na = 0;
    for (let l = 1; l <= 8; l++) { if (bloqueado(p.c, l)) continue; const v = valor(p.c, l); if (texto ? String(v).trim() : qtdInt(v) > 0) { na++; qa += texto ? 0 : qtdInt(v); } }
    const u = ultProd[p.c] || { q: 0, n: 0 };
    const a = texto ? na : qa, b = texto ? u.n : u.q;
    return { p, a, b, d: a - b };
  }).filter((x) => x.a || x.b);
  // ---- 3. pedido × média, por loja, nos ciclos guardados ----
  const razao = {};
  for (const c of ciclos) for (const [l, o] of Object.entries(c.lojas || {})) {
    if (!(Number(o.m) > 0)) continue;
    const r = (razao[l] ||= { qm: 0, m: 0, n: 0 }); r.qm += Number(o.qm) || 0; r.m += Number(o.m); r.n++;
  }
  const temMedia = !texto && Object.keys(razao).length > 0;
  const celCiclo = (o) => {
    if (!o) return '<td class="num vazio-h">–</td>';
    const r = !texto && Number(o.m) > 0 ? Number(o.qm) / Number(o.m) : null;
    const cls = r === null ? "" : r > FAIXA_MEDIA.alto ? "alto" : r < FAIXA_MEDIA.baixo ? "baixo" : "";
    return `<td class="num ${cls}">${fmtNum(medida(o))}${r !== null ? `<small title="pedido ÷ média de venda">${fmtX(r)}</small>` : ""}</td>`;
  };
  $("#conteudo").innerHTML = `<div class="pagina">
    <div class="grade-2">
      <div class="cartao"><h3>Pedido atual × último ciclo <small>fechado em ${esc(ult.em)}</small></h3>
        <div class="tabela-wrap"><table class="tabela-simples"><thead><tr><th>Loja</th><th class="num">Atual</th><th class="num">Último ciclo</th><th class="num">Diferença</th></tr></thead>
        <tbody>${linhaLojas}</tbody></table></div>
        <p class="dica" style="margin-top:8px">Em ${unid}. “Atual” é o que está na Separação agora (pode ainda estar sendo digitado).</p></div>
      ${temMedia ? `<div class="cartao"><h3>Pedido × média de venda <small>últimos ${ciclos.length} ciclos</small></h3>
        <div class="tabela-wrap"><table class="tabela-simples"><thead><tr><th>Loja</th><th class="num">Ciclos</th><th class="num">Pedido ÷ média</th><th>Leitura</th></tr></thead><tbody>
        ${LOJAS.map((nome, i) => { const r = razao[i + 1]; if (!r) return `<tr><td>${nome}</td><td class="num">0</td><td></td><td class="dica">sem pedido com média</td></tr>`;
          const x = r.qm / r.m;
          return `<tr><td>${nome}</td><td class="num">${r.n}</td><td class="num"><b>${fmtX(x)}</b></td><td>${x > FAIXA_MEDIA.alto ? '<span class="tag alerta">pede bem acima da média</span>' : x < FAIXA_MEDIA.baixo ? '<span class="tag novo">pede abaixo da média</span>' : '<span class="tag ok">equilibrado</span>'}</td></tr>`; }).join("")}
        </tbody></table></div>
        <p class="dica" style="margin-top:8px">Soma do pedido ÷ soma da média de venda dos mesmos produtos (só produtos com média). Acima de ${fmtX(FAIXA_MEDIA.alto)}: confira sobra e quebra. Abaixo de ${fmtX(FAIXA_MEDIA.baixo)}: risco de faltar na loja.</p></div>` : ""}
    </div>
    <div class="cartao"><h3>Ciclos anteriores <small>${ciclos.length} ${ciclos.length === 1 ? "ciclo" : "ciclos"}</small></h3>
      <div class="tabela-wrap"><table class="tabela-simples hist-grade"><thead><tr><th>Fechado em</th>${LOJAS.map((l, i) => `<th class="num">L${p2(i + 1)}</th>`).join("")}<th class="num">Total</th></tr></thead><tbody>
      ${ciclos.map((c) => `<tr><td class="num">${esc(c.em)}</td>${LOJAS.map((_, i) => celCiclo(c.lojas?.[i + 1])).join("")}<td class="num"><b>${fmtNum(Object.values(c.lojas || {}).reduce((n, o) => n + Number(medida(o)), 0))}</b></td></tr>`).join("")}
      </tbody></table></div>
      <p class="dica" style="margin-top:8px">${texto ? "Número de itens pedidos por loja em cada ciclo." : "Unidades pedidas por loja em cada ciclo; embaixo, o pedido dividido pela média de venda (vermelho: acima de " + fmtX(FAIXA_MEDIA.alto) + "; azul: abaixo de " + fmtX(FAIXA_MEDIA.baixo) + ")."}</p></div>
    <div class="cartao"><h3>Por produto: atual × último ciclo <small>${prods.length} produtos</small></h3>
      <div class="barra-filtros"><input class="busca" id="h-busca" type="search" placeholder="Buscar produto ou código" autocomplete="off"></div>
      <div class="tabela-wrap"><table class="tabela-simples"><thead><tr><th>Produto</th><th class="num">${texto ? "Lojas agora" : "Atual"}</th><th class="num">${texto ? "Lojas no último" : "Último ciclo"}</th><th class="num">Diferença</th></tr></thead><tbody id="h-prods"></tbody></table></div>
      <div id="h-mais"></div></div>
  </div>`;
  const pintar = () => {
    const q = norm($("#h-busca").value);
    const vis = prods.filter((x) => !q || norm(`${x.p.d} ${x.p.erp}`).includes(q)).sort((a, b) => Math.abs(b.d) - Math.abs(a.d) || b.a - a.a);
    $("#h-prods").innerHTML = vis.slice(0, histLimite).map((x) => `<tr><td>${esc(x.p.d)}<small class="dica" style="display:block">${esc(x.p.f)}${cR.erp ? " · " + esc(x.p.erp) : ""}</small></td>
      <td class="num">${fmtNum(x.a)}</td><td class="num">${fmtNum(x.b)}</td>
      <td class="num">${x.d === 0 ? "=" : `<span class="${x.d > 0 ? "delta-mais" : "delta-menos"}">${x.d > 0 ? "+" : ""}${fmtNum(x.d)}</span>`}</td></tr>`).join("")
      || '<tr><td colspan="4" class="vazio">Nada com esse filtro.</td></tr>';
    $("#h-mais").innerHTML = vis.length > histLimite ? `<button class="btn btn-secundario btn-pequeno" type="button" id="h-mais-btn" style="margin-top:10px">Mostrar mais (${vis.length - histLimite})</button>` : "";
    $("#h-mais-btn")?.addEventListener("click", () => { histLimite += 40; pintar(); });
  };
  $("#h-busca").addEventListener("input", () => { histLimite = 40; pintar(); });
  pintar();
}

// ============================================================
// Limpar Pedidos (fecha o ciclo do setor)
// ============================================================
function confirmarLimpar(s) {
  const mp = norm(s.nome) === "materia prima";
  abrirJanela(`<h3>Limpar todos os pedidos</h3>
    <p>Tem certeza que deseja <b>limpar todos os pedidos</b> de <b>${esc(s.titulo)}</b>?</p>
    <p>${mp ? "Apaga TODOS os pedidos pendentes do setor (de qualquer data) e as observações das lojas. As observações por produto são mantidas." : "Apaga TODOS os pedidos pendentes do setor (de qualquer data), preços/observações e observações das lojas — de todas as lojas."} <b>Não pode ser desfeito.</b></p>
    <p class="dica">Antes de apagar, o pedido final de cada loja (com os ajustes da Separação) fica guardado na aba <b>Histórico</b>.</p>
    <div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Não, cancelar</button><button class="btn btn-perigo-forte" type="button" id="btn-limpar-sim">Sim, limpar tudo</button></div>`);
  $("#btn-limpar-sim").onclick = async (e) => {
    const b = e.currentTarget; b.disabled = true; b.textContent = "Limpando...";
    try {
      const r = await rpc("pedidos_web_limpar", { p_token: sessao.token, p_setor: s.nome, p_confirmacao: "LIMPAR" });
      fecharJanela();
      cd = null; centralPainel = null;
      await rotaCentral(s, abaAtual === "log" ? "log" : abaAtual);
      toast(`Pedidos de ${s.titulo} limpos (${r.apagados} lançamentos)${r.historico ? ` · ${r.historico} guardados no Histórico` : ""}.`, "ok");
    } catch (err) { b.disabled = false; b.textContent = "Sim, limpar tudo"; toast(err.message, "erro"); }
  };
}

// ============================================================
// Impressão (mesmos modelos do Streamlit)
// ============================================================
function imprimirHtml(html, orientacao, classe = "") {
  const el = $("#impressao");
  el.className = "so-impressao " + classe;
  el.innerHTML = html;
  let st = $("#estilo-paginas");
  if (!st) { st = document.createElement("style"); st.id = "estilo-paginas"; document.head.appendChild(st); }
  const num = cR?.obs ? ' @bottom-right { content: "Página " counter(page) " de " counter(pages); font-size: 7.5pt; color: #333; }' : "";
  st.textContent = `@media print { @page { size: A4 ${orientacao}; margin: 10mm 10mm ${cR?.obs ? "14mm" : "10mm"} 10mm;${num} } }`;
  window.print();
}
function obsLojasPrint() {
  if (!cR.obs || !(cd.obs_lojas || []).length) return "";
  return `<h3 style="margin-top:14px">📝 Observações das Lojas</h3><table><thead><tr><th style="width:12%">Loja</th><th>Observação</th></tr></thead><tbody>
    ${cd.obs_lojas.map((o) => `<tr><td>${nomeLoja(o.l)}</td><td class="esq pre">${esc(o.o)}</td></tr>`).join("")}</tbody></table>`;
}
const semTraco = (v) => (String(v).trim() === "-" ? "" : v);
function imprimirSeparacao() {
  const flv = ehFLV(cd.setor);
  const linhas = linhasSeparacao();
  const cab = [cR.flores ? "Cód. Barra" : "Fornecedor"];
  if (flv) cab.push("Cód. Iceasa"); else if (cR.erp) cab.push("Cód. ERP");
  cab.push("Descrição", ...LOJAS);
  if (!cR.texto) cab.push("TOTAL GERAL");
  const corpo = linhas.map((p) => {
    const r = [cR.flores ? p.b : p.f];
    if (flv) r.push(iceasaImpressao(p.ice)); else if (cR.erp) r.push(p.erp);
    r.push(p.d);
    for (let l = 1; l <= 8; l++) r.push(bloqueado(p.c, l) ? "" : semTraco(normV(valor(p.c, l))));
    if (!cR.texto) r.push(totalLinha(p) || "");
    return r;
  });
  imprimirHtml(`<h3>📊 Separação e Fechamento — ${esc(cd.setor)} (${esc(cFiltro)})</h3><p class="emitido">Emitido em ${dataHoraBR()}</p>
    ${tabelaPrint(cab, corpo, [0, cab.indexOf("Descrição")])}${obsLojasPrint()}`, "landscape", "sep");
}
function imprimirFornecedores() {
  const flv = ehFLV(cd.setor), mp = norm(cd.setor) === "materia prima";
  const partes = gruposFornecedores().map((g) => {
    const cab = [];
    if (flv) cab.push("Cód. Iceasa"); else if (cR.erp) cab.push("Cód. ERP");
    if (cR.flores) cab.push("Cód. Barra");
    cab.push("Produto", ...g.lojas.map((l) => nomeLoja(l)));
    if (!cR.texto) cab.push("TOTAL GERAL");
    if (mp) cab.push("Observação");
    const corpo = g.prods.map((p) => {
      const r = [];
      if (flv) r.push(iceasaImpressao(p.ice)); else if (cR.erp) r.push(p.erp);
      if (cR.flores) r.push(p.b);
      r.push(p.d);
      for (const l of g.lojas) r.push(bloqueado(p.c, l) ? "-" : normV(valor(p.c, l)));
      if (!cR.texto) r.push(totalLinha(p) || "");
      if (mp) r.push(cExt[p.c]?.o || "");
      return r;
    });
    return `<div class="bloco-forn"><h4 class="forn">🚚 Fornecedor: ${esc(g.f)}</h4>${tabelaPrint(cab, corpo, [cab.indexOf("Produto")])}</div>`;
  }).join("");
  imprimirHtml(`<h3>🚚 Resumo por Fornecedor — ${esc(cd.setor)}</h3><p class="emitido">Emitido em ${dataHoraBR()}</p>${partes}${obsLojasPrint()}`, "landscape", "forn");
}
function tabelaPrint(cab, linhas, colsEsq = []) {
  return `<table><thead><tr>${cab.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>
    ${linhas.map((r) => `<tr>${r.map((v, i) => `<td class="${colsEsq.includes(i) ? "esq" : ""}">${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

// ============================================================
// Excel (mesmos 3 modelos do Streamlit: padrão, PEDIDO BOX e Fornecedores)
// ============================================================
const BORDA = { style: "thin", color: { argb: "FF000000" } };
const BORDAS = { top: BORDA, left: BORDA, bottom: BORDA, right: BORDA };
const BORDA_CAB = { style: "dashed", color: { argb: "FFFFFFFF" } };
const BORDAS_CAB = { top: BORDA_CAB, left: BORDA_CAB, bottom: BORDA_CAB, right: BORDA_CAB };
const FILL = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const LARANJA = "FFC55A11", SALMAO = "FFFCE4D6";
const letra = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const hojeBR = () => { const d = agoraSP(); return `${p2(d.getDate())}/${p2(d.getMonth() + 1)}/${d.getFullYear()}`; };
function margens(ws) {
  ws.pageSetup = { ...ws.pageSetup, paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: 1 / 2.54, right: 1 / 2.54, top: 1 / 2.54, bottom: 1 / 2.54, header: .5 / 2.54, footer: .5 / 2.54 } };
}
async function baixar(wb, nome) {
  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = nome; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
// limpar do gerar_excel_fornecedores: vazio, "-" e 0 ficam em branco; número vira número
function limparForn(v) {
  const s = String(v ?? "").trim();
  if (["", "-", "0", "0.0", "nan"].includes(s)) return null;
  const f = Number(s);
  if (Number.isFinite(f)) return f === 0 ? null : f;
  return s;
}

// modelo padrão (gerar_excel_download): cabeçalho laranja, lojas alternando, Total com fórmula
function planilhaPadrao(wb, nomeAba, cab, linhas) {
  const ws = wb.addWorksheet(nomeAba.slice(0, 30));
  margens(ws);
  ws.getCell("A1").value = `Pedidos do dia ${hojeBR()}`; ws.getCell("A1").font = { bold: true };
  const U = cab.map((c) => c.toUpperCase());
  const lojasIdx = [], verdes = new Set(); let totalIdx = null, nLoja = 0;
  U.forEach((c, i) => {
    if (c.includes("LOJA")) { lojasIdx.push(i + 1); nLoja++; if (nLoja % 2) verdes.add(i + 1); }
    else if (["PEDIDO", "TOTAL", "MÉDIA", "ESTOQUE"].some((x) => c.includes(x))) verdes.add(i + 1);
    if (c.includes("TOTAL")) totalIdx = i + 1;
  });
  cab.forEach((t, i) => {
    const cel = ws.getRow(2).getCell(i + 1);
    cel.value = t; cel.fill = FILL(LARANJA); cel.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cel.border = BORDAS_CAB; cel.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    const c = U[i];
    ws.getColumn(i + 1).width = c.includes("FORNECEDOR") ? 30 : (c.includes("DESCRI") || c.includes("PRODUTO")) ? 45 : c.includes("BARRA") ? 16
      : c.includes("CÓDIGO") ? 12 : (c.includes("MÉDIA") || c.includes("ESTOQUE")) ? 12 : c.includes("OBSERV") ? 22 : 9;
  });
  linhas.forEach((r, ri) => {
    const row = ws.getRow(3 + ri);
    r.forEach((v, i) => {
      const c = U[i];
      if (["LOJA", "PEDIDO", "ESTOQUE", "TOTAL", "MÉDIA"].some((x) => c.includes(x))) v = limparExcel(v);
      const cel = row.getCell(i + 1);
      cel.value = v === "" || v === undefined ? null : v;
      cel.border = BORDAS; cel.font = { bold: true };
      cel.alignment = { horizontal: ["FORNECEDOR", "DESCRIÇÃO", "PRODUTO"].some((x) => c.includes(x)) ? "left" : "center", vertical: "middle", wrapText: true };
      if (verdes.has(i + 1)) cel.fill = FILL(SALMAO);
    });
    if (totalIdx && lojasIdx.length) {
      const n = 3 + ri, a = letra(Math.min(...lojasIdx)), b = letra(Math.max(...lojasIdx));
      row.getCell(totalIdx).value = { formula: `IF(SUM(${a}${n}:${b}${n})>0, SUM(${a}${n}:${b}${n}), "")` };
    }
  });
  return ws;
}
function abaObsLojas(wb) {
  if (!cR.obs || !(cd.obs_lojas || []).length) return;
  const ws = wb.addWorksheet("Observações das Lojas");
  ["Loja", "Observação"].forEach((t, i) => {
    const c = ws.getRow(1).getCell(i + 1);
    c.value = t; c.fill = FILL(LARANJA); c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.border = BORDAS; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  cd.obs_lojas.forEach((o, i) => {
    const r = ws.getRow(2 + i);
    r.getCell(1).value = nomeLoja(o.l); r.getCell(1).alignment = { horizontal: "center", vertical: "middle" }; r.getCell(1).border = BORDAS;
    r.getCell(2).value = o.o; r.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true }; r.getCell(2).border = BORDAS;
  });
  ws.getColumn(1).width = 12; ws.getColumn(2).width = 110;
  ws.pageSetup = { ...ws.pageSetup, paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}
// PEDIDO BOX (FLV Normal/Ofertas): A Iceasa · B produto · C-J lojas 291..298 · K-P ocultas · Q total · R preço · S obs
function abaBox(wb, nome, prods) {
  const ws = wb.addWorksheet(nome.replace(/[:\\/?*[\]]/g, " ").slice(0, 31) || "PEDIDO");
  const fonte = { name: "Arial", size: 9, bold: true };
  const cab = ["CODIGO", "PRODUTOS MOLICENTER", "291", "292", "293", "294", "295", "296", "297", "298", "", "", "", "", "", "", "TOTAL", "PREÇO", "OBS:"];
  cab.forEach((t, i) => {
    const c = ws.getRow(2).getCell(i + 1);
    c.value = t; c.fill = FILL(LARANJA); c.font = { ...fonte, color: { argb: "FFFFFFFF" } }; c.border = BORDAS;
    c.alignment = { horizontal: i === 1 ? "left" : "center", vertical: "middle" };
  });
  const fundo = {}; for (let i = 0; i < 8; i++) fundo[3 + i] = i % 2 === 0 ? "FFE2EFDA" : "FFFFFFFF";
  fundo[17] = "FFC6EFCE"; fundo[18] = "FFFCE4D6";
  const qtdBox = (v) => { const s = String(v ?? "").trim(); if (["", "-", "0", "0.0"].includes(s)) return null; const f = Number(s.replace(",", ".")); return Number.isFinite(f) && f !== 0 ? f : null; };
  const lista = [...prods].sort((a, b) => a.d.toLowerCase().localeCompare(b.d.toLowerCase()));
  lista.forEach((p, i) => {
    const n = 3 + i, row = ws.getRow(n);
    const ice = Math.trunc(Number(p.ice));
    row.getCell(1).value = p.ice !== null && p.ice !== undefined && Number.isFinite(ice) && ice <= 9000 ? ice : null;
    row.getCell(2).value = String(p.d).trim();
    for (let l = 1; l <= 8; l++) row.getCell(2 + l).value = bloqueado(p.c, l) ? null : qtdBox(valor(p.c, l));
    row.getCell(17).value = { formula: `IF(SUM(C${n}:J${n})=0,"",SUM(C${n}:J${n}))` };
    for (let col = 1; col <= 19; col++) {
      const c = row.getCell(col);
      c.font = fonte; c.border = BORDAS; c.alignment = { horizontal: col === 2 ? "left" : "center", vertical: "middle" };
      c.fill = FILL(fundo[col] || "FFFFFFFF");
      if (col === 18) c.numFmt = '[$R$-416]\\ #,##0.00';
    }
    row.height = 15;
  });
  const larg = { A: 9, B: 34, C: 8.5, D: 8.5, E: 8.5, F: 8.5, G: 8.5, H: 8.5, I: 8.5, J: 8.5, K: 3, L: 13, M: 13, N: 13, O: 13, P: 13, Q: 8.5, R: 12, S: 22 };
  Object.entries(larg).forEach(([L, w]) => (ws.getColumn(L).width = w));
  ["K", "L", "M", "N", "O", "P"].forEach((L) => (ws.getColumn(L).hidden = true));
  ws.getRow(1).height = 6; ws.getRow(2).height = 18;
  ws.autoFilter = "A2:S2";
  ws.views = [{ state: "frozen", ySplit: 2 }];
  ws.pageSetup = { ...ws.pageSetup, orientation: "portrait", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: .75, right: .75, top: 1, bottom: 1, header: .3, footer: .3 } };
}
async function comExcel(btnSel, fn) {
  const btn = $(btnSel); btn.disabled = true; const t = btn.textContent; btn.textContent = "Gerando...";
  try { await carregarExcelJS(); await fn(); } catch (err) { toast(err.message, "erro"); }
  finally { btn.disabled = false; btn.textContent = t; }
}
function excelSeparacao() {
  return comExcel("#c-excel", async () => {
    const wb = new ExcelJS.Workbook();
    const linhas = linhasSeparacao();
    if (ehFLV(cd.setor)) {
      if (cFiltro === "Todos") {
        // uma aba por grupo de fornecedor (Box, Pedra...)
        const grupos = ordenarGrupos(linhas.map((p) => grupoFornecedor(p.f)));
        for (const g of grupos) abaBox(wb, `PEDIDO ${g.toUpperCase()}`, linhas.filter((p) => grupoFornecedor(p.f) === g));
        if (!grupos.length) abaBox(wb, "PEDIDO BOX", linhas);
      } else abaBox(wb, `PEDIDO ${cFiltro.toUpperCase()}`, linhas);
      return baixar(wb, "molicenter.xlsx");
    }
    const mp = norm(cd.setor) === "materia prima";
    const cab = [cR.flores ? "Cód. Barra" : "Fornecedor"];
    if (cR.erp) cab.push("Código");
    cab.push("Descrição", ...LOJAS);
    if (!cR.texto) cab.push("Total");
    if (mp) cab.push("Observação:");
    const corpo = linhas.map((p) => {
      const r = [cR.flores ? p.b : p.f];
      if (cR.erp) r.push(Number(p.erp));
      r.push(p.d);
      for (let l = 1; l <= 8; l++) r.push(bloqueado(p.c, l) ? "-" : normV(valor(p.c, l)));
      if (!cR.texto) r.push(totalLinha(p) || "");
      if (mp) r.push(cExt[p.c]?.o || "");
      return r;
    });
    planilhaPadrao(wb, `Fechamento ${cd.setor}`, cab, corpo);
    return baixar(wb, `Separacao_Fechamento_${cd.setor}.xlsx`);
  });
}
function excelFornecedores() {
  return comExcel("#c-excel", async () => {
    const wb = new ExcelJS.Workbook();
    const grupos = gruposFornecedores();
    const mp = norm(cd.setor) === "materia prima";
    if (cR.flores) {
      // Flores: modelo padrão (o layout por fornecedor não comporta o Cód. Barra)
      const cab = ["Código", "Cód. Barra", "Descrição", ...LOJAS, "Total"];
      const corpo = grupos.flatMap((g) => g.prods.map((p) => [Number(p.erp), p.b, p.d,
        ...[1, 2, 3, 4, 5, 6, 7, 8].map((l) => (bloqueado(p.c, l) ? "-" : normV(valor(p.c, l)))), totalLinha(p) || ""]));
      planilhaPadrao(wb, "Fornecedores", cab, corpo);
      return baixar(wb, `Resumo_Fornecedores_${cd.setor}.xlsx`);
    }
    const ws = wb.addWorksheet("Fornecedores");
    const totalCol = cR.texto ? null : 11;
    const obsCol = mp ? (totalCol ? totalCol + 1 : 11) : null;
    const ultima = Math.max(10, totalCol || 0, obsCol || 0);
    ws.getColumn(1).width = 12; ws.getColumn(2).width = 45;
    for (let c = 3; c <= ultima; c++) ws.getColumn(c).width = 9;
    if (obsCol) ws.getColumn(obsCol).width = 22;
    const pintaCab = (row) => { for (let c = 1; c <= ultima; c++) { const x = ws.getRow(row).getCell(c); x.fill = FILL(LARANJA); x.font = { bold: true, color: { argb: "FFFFFFFF" } }; x.border = BORDAS_CAB; x.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; } };
    ws.getCell("A1").value = "Tipo"; ws.getCell("B1").value = "Descrição"; ws.getCell("C1").value = `Pedidos do dia ${hojeBR()}`;
    ws.mergeCells(1, 3, 1, ultima); pintaCab(1);
    let r = 2;
    for (const g of grupos) {
      ws.getRow(r).getCell(2).value = g.f;
      LOJAS.forEach((l, i) => (ws.getRow(r).getCell(3 + i).value = l));
      if (totalCol) ws.getRow(r).getCell(totalCol).value = "TOTAL";
      if (obsCol) ws.getRow(r).getCell(obsCol).value = "Observação";
      pintaCab(r); r++;
      for (const p of g.prods) {
        const row = ws.getRow(r);
        const c1 = row.getCell(1); c1.value = cR.erp ? limparForn(p.erp) : null; c1.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c1.border = BORDAS;
        const c2 = row.getCell(2); c2.value = p.d; c2.alignment = { horizontal: "left", vertical: "middle", wrapText: true }; c2.border = BORDAS;
        for (let l = 1; l <= 8; l++) {
          const c = row.getCell(2 + l);
          c.value = bloqueado(p.c, l) ? null : limparForn(normV(valor(p.c, l)));
          c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = BORDAS; c.font = { bold: true };
          if ((l - 1) % 2 === 0) c.fill = FILL(SALMAO);
        }
        if (totalCol) { const c = row.getCell(totalCol); c.value = { formula: `IF(SUM(C${r}:J${r})>0, SUM(C${r}:J${r}), "")` }; c.alignment = { horizontal: "center", vertical: "middle" }; c.border = BORDAS; c.font = { bold: true }; c.fill = FILL(SALMAO); }
        if (obsCol) { const c = row.getCell(obsCol); c.value = String(cExt[p.c]?.o || "").trim() || null; c.alignment = { horizontal: "left", vertical: "middle", wrapText: true }; c.border = BORDAS; }
        r++;
      }
    }
    margens(ws);
    abaObsLojas(wb);
    return baixar(wb, `Resumo_Fornecedores_${cd.setor}.xlsx`);
  });
}


// ============================================================
// Catálogo de Produtos (etapa 3)
// ============================================================
let cat = null;          // pedidos_web_catalogo
let catLinhas = [];      // [{ id, c (null = novo), v: {campos}, o: {campos originais}, l: [8], lo: [8], x: excluir }]
let catFiltro = "Todos";
let catSeq = 0;
const CAMPOS_CAT = ["fornecedor", "codigo_erp", "codigo_iceasa", "codigo_barra", "link_foto", "descricao", "nome_personalizado", "nome_ceasa"];
function catalogoSujo() { return !!cat && catLinhas.some((r) => r.x || (!r.c && !novoVazio(r)) || camposMudados(r).length || lojasMudaram(r)); }
const strV = (v) => (v === null || v === undefined ? "" : String(v).trim());
function camposMudados(r) { return r.c ? CAMPOS_CAT.filter((k) => strV(r.v[k]) !== strV(r.o[k])) : []; }
const lojasMudaram = (r) => r.c && r.l.some((v, i) => v !== r.lo[i]);
function novoVazio(r) { return !r.c && !strV(r.v.descricao) && !strV(r.v.codigo_erp); }

async function abrirCatalogo(s, recarregar) {
  telaAtual = "central"; cSetor = s; cR = regras(s.nome);
  if (!cat || cat.setor !== s.nome || recarregar) {
    $("#conteudo").innerHTML = '<div class="vazio">Carregando catálogo...</div>';
    esconderBarraCentral();
    try {
      cat = await rpc("pedidos_web_catalogo", { p_token: sessao.token, p_setor: s.nome });
    } catch (err) {
      $("#conteudo").innerHTML = `<div class="cartao"><p class="faixa erro">${esc(err.message)}</p></div>`;
      return;
    }
    catFiltro = "Todos";
    catLinhas = cat.produtos.map((p) => {
      const v = { fornecedor: p.f, codigo_erp: p.erp ?? "", codigo_iceasa: p.ice ?? "", codigo_barra: p.b, link_foto: p.foto, descricao: p.d, nome_personalizado: p.nm, nome_ceasa: p.nc };
      return { id: ++catSeq, c: p.c, v, o: { ...v }, l: [...p.l], lo: [...p.l], x: false };
    });
  }
  renderCatalogo();
}
function colunasCatalogo() {
  const c = [{ k: "fornecedor", t: "Fornecedor", w: "fornecedor" }];
  if (cR.erp) c.push({ k: "codigo_erp", t: "Cód. ERP", w: "cod" });
  if (ehFLV(cat.setor)) c.push({ k: "codigo_iceasa", t: "Cód. Iceasa", w: "cod" });
  if (cR.flores) c.push({ k: "codigo_barra", t: "Cód. Barra", w: "barra" });
  if (cR.foto) c.push({ k: "link_foto", t: "Link Foto", w: "foto" });
  c.push({ k: "descricao", t: "Nome Prime", w: "nome" }, { k: "nome_personalizado", t: "Nome Manual", w: "nome" });
  if (cat.nome_ceasa) c.push({ k: "nome_ceasa", t: "Nome Ceasa", w: "nome" });
  return c;
}
function linhasCatalogo() {
  const q = norm($("#cat-busca")?.value);
  return catLinhas.filter((r) => !r.c || ((catFiltro === "Todos" || grupoFornecedor(r.v.fornecedor) === catFiltro)
    && (!q || norm(`${r.v.descricao} ${r.v.nome_personalizado} ${r.v.nome_ceasa} ${r.v.codigo_erp} ${r.v.codigo_iceasa} ${r.v.codigo_barra}`).includes(q))));
}
function renderCatalogo() {
  const grupos = ordenarGruposCat();
  const nomes = cat.nomes;
  $("#conteudo").innerHTML = `<div class="pagina">
    ${cat.espelho ? `<p class="faixa info"><span><b>Catálogo-mestre:</b> tudo o que você salvar aqui é replicado automaticamente no <b>${esc(cat.espelho)}</b> — inclusões, nomes, fornecedor, códigos, exclusões e as marcações de loja. Não precisa repetir lá.</span></p>` : ""}
    ${cat.mestre ? `<p class="faixa aviso"><span>O cadastro deste setor é comandado pelo <b>${esc(cat.mestre)}</b>. Alterações feitas aqui <b>não sobem</b> para o ${esc(cat.mestre)} e podem ser sobrescritas no próximo salvamento de lá. Para mexer no catálogo do FLV, use o ${esc(cat.mestre)}.</span></p>` : ""}
    <div class="cartao">
      <div class="barra-filtros">
        <button class="btn btn-secundario btn-pequeno" type="button" id="cat-novo">+ Novo produto</button>
        ${grupos.length > 1 ? `<div class="chips" id="cat-chips"><button type="button" class="sel" data-f="Todos">Todos</button>${grupos.map((g) => `<button type="button" data-f="${esc(g)}">${esc(g)}</button>`).join("")}</div>` : ""}
        <input class="busca" id="cat-busca" type="search" placeholder="Buscar por nome ou código" autocomplete="off">
        <span class="metrica" id="cat-metrica"></span>
      </div>
      <p class="dica" style="margin:-4px 0 12px">Desmarque a loja que <b>não</b> pode pedir o produto. Clique no nome da loja (L01, L02...) para marcar ou desmarcar todos os produtos da lista de uma vez.
        ${nomes ? ` Nome Prime atualizado pelo ERP em ${esc(nomes.em)}.` : ""}</p>
      <div class="tabela-wrap grade-wrap"><table class="grade catalogo" id="cat-grade"></table></div>
    </div>
  </div>`;
  $("#cat-chips")?.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-f]"); if (!b) return;
    catFiltro = b.dataset.f;
    $$("#cat-chips button").forEach((x) => x.classList.toggle("sel", x === b));
    desenharCatalogo();
  });
  $("#cat-busca").addEventListener("input", desenharCatalogo);
  $("#cat-novo").onclick = () => {
    const forn = catFiltro !== "Todos" ? catFiltro : "";
    catLinhas.unshift({ id: ++catSeq, c: null, v: { fornecedor: forn, codigo_erp: "", codigo_iceasa: "", codigo_barra: "", link_foto: "", descricao: "", nome_personalizado: "", nome_ceasa: "" }, o: {}, l: Array(8).fill(true), lo: Array(8).fill(true), x: false });
    desenharCatalogo();
    $(`#cat-grade tr[data-id="${catSeq}"] input[data-campo="${cR.erp ? "codigo_erp" : "descricao"}"]`)?.focus();
  };
  const g = $("#cat-grade");
  g.addEventListener("input", (e) => {
    const i = e.target; const r = catLinhas.find((x) => String(x.id) === i.closest("tr")?.dataset.id); if (!r) return;
    if (i.dataset.campo) {
      if (["codigo_erp", "codigo_iceasa"].includes(i.dataset.campo)) { const l = i.value.replace(/\D/g, ""); if (l !== i.value) i.value = l; }
      r.v[i.dataset.campo] = i.value;
      i.classList.toggle("mudou", !!r.c && strV(i.value) !== strV(r.o[i.dataset.campo]));
    }
    resumoCatalogo();
  });
  g.addEventListener("change", (e) => {
    const i = e.target; const r = catLinhas.find((x) => String(x.id) === i.closest("tr")?.dataset.id); if (!r) return;
    if (i.dataset.loja !== undefined) { r.l[Number(i.dataset.loja)] = i.checked; i.closest("td").classList.toggle("mudou", !!r.c && r.l[i.dataset.loja] !== r.lo[i.dataset.loja]); resumoCatalogo(); }
  });
  g.addEventListener("click", (e) => {
    const b = e.target.closest("[data-excluir]");
    if (b) {
      const r = catLinhas.find((x) => String(x.id) === b.dataset.excluir);
      if (!r.c) catLinhas = catLinhas.filter((x) => x !== r); else r.x = !r.x;
      desenharCatalogo(); return;
    }
    const th = e.target.closest("th[data-todas]");
    if (th) {
      const l = Number(th.dataset.todas);
      const vis = linhasCatalogo().filter((r) => !r.x);
      const marcar = !vis.every((r) => r.l[l]);
      vis.forEach((r) => (r.l[l] = marcar));
      desenharCatalogo();
    }
  });
  g.addEventListener("keydown", (e) => {
    const i = e.target.closest("input[data-campo]"); if (!i || e.key !== "Enter") return;
    e.preventDefault();
    const td = i.closest("td"), idx = [...td.parentElement.children].indexOf(td);
    const prox = td.parentElement.nextElementSibling?.children[idx]?.querySelector("input");
    if (prox) { prox.focus(); prox.select(); }
  });
  desenharCatalogo();
  barraCatalogo();
}
function ordenarGruposCat() {
  const mapa = cat.ordem || {};
  return [...new Set(catLinhas.filter((r) => r.c).map((r) => grupoFornecedor(r.v.fornecedor)).filter((x) => String(x).trim()))]
    .sort((a, b) => ((mapa[a] ?? 1e6) - (mapa[b] ?? 1e6)) || a.toLowerCase().localeCompare(b.toLowerCase(), "pt-BR"));
}
function desenharCatalogo() {
  const cols = colunasCatalogo();
  const linhas = linhasCatalogo();
  $("#cat-grade").innerHTML = `<thead><tr>${cols.map((c) => `<th>${c.t}</th>`).join("")}
      ${LOJAS.map((l, i) => `<th class="num-col clicavel" data-todas="${i}" title="Marcar/desmarcar ${l} em todos os produtos da lista">L${p2(i + 1)}</th>`).join("")}<th></th></tr></thead>
    <tbody>${linhas.map((r) => `<tr data-id="${r.id}" class="${r.x ? "excluir" : !r.c ? "novo" : ""}">
      ${cols.map((c) => {
        const mudou = r.c && strV(r.v[c.k]) !== strV(r.o[c.k]);
        const ph = !r.c && c.k === "codigo_erp" ? "obrigatório" : !r.c && c.k === "descricao" ? "nome do produto" : "";
        return `<td><input class="cat-${c.w} ${mudou ? "mudou" : ""}" data-campo="${c.k}" value="${esc(r.v[c.k])}" ${ph ? `placeholder="${ph}"` : ""} ${["codigo_erp", "codigo_iceasa"].includes(c.k) ? 'inputmode="numeric"' : ""} ${r.x ? "disabled" : ""} autocomplete="off"></td>`;
      }).join("")}
      ${r.l.map((v, i) => `<td class="ck ${r.c && v !== r.lo[i] ? "mudou" : ""}"><input type="checkbox" data-loja="${i}" ${v ? "checked" : ""} ${r.x ? "disabled" : ""} aria-label="${nomeLoja(i + 1)}"></td>`).join("")}
      <td><button class="btn-lixo ${r.x ? "ativo" : ""}" type="button" data-excluir="${r.id}" title="${r.x ? "Desfazer exclusão" : "Excluir produto"}">${r.x ? "desfazer" : "excluir"}</button></td>
    </tr>`).join("") || `<tr><td colspan="${cols.length + 9}" class="vazio">Nenhum produto com esse filtro.</td></tr>`}</tbody>`;
  resumoCatalogo();
}
function resumoCatalogo() {
  const m = $("#cat-metrica");
  if (m) m.innerHTML = `<b class="num">${catLinhas.filter((r) => r.c).length}</b> produtos no setor`;
  const r = $("#c-resumo"); if (!r) return;
  const exc = catLinhas.filter((x) => x.x).length, nov = catLinhas.filter((x) => !x.c && !novoVazio(x)).length;
  const ed = catLinhas.filter((x) => x.c && !x.x && (camposMudados(x).length || lojasMudaram(x))).length;
  const partes = [ed && `${ed} ${ed === 1 ? "alterado" : "alterados"}`, nov && `${nov} ${nov === 1 ? "novo" : "novos"}`, exc && `${exc} para excluir`].filter(Boolean);
  r.innerHTML = partes.length
    ? `<strong style="color:var(--laranja)">${partes.join(" · ")}</strong><small>Clique em Salvar catálogo para gravar${cat.espelho ? " (e replicar no " + esc(cat.espelho) + ")" : ""}.</small>`
    : "<strong>Catálogo salvo</strong><small>Nenhuma alteração pendente.</small>";
}
function barraCatalogo() {
  const b = $("#barra-total");
  b.hidden = false; document.body.classList.add("com-barra");
  b.innerHTML = `<div class="resumo" id="c-resumo"></div>
    <button class="btn btn-secundario" type="button" id="cat-ordem">Ordenar fornecedores</button>
    ${cR.erp ? '<button class="btn btn-secundario" type="button" id="cat-nomes">Puxar nomes do ERP</button>' : ""}
    ${cat.espelho ? `<button class="btn btn-secundario" type="button" id="cat-espelho">Sincronizar ${esc(cat.espelho)}</button>` : ""}
    <button class="btn btn-secundario" type="button" id="cat-imprimir">Imprimir</button>
    <button class="btn btn-primario" type="button" id="cat-salvar">Salvar catálogo</button>`;
  $("#cat-salvar").onclick = salvarCatalogo;
  $("#cat-ordem").onclick = janelaOrdem;
  $("#cat-nomes")?.addEventListener("click", pedirNomes);
  $("#cat-espelho")?.addEventListener("click", janelaEspelho);
  $("#cat-imprimir").onclick = imprimirCatalogo;
  resumoCatalogo();
}
async function salvarCatalogo() {
  const novos = catLinhas.filter((r) => !r.c && !r.x && !novoVazio(r));
  const semErp = cR.erp ? novos.filter((r) => !strV(r.v.codigo_erp)) : [];
  if (semErp.length && !confirm(`${semErp.length} produto(s) novo(s) sem Cód. ERP: vão ganhar um código interno novo e ficam sem estoque/média do ERP${cat.espelho ? " e sem espelho no " + cat.espelho : ""}. Salvar assim mesmo?`)) return;
  const exc = catLinhas.filter((r) => r.c && r.x);
  if (exc.length && !confirm(`Excluir ${exc.length} produto(s)? Isso apaga também os pedidos pendentes e as médias deles${cat.espelho ? ", e o mesmo produto no " + cat.espelho : ""}. Não pode ser desfeito.`)) return;
  const dados = {
    excluir: exc.map((r) => r.c),
    editar: catLinhas.filter((r) => r.c && !r.x && camposMudados(r).length).map((r) => ({ c: r.c, campos: Object.fromEntries(camposMudados(r).map((k) => [k, strV(r.v[k])])) })),
    novos: novos.map((r) => ({ campos: Object.fromEntries(CAMPOS_CAT.map((k) => [k, strV(r.v[k])])), l: r.l })),
    lojas: catLinhas.filter((r) => r.c && !r.x && lojasMudaram(r)).map((r) => ({ c: r.c, l: r.l })),
  };
  if (!dados.excluir.length && !dados.editar.length && !dados.novos.length && !dados.lojas.length) { toast("Nada para salvar."); return; }
  const btn = $("#cat-salvar");
  btn.disabled = true; btn.textContent = "Gravando...";
  try {
    const r = await rpc("pedidos_web_catalogo_salvar", { p_token: sessao.token, p_setor: cat.setor, p_dados: dados });
    cd = null; centralPainel = null;
    await abrirCatalogo(cSetor, true);
    const esp = r.espelho ? ` Replicado no ${r.espelho}: ${r.esp_novos} incluído(s), ${r.esp_editados} alterado(s), ${r.esp_lojas} com lojas, ${r.esp_excluidos} excluído(s).${r.esp_sem_erp ? ` ${r.esp_sem_erp} sem Cód. ERP não foram espelhados.` : ""}` : "";
    toast(`Catálogo salvo: ${r.novos} novo(s), ${r.editados} alterado(s), ${r.lojas} com lojas, ${r.excluidos} excluído(s).${esp}`, "ok");
  } catch (err) {
    toast(err.message, "erro");
    btn.disabled = false; btn.textContent = "Salvar catálogo";
  }
}
function janelaOrdem() {
  // nomes completos dos fornecedores (como no Streamlit): valem para a Visão Fornecedores, o Excel e a impressão
  const forns = [...new Set(catLinhas.filter((r) => r.c).map((r) => strV(r.v.fornecedor)).filter((x) => x))];
  const mapa = cat.ordem || {};
  forns.sort((a, b) => ((mapa[a] ?? 1e6) - (mapa[b] ?? 1e6)) || a.toLowerCase().localeCompare(b.toLowerCase(), "pt-BR"));
  abrirJanela(`<h3>Ordenar fornecedores</h3>
    <p>Menor número aparece primeiro. Fornecedor sem número vai para o fim, em ordem alfabética. A ordem vale para a tela, a impressão e o Excel deste setor.</p>
    <div class="lista-ordem">${forns.map((f) => `<label><span>${esc(f)}</span><input type="number" min="0" step="1" data-f="${esc(f)}" value="${mapa[f] ?? 999}"></label>`).join("")}</div>
    <div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Cancelar</button><button class="btn btn-primario" type="button" id="ordem-salvar">Salvar ordem</button></div>`);
  $("#ordem-salvar").onclick = async (e) => {
    const b = e.currentTarget; b.disabled = true;
    try {
      const itens = $$(".lista-ordem input").map((i) => ({ f: i.dataset.f, o: i.value }));
      await rpc("pedidos_web_ordem_salvar", { p_token: sessao.token, p_setor: cat.setor, p_itens: itens });
      fecharJanela();
      cat.ordem = Object.fromEntries(itens.map((x) => [x.f, Number(x.o) || 999]));
      if (cd && cd.setor === cat.setor) cd.ordem = cat.ordem;
      renderCatalogo();
      toast("Ordem dos fornecedores salva.", "ok");
    } catch (err) { b.disabled = false; toast(err.message, "erro"); }
  };
}
async function pedirNomes() {
  if (!confirm("Atualizar o Nome Prime de todos os produtos pelo cadastro do ERP? A rotina do servidor faz isso em até 30 minutos. O Nome Manual não é alterado.")) return;
  try {
    await rpc("pedidos_web_pedir_nomes", { p_token: sessao.token, p_setor: cat.setor });
    toast("Pedido registrado: o servidor atualiza os nomes na próxima rodada do estoque (até 30 min).", "ok");
  } catch (err) { toast(err.message, "erro"); }
}
function janelaEspelho() {
  abrirJanela(`<h3>Sincronizar ${esc(cat.espelho)}</h3>
    <p>Copia o catálogo do <b>${esc(cat.setor)}</b> por cima do <b>${esc(cat.espelho)}</b>: nomes, fornecedor, códigos e marcações de loja. O ${esc(cat.setor)} não é alterado.</p>
    <p>Use para acertar divergências antigas; no dia a dia o Salvar catálogo já espelha sozinho.</p>
    <label class="marcar"><input type="checkbox" id="esp-sobras"> Excluir do ${esc(cat.espelho)} os produtos que não existem no ${esc(cat.setor)}</label>
    <p class="dica">Cuidado: excluir apaga também os pedidos pendentes e as médias desses produtos no ${esc(cat.espelho)}.</p>
    <div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Cancelar</button><button class="btn btn-primario" type="button" id="esp-sim">Sincronizar agora</button></div>`);
  $("#esp-sim").onclick = async (e) => {
    if (catalogoSujo() && !confirm("Há alterações não salvas no catálogo. Elas não entram na sincronização. Continuar?")) return;
    const b = e.currentTarget; b.disabled = true; b.textContent = "Alinhando...";
    try {
      const r = await rpc("pedidos_web_catalogo_espelhar", { p_token: sessao.token, p_setor: cat.setor, p_excluir_sobras: $("#esp-sobras").checked });
      fecharJanela();
      toast(`${r.espelho} sincronizado: ${r.criados} incluído(s), ${r.atualizados} atualizado(s), ${r.removidos} removido(s).${r.sem_erp ? ` ${r.sem_erp} sem Cód. ERP ficaram de fora.` : ""}`, "ok");
    } catch (err) { b.disabled = false; b.textContent = "Sincronizar agora"; toast(err.message, "erro"); }
  };
}
function imprimirCatalogo() {
  const cols = colunasCatalogo();
  const linhas = linhasCatalogo().filter((r) => r.c && !r.x);
  const cab = [...cols.map((c) => c.t), ...LOJAS];
  const corpo = linhas.map((r) => [...cols.map((c) => strV(r.v[c.k])), ...r.l.map((v) => (v ? "✓" : "—"))]);
  imprimirHtml(`<h3>🗂️ Catálogo Geral — ${esc(cat.setor)}</h3><p class="emitido">Emitido em ${dataHoraBR()}</p>
    ${tabelaPrint(cab, corpo, cols.map((c, i) => (["fornecedor", "descricao", "nome_personalizado", "nome_ceasa", "link_foto"].includes(c.k) ? i : -1)).filter((i) => i >= 0))}`, "landscape", "forn");
}

// clique na miniatura da Separação/Fornecedores amplia a foto
document.addEventListener("click", (e) => {
  const img = e.target.closest(".grade img[data-ampliar]");
  if (img) abrirJanela(`<img class="foto-grande" src="${esc(img.src)}" alt=""><div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Fechar</button></div>`);
});
