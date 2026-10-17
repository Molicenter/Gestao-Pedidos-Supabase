// ============================================================
// Gestão de Pedidos · Molicenter (app web, etapa 1: painel + digitação das lojas)
// Grava nas MESMAS tabelas do Streamlit, pelas funções pedidos_web_* do Supabase.
// ============================================================
const CFG = window.CONFIG;
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const norm = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
const LOJAS = Array.from({ length: 8 }, (_, i) => `Loja ${String(i + 1).padStart(2, "0")}`);
const nomeLoja = (n) => `Loja ${String(n).padStart(2, "0")}`;

// ---------- setores (nome = exatamente como está no banco) ----------
const UNSPLASH = (id) => `https://images.unsplash.com/${id}?w=480&q=70`;
const GRUPOS = [
  { id: "verde", nome: "Setor Hortifruti (FLV)" },
  { id: "coral", nome: "Setor Açougue e Aves" },
  { id: "ambar", nome: "Outros Setores e Logística" },
];
// prazos: dias da semana (0 = domingo) e horário-limite. O horário é só informativo,
// igual ao Streamlit: o app não bloqueia a digitação depois do prazo.
const SETORES = [
  { nome: "Folhagem", titulo: "Folhagem", g: "verde", ic: "🥬", foto: UNSPLASH("photo-1574316071802-0d684efa7bf5"), prazo: "Seg a Sáb até 12:00", limites: [{ d: [1, 2, 3, 4, 5, 6], h: "12:00" }] },
  { nome: "FLV Normal", titulo: "FLV Normal", g: "verde", ic: "🍎", foto: UNSPLASH("photo-1610348725531-843dff563e2c"), prazo: "Terça até 17:00 · Quinta até 14:00", limites: [{ d: [2], h: "17:00" }, { d: [4], h: "14:00" }] },
  { nome: "FLV Ofertas", titulo: "FLV Ofertas", g: "verde", ic: "🏷️", foto: UNSPLASH("photo-1610348725531-843dff563e2c"), prazo: "Quinta até 14:00", limites: [{ d: [4], h: "14:00" }] },
  { nome: "FLV Oriental", titulo: "FLV Oriental", g: "verde", ic: "🍣", foto: UNSPLASH("photo-1579871494447-9811cf80d66c"), prazo: "Quinta até 14:00", limites: [{ d: [4], h: "14:00" }] },
  { nome: "Açougue Especiais", titulo: "Pioneiro + BF + Paraná", g: "coral", ic: "🍗", foto: "pioneiros.jpg", prazo: "Seg a Sex até 08:30", limites: [{ d: [1, 2, 3, 4, 5], h: "08:30" }] },
  { nome: "Açougue Adriano", titulo: "Açougue Adriano", g: "coral", ic: "🔪", foto: UNSPLASH("photo-1544025162-d76694265947"), prazo: "Quarta até 15:00 · Sábado até 15:00", limites: [{ d: [3, 6], h: "15:00" }] },
  { nome: "Peças Açougue - Manoel", titulo: "Peças Açougue - Manoel", g: "coral", ic: "🥩", foto: UNSPLASH("photo-1544025162-d76694265947"), prazo: "Seg/Qua/Sex Arapongas 15:00 · Ter/Qui/Sáb Maringá 15:00", limites: [{ d: [1, 2, 3, 4, 5, 6], h: "15:00" }] },
  { nome: "Embalagens", titulo: "Embalagens", g: "ambar", ic: "🥡", foto: "embalagens.jpg", prazo: "Sexta até 17:30", limites: [{ d: [5], h: "17:30" }] },
  { nome: "Padaria e Confeitaria", titulo: "Padaria e Confeitaria", g: "ambar", ic: "🥖", foto: UNSPLASH("photo-1509440159596-0249088772ff"), prazo: "Sábado", limites: [] },
  { nome: "Matéria Prima", titulo: "Matéria Prima", g: "ambar", ic: "🌾", foto: "materiaprima.jpg", prazo: "Até sábado", limites: [] },
  { nome: "Flores", titulo: "Flores", g: "ambar", ic: "🌸", foto: UNSPLASH("photo-1490750967868-88aa4486c946"), prazo: "Segunda (15 em 15 dias) até 17:30", limites: [] },
];
const slug = (s) => norm(s).replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
const setorPorSlug = (sl) => SETORES.find((s) => slug(s.nome) === sl);

// regras por setor (as mesmas funções do tela_pedidos_generica.py)
function regras(setor) {
  const n = norm(setor);
  const texto = /materia prima|embalag|padaria|confeitaria/.test(n);
  const manoel = n.includes("manoel");
  const flores = n.includes("flores");
  const oriental = n.includes("oriental");
  return {
    texto,                                   // pede em texto ("1 cx", "2 fd")
    obs: /embalag|padaria|confeitaria/.test(n), // Observação Geral da Loja
    semPedido: n === "acougue adriano" || n === "acougue especiais",
    media: !texto && !manoel,                // coluna Média
    erp: !manoel,                            // Cód. ERP e Estoque
    flores,                                  // Cód. Barra no lugar do Fornecedor
    foto: flores || oriental,                // coluna Foto
    miniatura: oriental,                     // Oriental: miniatura · Flores: link "Ver"
  };
}
// agrupa "BIG FRANGO - *" e "PIONEIRO - *" num rótulo só no filtro
function grupoFornecedor(f) {
  const n = norm(f);
  if (n.startsWith("big frango")) return "BIG FRANGO";
  if (n.startsWith("pioneiro")) return "PIONEIRO";
  return String(f || "").trim();
}

// ---------- utilidades ----------
function agoraSP() { return new Date(new Date().toLocaleString("en-US", { timeZone: "America/Sao_Paulo" })); }
const DIAS = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const p2 = (n) => String(n).padStart(2, "0");
function dataHoraBR(d = agoraSP()) { return `${p2(d.getDate())}/${p2(d.getMonth() + 1)}/${d.getFullYear()} ${p2(d.getHours())}:${p2(d.getMinutes())}`; }
function horaBR(d = agoraSP()) { return `${p2(d.getHours())}:${p2(d.getMinutes())}`; }
const fmtNum = (v, casas = 0) => Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas });
// mesma regra do Streamlit: "2,5" → 2; vazio/inválido → 0
function qtdInt(v) {
  const t = String(v ?? "").trim().replace(",", ".");
  if (!t || ["-", "nan", "none", "null", "<na>"].includes(t.toLowerCase())) return 0;
  const n = Number(t);
  return Number.isFinite(n) ? Math.max(0, Math.trunc(n)) : 0;
}
const temValor = (v, texto) => (texto ? String(v ?? "").trim() !== "" : qtdInt(v) > 0);

function toast(msg, tipo = "") {
  const t = $("#toast");
  t.textContent = msg; t.className = "toast " + tipo; t.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), tipo === "erro" ? 6000 : 3200);
}
function carregando(sim) { $("#carregando").hidden = !sim; }

// ---------- sessão ----------
const CHAVE_SESSAO = "pedidos_web_sessao";
function lerSessao() { try { return JSON.parse(localStorage.getItem(CHAVE_SESSAO)) || null; } catch { return null; } }
function gravarSessao(s) { try { s ? localStorage.setItem(CHAVE_SESSAO, JSON.stringify(s)) : localStorage.removeItem(CHAVE_SESSAO); } catch {} }
let sessao = lerSessao();
let lojaAtual = null;

// ---------- Supabase ----------
const MENSAGENS = {
  SENHA_INCORRETA: "Senha incorreta. Tente de novo.",
  SENHA_NAO_CONFIGURADA: "A senha deste usuário ainda não foi cadastrada no banco (falta rodar o SQL 02).",
  MUITAS_TENTATIVAS: "Muitas tentativas erradas. Espere 10 minutos e tente de novo.",
  USUARIO_INVALIDO: "Escolha um usuário.",
  SESSAO_INVALIDA: "Sua sessão expirou. Entre de novo.",
  LOJA_INVALIDA: "Escolha a loja.",
  SETOR_INVALIDO: "Setor não encontrado.",
  PRODUTO_INVALIDO: "Produto não encontrado neste setor. Atualize a tela.",
  SEM_PEDIDO_NAO_PERMITIDO: "Este setor não tem a opção Sem Pedido Hoje.",
  TELEGRAM_NAO_CONFIGURADO: "O aviso do Telegram ainda não foi configurado (falta rodar o SQL 02).",
  TELEGRAM_FALHOU: "O Telegram não aceitou o aviso agora.",
  "Could not find the function": "O banco ainda não tem as funções do app web (falta rodar o SQL 01).",
};
function traduzir(msg) {
  for (const k in MENSAGENS) if (String(msg).includes(k)) return MENSAGENS[k];
  return msg;
}
async function rpc(fn, args = {}) {
  let resp;
  try {
    resp = await fetch(`${CFG.SUPABASE_URL}/rest/v1/rpc/${fn}`, {
      method: "POST",
      headers: { apikey: CFG.SUPABASE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify(args),
    });
  } catch {
    throw new Error("Sem conexão com a internet. Tente de novo.");
  }
  const texto = await resp.text();
  let corpo = null;
  try { corpo = texto ? JSON.parse(texto) : null; } catch { corpo = texto; }
  if (!resp.ok) {
    const msg = (corpo && (corpo.message || corpo.hint)) || `Erro ${resp.status}`;
    if (String(msg).includes("SESSAO_INVALIDA")) { sair(true); throw new Error(MENSAGENS.SESSAO_INVALIDA); }
    throw new Error(traduzir(msg));
  }
  return corpo;
}
async function avisos(action, extra) {
  const resp = await fetch(`${CFG.SUPABASE_URL}/functions/v1/pedidos-avisos`, {
    method: "POST",
    headers: { apikey: CFG.SUPABASE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ action, token: sessao?.token, ...extra }),
  });
  const j = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(traduzir(j.erro || `Erro ${resp.status}`));
  return j;
}

// ============================================================
// Login
// ============================================================
function montarLogin() {
  const sel = $("#login-usuario");
  sel.innerHTML = '<option value="">Selecione...</option>' + ["Administrador", "Ceasa", ...LOJAS].map((u) => `<option>${u}</option>`).join("");
  try { const u = localStorage.getItem("pedidos_web_ultimo_usuario"); if (u) sel.value = u; } catch {}
  $("#versao-login").textContent = "versão " + CFG.VERSAO;
}
$("#form-login").addEventListener("submit", async (e) => {
  e.preventDefault();
  const usuario = $("#login-usuario").value;
  const senha = $("#login-senha").value;
  const erro = $("#login-erro");
  erro.hidden = true;
  if (!usuario) { erro.textContent = "Escolha o usuário."; erro.hidden = false; return; }
  if (!senha) { erro.textContent = "Digite a senha."; erro.hidden = false; return; }
  const btn = $("#btn-entrar");
  btn.disabled = true; btn.textContent = "Entrando...";
  try {
    sessao = await rpc("pedidos_web_login", { p_usuario: usuario, p_senha: senha });
    gravarSessao(sessao);
    try { localStorage.setItem("pedidos_web_ultimo_usuario", usuario); } catch {}
    $("#login-senha").value = "";
    iniciarApp();
  } catch (err) {
    erro.textContent = err.message; erro.hidden = false;
  } finally {
    btn.disabled = false; btn.textContent = "Entrar";
  }
});
async function sair(silencioso) {
  const tk = sessao?.token;
  sessao = null; gravarSessao(null);
  if (tk && !silencioso) rpc("pedidos_web_sair", { p_token: tk }).catch(() => {});
  $("#app").hidden = true; $("#barra-total").hidden = true; document.body.classList.remove("com-barra");
  $("#tela-login").hidden = false;
  montarLogin();
}
$("#btn-sair").addEventListener("click", async () => {
  if (temAlteracaoPendente() && !confirm("Ainda há alterações sendo salvas. Sair mesmo assim?")) return;
  sair(false);
});

// ============================================================
// Casca: menu, topo, rotas
// ============================================================
let painel = null;          // retorno de pedidos_web_painel
let telaAtual = "painel";   // painel | setor

function iniciarApp() {
  $("#tela-login").hidden = true;
  $("#app").hidden = false;
  const admin = sessao.perfil !== "loja";   // Administrador ou Ceasa (central)
  if (admin) {
    let salva = null;
    try { salva = Number(localStorage.getItem("pedidos_web_loja_admin")); } catch {}
    lojaAtual = salva >= 1 && salva <= 8 ? salva : 1;
  } else {
    lojaAtual = sessao.loja;
  }
  $("#conta-nome").textContent = sessao.usuario;
  $("#conta-perfil").textContent = sessao.perfil === "ceasa" ? "Ceasa · central de pedidos" : admin ? "Administrador · central de pedidos" : "Pedidos da loja";
  montarPilulaLoja();
  const d = agoraSP();
  $("#pil-data").textContent = `${DIAS[d.getDay()]}, ${p2(d.getDate())}/${p2(d.getMonth() + 1)}`;
  rota();
}
function montarPilulaLoja() {
  const pil = $("#pil-loja");
  if (sessao.perfil !== "loja") {
    pil.innerHTML = `Ver como <select id="sel-loja" aria-label="Visualizar como">${LOJAS.map((l, i) => `<option value="${i + 1}" ${i + 1 === lojaAtual ? "selected" : ""}>${l}</option>`).join("")}</select>`;
    $("#sel-loja").addEventListener("change", (e) => {
      if (temAlteracaoPendente() && !confirm("Ainda há alterações sendo salvas. Trocar de loja mesmo assim?")) { e.target.value = lojaAtual; return; }
      lojaAtual = Number(e.target.value);
      try { localStorage.setItem("pedidos_web_loja_admin", lojaAtual); } catch {}
      rota();
    });
  } else {
    pil.textContent = nomeLoja(lojaAtual);
  }
}

function situacaoSetor(nome) {
  const st = painel?.setores?.[nome];
  if ((painel?.sem_pedido || []).includes(nome)) return { cls: "sem", txt: "Sem pedido hoje", mini: "sem" };
  if (!st || !st.itens) return { cls: "novo", txt: "Não iniciado", mini: "" };
  if (st.rascunho > 0) return { cls: "rasc", txt: `Rascunho · ${st.itens} ${st.itens === 1 ? "item" : "itens"}`, mini: "rasc" };
  return { cls: "ok", txt: `Salvo ${st.ultimo || ""} · ${st.itens} ${st.itens === 1 ? "item" : "itens"}`, mini: "ok" };
}
function renderNav() {
  if (ehCentral()) return renderNavCentral();
  const ativo = telaAtual === "setor" ? setorAberto?.nome : null;
  const icPainel = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>';
  $("#nav").innerHTML = `<div class="nav-rotulo">Início</div>
    <button type="button" data-ir="painel" class="${telaAtual === "painel" ? "ativo" : ""}">${icPainel}Painel de pedidos</button>
    <div class="nav-rotulo">Setores</div>
    ${SETORES.map((s) => {
      const st = situacaoSetor(s.nome);
      const mini = st.mini === "ok" ? '<span class="estado-mini ok">ok</span>' : st.mini === "rasc" ? '<span class="estado-mini rasc">rasc.</span>' : st.mini === "sem" ? '<span class="estado-mini sem">sem</span>' : "";
      return `<button type="button" data-setor="${slug(s.nome)}" class="${ativo === s.nome ? "ativo" : ""}"><span class="bolinha" style="background:var(--g-${s.g})"></span>${esc(s.titulo)}${mini}</button>`;
    }).join("")}`;
}
document.addEventListener("click", (e) => {
  const b = e.target.closest("[data-setor],[data-ir]");
  if (!b || b.closest(".janela")) return;
  if (b.dataset.setor) irPara("#/setor/" + b.dataset.setor);
  else if (b.dataset.ir === "painel") irPara("#/painel");
});
function irPara(hash) {
  if (location.hash === hash) rota(); else location.hash = hash;
}
window.addEventListener("hashchange", () => { if (sessao) rota(); });

async function rota() {
  const m = location.hash.match(/^#\/setor\/([a-z0-9-]+)(?:\/([a-z]+))?/);
  const s = m && setorPorSlug(m[1]);
  if (ehCentral()) { await rotaCentral(s, m && m[2]); window.scrollTo({ top: 0 }); return; }
  $("#abas").hidden = true; $("#pil-loja").hidden = false;
  if (s) await abrirSetor(s); else await abrirPainel();
  window.scrollTo({ top: 0 });
}

// ============================================================
// Painel de pedidos (cartões dos 11 setores)
// ============================================================
function prazoHoje(s) {
  const agora = agoraSP();
  const lim = s.limites.find((l) => l.d.includes(agora.getDay()));
  if (!lim) return null;
  const [h, mi] = lim.h.split(":").map(Number);
  const fim = new Date(agora); fim.setHours(h, mi, 0, 0);
  const min = Math.round((fim - agora) / 60000);
  if (min < 0) return { txt: `Prazo de hoje: ${lim.h}`, passou: true };
  if (min <= 180) return { txt: `Fecha em ${min >= 60 ? Math.floor(min / 60) + "h" + p2(min % 60) : min + " min"}`, passou: false, min };
  return { txt: `Hoje até ${lim.h}`, passou: false, min };
}
let timerPainel = null;
async function abrirPainel() {
  if (!(await podeSairDoSetor())) return;
  telaAtual = "painel"; setorAberto = null;
  $("#barra-total").hidden = true; document.body.classList.remove("com-barra");
  $("#pg-titulo").textContent = "Painel de pedidos";
  $("#pg-sub").textContent = "Escolha o setor para fazer o pedido da loja. Cada cartão mostra o prazo e como está o pedido.";
  if (!painel) $("#conteudo").innerHTML = '<div class="vazio">Carregando...</div>';
  try {
    painel = await rpc("pedidos_web_painel", { p_token: sessao.token, p_loja: lojaAtual });
  } catch (err) {
    toast(err.message, "erro");
    if (!painel) return;
  }
  renderNav();
  renderPainel();
  clearInterval(timerPainel);
  timerPainel = setInterval(async () => {
    if (telaAtual !== "painel" || document.hidden || !sessao) return;
    try { painel = await rpc("pedidos_web_painel", { p_token: sessao.token, p_loja: lojaAtual }); renderNav(); renderPainel(); } catch {}
  }, 120000);
}
function renderPainel() {
  const sit = SETORES.map((s) => ({ s, st: situacaoSetor(s.nome), pz: prazoHoje(s) }));
  const enviados = sit.filter((x) => x.st.cls === "ok" || x.st.cls === "sem").length;
  const rasc = sit.filter((x) => x.st.cls === "rasc");
  const hoje = sit.filter((x) => x.pz && !x.pz.passou).sort((a, b) => a.pz.min - b.pz.min);
  const prox = hoje[0];
  $("#conteudo").innerHTML = `<div class="pagina">
    <div class="kpis">
      <div class="kpi destaque"><small>Prazos hoje</small><strong class="num">${hoje.length}</strong><span>${hoje.length ? esc(hoje.map((x) => x.s.titulo).join(", ")) : "nenhum setor fecha hoje"}</span></div>
      <div class="kpi verde"><small>Pedidos salvos</small><strong class="num">${enviados}</strong><span>de ${SETORES.length} setores (inclui sem pedido)</span></div>
      <div class="kpi laranja"><small>Em rascunho</small><strong class="num">${rasc.length}</strong><span>${rasc.length ? "falta clicar em Salvar pedido" : "nada esperando confirmação"}</span></div>
      <div class="kpi magenta"><small>Próximo prazo</small><strong class="num">${prox ? prox.s.limites.find((l) => l.d.includes(agoraSP().getDay())).h : "—"}</strong><span>${prox ? esc(prox.s.titulo) + " · " + esc(prox.pz.txt.toLowerCase()) : "sem prazo pela frente hoje"}</span></div>
    </div>
    ${painel?.estoque_em ? `<p class="aviso-estoque">Estoque do ERP atualizado em ${esc(painel.estoque_em)} (atualiza sozinho a cada 30 minutos, das 6h às 22h).</p>` : ""}
    ${GRUPOS.map((g) => `<section class="grupo" style="--cor:var(--g-${g.id})">
      <h3 class="grupo-titulo" style="color:var(--g-${g.id}-texto)">${g.nome}</h3>
      <div class="cards">${sit.filter((x) => x.s.g === g.id).map(({ s, st, pz }) => `
        <button class="setor" type="button" data-setor="${slug(s.nome)}" style="--cor:var(--g-${s.g});--cor-fundo:var(--g-${s.g}-fundo);--cor-texto:var(--g-${s.g}-texto)">
          <span class="setor-faixa"></span>
          <span class="setor-foto" style="background-image:url('${s.foto}')" aria-hidden="true"></span>
          <span class="setor-corpo">
            <span class="setor-nome">${esc(s.titulo)}</span>
            <span class="setor-prazo">${esc(s.prazo)}</span>
            <span class="setor-rodape"><span class="tag ${st.cls}">${esc(st.txt)}</span>${pz ? `<span class="relogio ${pz.passou ? "passou" : ""}">${esc(pz.txt)}</span>` : ""}</span>
          </span>
        </button>`).join("")}</div></section>`).join("")}
  </div>`;
}

// ============================================================
// Digitação da loja (Visão das Lojas)
// ============================================================
let setorAberto = null;   // item de SETORES
let dados = null;         // retorno de pedidos_web_setor
let R = null;             // regras do setor
let valores = {};         // código → valor na tela
let noBanco = {};         // código → valor já gravado (para não regravar à toa)
let filas = {};           // código → promessa da última gravação
let timers = {};          // código → atraso da digitação
let gravando = 0, falhou = null, ultimaGravacao = null;
let tudoConfirmado = true;
let obsOriginal = "";
let filtroForn = "Todos";

function temAlteracaoPendente() {
  if (telaAtual === "central") return centralSujo();
  if (telaAtual !== "setor") return false;
  const obsMudou = R?.obs && $("#obs-loja") && $("#obs-loja").value.trim() !== obsOriginal.trim();
  return gravando > 0 || Object.keys(timers).length > 0 || !!falhou || obsMudou;
}
async function podeSairDoSetor() {
  if (telaAtual !== "setor") return true;
  await descarregar();
  if (temAlteracaoPendente()) {
    const ok = confirm("Há alterações que ainda não foram salvas. Sair mesmo assim?");
    if (!ok) { history.replaceState(null, "", "#/setor/" + slug(setorAberto.nome)); return false; }
  }
  return true;
}
window.addEventListener("beforeunload", (e) => { if (temAlteracaoPendente()) { e.preventDefault(); e.returnValue = ""; } });

async function abrirSetor(s) {
  if (telaAtual === "setor" && setorAberto && setorAberto.nome !== s.nome && !(await podeSairDoSetor())) return;
  telaAtual = "setor"; setorAberto = s; R = regras(s.nome);
  valores = {}; noBanco = {}; filas = {}; timers = {}; gravando = 0; falhou = null; ultimaGravacao = null; filtroForn = "Todos";
  clearInterval(timerPainel);
  $("#pg-titulo").textContent = s.titulo;
  $("#pg-sub").textContent = `Prazo: ${s.prazo}. O que você digita é salvo sozinho; ao terminar, clique em Salvar pedido.`;
  $("#conteudo").innerHTML = '<div class="vazio">Carregando produtos...</div>';
  $("#barra-total").hidden = true;
  renderNav();
  try {
    dados = await rpc("pedidos_web_setor", { p_token: sessao.token, p_loja: lojaAtual, p_setor: s.nome });
  } catch (err) {
    $("#conteudo").innerHTML = `<div class="cartao"><p class="faixa erro">${esc(err.message)}</p></div>`;
    return;
  }
  for (const p of dados.pedido) {
    const v = R.texto ? String(p.t || "").trim() : (Number(p.q) > 0 ? String(Math.trunc(Number(p.q))) : "");
    valores[p.c] = v; noBanco[p.c] = v;
  }
  tudoConfirmado = dados.pedido.every((p) => p.ok !== false);
  obsOriginal = dados.obs || "";
  // ordenação: Fornecedor e Descrição (Flores: só Descrição)
  const cmp = (a, b) => a.localeCompare(b, "pt-BR", { sensitivity: "base" });
  dados.produtos.sort((a, b) => (R.flores ? 0 : cmp(a.f, b.f)) || cmp(a.d, b.d));
  renderSetor();
}

function colunasSetor() {
  const c = [];
  if (R.erp) c.push(["cod", "Cód. ERP", "74px"]);
  c.push(R.flores ? ["barra", "Cód. Barra", "128px"] : ["forn", "Fornecedor", "minmax(110px, .55fr)"]);
  if (R.foto) c.push(["foto", "Foto", R.miniatura ? "64px" : "54px"]);
  c.push(["desc", "Descrição", "minmax(0, 1.6fr)"]);
  if (R.media) c.push(["media", "Média", "68px"]);
  if (R.erp) c.push(["est", "Estoque", "68px"]);
  c.push(["qtd", R.texto ? "Pedido" : "Qtde", R.texto ? "140px" : "150px"]);
  return c;
}
function gruposDoSetor() {
  if (R.flores) return [];
  const mapa = dados.ordem || {};
  const g = [...new Set(dados.produtos.map((p) => grupoFornecedor(p.f)).filter((x) => x))];
  return g.sort((a, b) => ((mapa[a] ?? 1e6) - (mapa[b] ?? 1e6)) || a.toLowerCase().localeCompare(b.toLowerCase(), "pt-BR"));
}

function renderSetor() {
  const cols = colunasSetor();
  const grupos = gruposDoSetor();
  const temProdutos = dados.produtos.length > 0;
  const linhas = dados.produtos.map((p) => {
    const v = valores[p.c] ?? "";
    const cel = {
      cod: `<span class="p-cod num">${esc(p.erp)}</span>`,
      barra: `<span class="p-forn num">${esc(p.b)}</span>`,
      forn: `<span class="p-forn">${esc(p.f)}</span>`,
      foto: `<span class="p-foto">${fotoHtml(p)}</span>`,
      desc: `<span class="p-nome">${esc(p.d)}<small>${esc(R.flores ? p.b : p.f)}${R.erp ? " · " + esc(p.erp) : ""}</small></span>`,
      media: `<span class="p-num num">${fmtNum(p.m, 2)}</span>`,
      est: `<span class="p-num num">${fmtNum(Math.trunc(p.e))}</span>`,
      qtd: `<span class="qtd">${R.texto
        ? `<input class="texto" type="text" data-c="${p.c}" value="${esc(v)}" aria-label="Pedido de ${esc(p.d)}" autocomplete="off">`
        : `<button type="button" data-passo="-1" data-c="${p.c}" aria-label="Menos">−</button><input class="num" type="text" inputmode="numeric" data-c="${p.c}" value="${esc(v)}" aria-label="Quantidade de ${esc(p.d)}" autocomplete="off"><button type="button" data-passo="1" data-c="${p.c}" aria-label="Mais">+</button>`}</span>`,
    };
    const info = [R.media ? `média ${fmtNum(p.m, 2)}` : "", R.erp ? `estoque ${fmtNum(Math.trunc(p.e))}` : ""].filter(Boolean).join(" · ");
    return `<div class="produto" data-linha="${p.c}" data-g="${esc(grupoFornecedor(p.f))}" data-busca="${esc(norm(p.d + " " + (R.erp ? p.erp : "") + " " + (R.flores ? p.b : "")))}">
      ${cols.map(([k]) => cel[k]).join("")}${info ? `<span class="p-info num">${info}</span>` : ""}<span class="aviso-curva" hidden></span></div>`;
  }).join("");

  $("#conteudo").innerHTML = `<div class="pagina">
    ${dados.sem_pedido ? '<p class="faixa ok"><span>Hoje esta loja informou que não fará pedido neste setor. Se lançar itens e salvar, a marca de "sem pedido" sai sozinha.</span></p>' : ""}
    <p class="faixa info" id="faixa-rascunho" hidden><span>Sua digitação está sendo salva automaticamente. Ao terminar, clique em <b>Salvar pedido</b> para confirmar o envio.</span></p>
    <div class="faixa erro" id="faixa-curva" hidden></div>
    <div class="cartao">
      <div class="barra-filtros">
        <button class="voltar" type="button" data-ir="painel">← Painel</button>
        ${grupos.length > 1 ? `<div class="chips" id="chips"><button type="button" class="sel" data-f="Todos">Todos</button>${grupos.map((g) => `<button type="button" data-f="${esc(g)}">${esc(g)}</button>`).join("")}</div>` : ""}
        <input class="busca" id="busca" type="search" placeholder="Buscar produto por nome${R.erp ? " ou código" : ""}${R.flores ? " ou cód. barra" : ""}" autocomplete="off">
        <span class="salvo" id="salvo"></span>
      </div>
      ${temProdutos ? `<div class="lista-produtos" style="--colunas:${cols.map((c) => c[2]).join(" ")}">
        <div class="cab-lista">${cols.map(([k, t]) => `<span class="${["media", "est", "qtd"].includes(k) ? "dir" : ""}">${t}</span>`).join("")}</div>
        ${linhas}
        <p class="vazio" id="nada-filtro" hidden>Nenhum produto com esse filtro.</p>
      </div>` : '<p class="vazio">Nenhum produto liberado para esta loja neste setor.</p>'}
      ${dados.estoque_em && R.erp ? `<p class="aviso-estoque" style="margin-top:12px">Estoque do ERP atualizado em ${esc(dados.estoque_em)}.</p>` : ""}
    </div>
    ${R.obs ? `<div class="cartao"><h3>Observação geral da loja</h3>
      <textarea class="obs-loja" id="obs-loja" placeholder="Ex.: 2 estiletes; 2 grampeadores; 2 sacos de enforca gato...">${esc(obsOriginal)}</textarea>
      <p class="dica" style="margin-top:8px">Recados para a separação (itens fora da lista, pedidos especiais). É salva junto com o pedido.</p></div>` : ""}
  </div>`;

  // eventos da lista
  const lista = $(".lista-produtos");
  if (lista) {
    lista.addEventListener("input", (e) => { const i = e.target.closest("input[data-c]"); if (i) aoDigitar(i, false); });
    lista.addEventListener("change", (e) => { const i = e.target.closest("input[data-c]"); if (i) aoDigitar(i, true); });
    lista.addEventListener("keydown", (e) => {
      const i = e.target.closest("input[data-c]");
      if (!i) return;
      if (e.key === "Enter" || (e.key === "ArrowDown" && !R.texto) || e.key === "ArrowUp" && !R.texto) {
        e.preventDefault();
        aoDigitar(i, true);
        const vis = $$(".produto:not([hidden]) input[data-c]");
        const k = vis.indexOf(i) + (e.key === "ArrowUp" ? -1 : 1);
        if (vis[k]) { vis[k].focus(); vis[k].select(); }
      }
    });
    lista.addEventListener("click", (e) => {
      const b = e.target.closest("button[data-passo]");
      if (b) {
        const i = $(`input[data-c="${b.dataset.c}"]`, lista);
        i.value = String(Math.max(0, qtdInt(i.value) + Number(b.dataset.passo)) || "");
        aoDigitar(i, false);
        return;
      }
      const img = e.target.closest(".p-foto img");
      if (img) abrirJanela(`<img class="foto-grande" src="${esc(img.src)}" alt=""><div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Fechar</button></div>`);
    });
  }
  $("#chips")?.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-f]"); if (!b) return;
    filtroForn = b.dataset.f;
    $$("#chips button").forEach((x) => x.classList.toggle("sel", x === b));
    aplicarFiltro();
  });
  $("#busca").addEventListener("input", aplicarFiltro);
  $("#obs-loja")?.addEventListener("input", atualizarResumo);

  aplicarFiltro();
  renderBarra();
}
function fotoHtml(p) {
  const url = String(p.foto || "").trim();
  if (!/^https?:\/\//.test(url)) return "";
  return R.miniatura ? `<img src="${esc(url)}" alt="" loading="lazy" onerror="this.remove()">` : `<a href="${esc(url)}" target="_blank" rel="noopener">Ver</a>`;
}
function linhasVisiveis() { return $$(".produto:not([hidden])"); }
function aplicarFiltro() {
  const q = norm($("#busca")?.value);
  let n = 0;
  $$(".produto").forEach((l) => {
    const ok = (filtroForn === "Todos" || l.dataset.g === filtroForn) && (!q || l.dataset.busca.includes(q));
    l.hidden = !ok; if (ok) n++;
  });
  const nada = $("#nada-filtro"); if (nada) nada.hidden = n > 0;
}

// ---------- digitação e salvamento automático ----------
function aoDigitar(input, agora) {
  const c = Number(input.dataset.c);
  let v = input.value;
  if (!R.texto) {
    // só números (aceita vírgula, que vira inteiro como no Streamlit)
    const limpo = v.replace(/[^\d.,]/g, "");
    if (limpo !== v) { v = limpo; input.value = limpo; }
  }
  valores[c] = v.trim();
  marcarLinha(c);
  atualizarResumo();
  clearTimeout(timers[c]); delete timers[c];
  if (agora) gravarItem(c);
  else { timers[c] = setTimeout(() => { delete timers[c]; gravarItem(c); }, 800); atualizarSalvo(); }
}
function normalizado(v) { return R.texto ? String(v ?? "").trim() : (qtdInt(v) > 0 ? String(qtdInt(v)) : ""); }
function gravarItem(c) {
  const alvo = normalizado(valores[c]);
  const anterior = filas[c] || Promise.resolve();
  filas[c] = anterior.then(async () => {
    if (normalizado(noBanco[c]) === alvo) return;
    gravando++; atualizarSalvo();
    try {
      await rpc("pedidos_web_salvar_item", { p_token: sessao.token, p_loja: lojaAtual, p_setor: setorAberto.nome, p_codigo: c, p_valor: alvo });
      noBanco[c] = alvo;
      tudoConfirmado = false;
      ultimaGravacao = horaBR();
      if (falhou === c) falhou = null;
      $(`input[data-c="${c}"]`)?.classList.remove("erro");
    } catch (err) {
      falhou = c;
      $(`input[data-c="${c}"]`)?.classList.add("erro");
      toast("Não consegui salvar o último item: " + err.message + " Confira a conexão; ele será salvo de novo ao clicar em Salvar pedido.", "erro");
    } finally {
      gravando--; atualizarSalvo(); atualizarResumo();
    }
  });
  return filas[c];
}
// espera tudo o que está na fila de gravação
async function descarregar() {
  for (const c of Object.keys(timers)) { clearTimeout(timers[c]); delete timers[c]; gravarItem(Number(c)); }
  await Promise.allSettled(Object.values(filas));
}
function atualizarSalvo() {
  const el = $("#salvo"); if (!el) return;
  if (gravando > 0 || Object.keys(timers).length) { el.className = "salvo gravando"; el.innerHTML = "<i></i>Salvando..."; }
  else if (falhou) { el.className = "salvo falhou"; el.innerHTML = "<i></i>Falhou ao salvar"; }
  else if (ultimaGravacao) { el.className = "salvo"; el.innerHTML = `<i></i>Salvo às ${ultimaGravacao}`; }
  else el.innerHTML = "";
}
function marcarLinha(c) {
  const l = $(`.produto[data-linha="${c}"]`); if (!l) return;
  const p = dados.produtos.find((x) => x.c === c);
  const v = valores[c];
  const q = qtdInt(v);
  const fora = R.media && p && p.m > 0 && q > CFG.FATOR_ALERTA_LOJA * p.m;
  l.classList.toggle("com-qtd", temValor(v, R.texto) && !fora);
  l.classList.toggle("fora-curva", !!fora);
  const av = $(".aviso-curva", l);
  av.hidden = !fora;
  if (fora) av.textContent = `Pedido ${fmtNum(q / p.m, 0)}x acima da média. Confira se não é erro de digitação.`;
}
function itensPreenchidos() { return dados.produtos.filter((p) => temValor(valores[p.c], R.texto)); }
function atualizarResumo() {
  if (telaAtual !== "setor" || !dados) return;
  const itens = itensPreenchidos();
  const fora = R.media ? itens.filter((p) => p.m > 0 && qtdInt(valores[p.c]) > CFG.FATOR_ALERTA_LOJA * p.m) : [];
  const fc = $("#faixa-curva");
  if (fc) {
    fc.hidden = !fora.length;
    if (fora.length) fc.innerHTML = `<div><b>${fora.length} ${fora.length === 1 ? "item" : "itens"} com pedido acima de ${CFG.FATOR_ALERTA_LOJA}x a média.</b> Confira se não há erro de digitação:
      <table><thead><tr><th>Cód. ERP</th><th>Descrição</th><th>Média</th><th>Pedido</th></tr></thead><tbody>
      ${fora.map((p) => `<tr><td class="num">${esc(p.erp)}</td><td>${esc(p.d)}</td><td class="num">${fmtNum(p.m, 2)}</td><td class="num">${qtdInt(valores[p.c])}</td></tr>`).join("")}</tbody></table></div>`;
  }
  const fr = $("#faixa-rascunho"); if (fr) fr.hidden = tudoConfirmado;
  const r = $("#resumo-barra");
  if (r) {
    const obsMudou = R.obs && $("#obs-loja") && $("#obs-loja").value.trim() !== obsOriginal.trim();
    r.innerHTML = `<strong class="num">${itens.length} ${itens.length === 1 ? "produto preenchido" : "produtos preenchidos"}</strong>
      <small>${fora.length ? `<span class="tag alerta">${fora.length} acima de ${CFG.FATOR_ALERTA_LOJA}x a média</span> ` : ""}${tudoConfirmado && !obsMudou ? (itens.length ? "Pedido salvo." : "Nada lançado ainda.") : "Rascunho salvo. Falta clicar em Salvar pedido."}</small>`;
  }
}
function renderBarra() {
  const b = $("#barra-total");
  b.hidden = false; document.body.classList.add("com-barra");
  b.innerHTML = `<div class="resumo" id="resumo-barra"></div>
    ${R.semPedido ? '<button class="btn btn-secundario" type="button" id="btn-sem-pedido">Sem pedido hoje</button>' : ""}
    <button class="btn btn-secundario" type="button" id="btn-excel">Exportar Excel</button>
    <button class="btn btn-secundario" type="button" id="btn-imprimir">Imprimir</button>
    <button class="btn btn-primario" type="button" id="btn-salvar">Salvar pedido</button>`;
  $("#btn-salvar").onclick = salvarPedido;
  $("#btn-excel").onclick = exportarExcel;
  $("#btn-imprimir").onclick = imprimir;
  $("#btn-sem-pedido")?.addEventListener("click", confirmarSemPedido);
  for (const p of dados.produtos) if (valores[p.c]) marcarLinha(p.c);
  atualizarResumo();
}

// ---------- Salvar pedido ----------
async function salvarPedido() {
  const btn = $("#btn-salvar");
  btn.disabled = true; btn.textContent = "Gravando...";
  try {
    await descarregar();
    const itens = dados.produtos.map((p) => ({ c: p.c, v: normalizado(valores[p.c]) }));
    const obs = R.obs ? $("#obs-loja").value : null;
    const r = await rpc("pedidos_web_salvar_pedido", { p_token: sessao.token, p_loja: lojaAtual, p_setor: setorAberto.nome, p_itens: itens, p_obs: obs });
    for (const it of itens) noBanco[it.c] = it.v;
    falhou = null; $$(".qtd input.erro").forEach((i) => i.classList.remove("erro"));
    tudoConfirmado = true;
    if (R.obs) obsOriginal = obs || "";
    ultimaGravacao = r.hora || horaBR();
    if (dados.sem_pedido && r.itens > 0) { dados.sem_pedido = false; $(".faixa.ok")?.remove(); }
    atualizarSalvo(); atualizarResumo();
    painel = null;   // painel recarrega na próxima visita
    abrirJanela(`<h3>Pedido salvo com sucesso</h3>
      <p>O pedido de <b>${esc(setorAberto.titulo)}</b> da <b>${nomeLoja(lojaAtual)}</b> foi salvo às ${esc(ultimaGravacao)} (${r.itens} ${r.itens === 1 ? "item" : "itens"}).</p>
      <p>Pode conferir, imprimir ou exportar quando quiser.</p>
      <div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Fechar</button><button class="btn btn-primario" type="button" data-fechar data-ir="painel">Voltar ao painel</button></div>`, "ok");
  } catch (err) {
    toast(err.message, "erro");
  } finally {
    btn.disabled = false; btn.textContent = "Salvar pedido";
  }
}

// ---------- Sem pedido hoje (Açougues) ----------
function confirmarSemPedido() {
  abrirJanela(`<h3>Sem pedido hoje</h3>
    <p>Confirma que a <b>${nomeLoja(lojaAtual)}</b> <b>não</b> fará pedido de <b>${esc(setorAberto.titulo)}</b> hoje?</p>
    <p>Isso apaga o que estiver lançado e avisa o supervisor no Telegram.</p>
    <div class="acoes"><button class="btn btn-secundario" type="button" data-fechar>Não, cancelar</button><button class="btn btn-primario" type="button" id="btn-sem-sim">Sim, sem pedido hoje</button></div>`);
  $("#btn-sem-sim").onclick = async (e) => {
    const b = e.currentTarget; b.disabled = true; b.textContent = "Registrando...";
    try {
      for (const c of Object.keys(timers)) { clearTimeout(timers[c]); delete timers[c]; }
      await Promise.allSettled(Object.values(filas));
      await rpc("pedidos_web_sem_pedido", { p_token: sessao.token, p_loja: lojaAtual, p_setor: setorAberto.nome });
      let aviso = true;
      try { await avisos("sem_pedido", { setor: setorAberto.nome, loja: lojaAtual }); } catch (err) { aviso = err.message; }
      fecharJanela();
      falhou = null; painel = null;
      await abrirSetor(setorAberto);
      if (aviso === true) toast("Supervisor avisado e pedido zerado. Nenhum pedido será feito hoje neste setor.", "ok");
      else toast("O pedido foi zerado, mas o aviso no Telegram falhou: " + aviso, "erro");
    } catch (err) {
      b.disabled = false; b.textContent = "Sim, sem pedido hoje";
      toast(err.message, "erro");
    }
  };
}

// ---------- janela ----------
function abrirJanela(html, tipo = "") {
  const j = $("#janela");
  j.className = "janela " + tipo; j.innerHTML = html; j.hidden = false; $("#veu").hidden = false;
  setTimeout(() => $("button", j)?.focus(), 30);
}
function fecharJanela() { $("#janela").hidden = true; $("#veu").hidden = true; }
$("#veu").addEventListener("click", fecharJanela);
$("#janela").addEventListener("click", (e) => {
  const b = e.target.closest("[data-fechar]");
  if (!b) return;
  fecharJanela();
  if (b.dataset.ir === "painel") irPara("#/painel");
});
document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#janela").hidden) fecharJanela(); });

// ============================================================
// Impressão e Excel (o que está na tela, com o filtro aplicado — igual ao Streamlit)
// ============================================================
function linhasExport() {
  const vis = new Set(linhasVisiveis().map((l) => Number(l.dataset.linha)));
  return dados.produtos.filter((p) => vis.has(p.c));
}
function colunasExport() {
  // mesmas colunas da tela, sem a Foto
  return colunasSetor().filter(([k]) => k !== "foto").map(([k]) => ({
    k,
    tela: { cod: "Cód. ERP", barra: "Cód. Barra", forn: "Fornecedor", desc: "Descrição", media: "Média (90d)", est: "Estoque ERP", qtd: "Pedido" }[k],
    excel: { cod: "Código", barra: "Cód. Barra", forn: "Fornecedor", desc: "Descrição", media: "Média (90d)", est: "Estoque", qtd: "Pedido" }[k],
  }));
}
function valorCelula(p, k) {
  switch (k) {
    case "cod": return p.erp;
    case "barra": return p.b;
    case "forn": return p.f;
    case "desc": return p.d;
    case "media": return Number(p.m || 0);
    case "est": return Math.trunc(Number(p.e || 0));
    case "qtd": return normalizado(valores[p.c]);
  }
  return "";
}
function imprimir() {
  const cols = colunasExport();
  const largura = { cod: 7, barra: 13, forn: 13, media: 9, est: 9, qtd: 11 };
  const usada = cols.reduce((s, c) => s + (largura[c.k] || 0), 0);
  const linhas = linhasExport();
  const obs = R.obs ? $("#obs-loja").value.trim() : "";
  $("#impressao").className = "so-impressao";
  $("#impressao").innerHTML = `<h3>📦 Pedido Oficial — ${nomeLoja(lojaAtual)}</h3>
    <p class="emitido">${esc(setorAberto.titulo)} · Emitido em ${dataHoraBR()}</p>
    <table><colgroup>${cols.map((c) => `<col style="width:${c.k === "desc" ? 100 - usada : largura[c.k]}%">`).join("")}</colgroup>
    <thead><tr>${cols.map((c) => `<th>${c.tela}</th>`).join("")}</tr></thead>
    <tbody>${linhas.map((p) => `<tr>${cols.map((c) => {
      let v = valorCelula(p, c.k);
      if (c.k === "media") v = fmtNum(v, 2);
      return `<td class="${["forn", "desc"].includes(c.k) ? "esq" : ""}">${esc(v)}</td>`;
    }).join("")}</tr>`).join("")}</tbody></table>
    ${obs ? `<div class="obs"><strong>📝 Observação Geral da Loja:</strong>\n${esc(obs)}</div>` : ""}`;
  // numeração de página (Embalagens/Padaria saem em várias folhas)
  let st = $("#estilo-paginas");
  if (!st) { st = document.createElement("style"); st.id = "estilo-paginas"; document.head.appendChild(st); }
  st.textContent = R.obs ? '@media print { @page { margin: 10mm 10mm 14mm 10mm; @bottom-right { content: "Página " counter(page) " de " counter(pages); font-size: 7.5pt; color: #333; } } }' : "";
  window.print();
}

let exceljsPromessa = null;
function carregarExcelJS() {
  if (window.ExcelJS) return Promise.resolve();
  if (!exceljsPromessa) exceljsPromessa = new Promise((ok, erro) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    s.onload = ok; s.onerror = () => { exceljsPromessa = null; erro(new Error("Não consegui carregar o gerador de Excel. Confira a internet.")); };
    document.head.appendChild(s);
  });
  return exceljsPromessa;
}
// mesma regra do Streamlit para Pedido/Estoque/Média: vazio, "-" ou 0 ficam em branco
function limparExcel(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (["", "-", "0", "0.0"].includes(s)) return null;
  const f = Number(s.replace(",", "."));
  if (Number.isFinite(f) && s !== "") return f === 0 ? null : f;
  return s;
}
async function exportarExcel() {
  const btn = $("#btn-excel");
  btn.disabled = true; btn.textContent = "Gerando...";
  try {
    await carregarExcelJS();
    const cols = colunasExport();
    const linhas = linhasExport();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Pedido", {
      pageSetup: { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        margins: { left: 1 / 2.54, right: 1 / 2.54, top: 1 / 2.54, bottom: 1 / 2.54, header: .5 / 2.54, footer: .5 / 2.54 } },
    });
    const d = agoraSP();
    ws.getCell("A1").value = `Pedidos do dia ${p2(d.getDate())}/${p2(d.getMonth() + 1)}/${d.getFullYear()}`;
    ws.getCell("A1").font = { bold: true };
    const borda = { style: "thin", color: { argb: "FF000000" } };
    const bordaCab = { style: "dashed", color: { argb: "FFFFFFFF" } };
    const verdes = new Set(["qtd", "media", "est"]);
    const cab = ws.getRow(2);
    cols.forEach((c, i) => {
      const cel = cab.getCell(i + 1);
      cel.value = c.excel;
      cel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC55A11" } };
      cel.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cel.border = { top: bordaCab, left: bordaCab, bottom: bordaCab, right: bordaCab };
      cel.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      const w = { forn: 30, desc: 45, barra: 16, cod: 12, media: 12, est: 12 }[c.k] || 9;
      ws.getColumn(i + 1).width = w;
    });
    linhas.forEach((p, r) => {
      const row = ws.getRow(3 + r);
      cols.forEach((c, i) => {
        let v = valorCelula(p, c.k);
        if (verdes.has(c.k)) v = limparExcel(v);
        else if (c.k === "cod") v = Number(v);
        const cel = row.getCell(i + 1);
        cel.value = v === "" ? null : v;
        cel.border = { top: borda, left: borda, bottom: borda, right: borda };
        cel.font = { bold: true };
        cel.alignment = { horizontal: ["forn", "desc"].includes(c.k) ? "left" : "center", vertical: "middle", wrapText: true };
        if (verdes.has(c.k)) cel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } };
      });
    });
    const obs = R.obs ? $("#obs-loja").value.trim() : "";
    if (obs) {
      const n = Math.max(1, cols.length);
      const rl = 3 + linhas.length + 1;
      ws.getCell(rl, 1).value = "📝 Observação Geral da Loja:";
      ws.getCell(rl, 1).font = { bold: true };
      if (n > 1) ws.mergeCells(rl, 1, rl, n);
      const rt = rl + 1;
      ws.getCell(rt, 1).value = obs;
      ws.getCell(rt, 1).alignment = { horizontal: "left", vertical: "top", wrapText: true };
      ws.getCell(rt, 1).border = { top: borda, left: borda, bottom: borda, right: borda };
      if (n > 1) ws.mergeCells(rt, 1, rt + 3, n);
      ws.getRow(rt).height = 60;
    }
    const buf = await wb.xlsx.writeBuffer();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = `Pedido_${nomeLoja(lojaAtual)}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  } catch (err) {
    toast(err.message, "erro");
  } finally {
    btn.disabled = false; btn.textContent = "Exportar Excel";
  }
}

// ============================================================
// Tema claro/escuro
// ============================================================
function aplicarTema(t) {
  document.documentElement.dataset.theme = t;
  $("#txt-tema").textContent = t === "dark" ? "Claro" : "Escuro";
  try { localStorage.setItem("pedidos_web_tema", t); } catch {}
}
$("#btn-tema").addEventListener("click", () => aplicarTema(document.documentElement.dataset.theme === "dark" ? "light" : "dark"));
$("#txt-tema").textContent = document.documentElement.dataset.theme === "dark" ? "Claro" : "Escuro";

// ---------- menu da esquerda: mostrar/esconder ----------
$("#btn-menu").addEventListener("click", () => {
  const fechado = document.documentElement.classList.toggle("menu-fechado");
  try { localStorage.setItem("pedidos_web_menu", fechado ? "fechado" : "aberto"); } catch {}
});

// ---------- início ----------
// espera o central.js carregar (as telas da central ficam nele)
document.addEventListener("DOMContentLoaded", () => {
  if (sessao?.token) iniciarApp();
  else { $("#tela-login").hidden = false; montarLogin(); }
});
