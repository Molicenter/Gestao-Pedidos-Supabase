// ============================================================
// Gestão de Pedidos · tela de Usuários (só o perfil Administrador) — versão 1.3.0
// ============================================================
const ROTULO_PERFIL_PED = { admin: "Administrador · central", ceasa: "Ceasa · central (FLV e Folhagem)", loja: "Loja" };
const rotuloPerfil = (u) => (u.perfil === "admin" && u.setores?.length ? "Responsável de setor" : u.perfil === "admin" ? "Administrador · acesso total" : ROTULO_PERFIL_PED[u.perfil] || u.perfil);
const tituloSetor = (n) => SETORES.find((s) => s.nome === n)?.titulo || n;
const ROTULO_SENHA = { propria: '<span class="tag ok">própria</span>', geral: '<span class="tag manual">senha geral das lojas</span>', sem: '<span class="tag pend">sem senha</span>' };
let usuariosPed = null;

async function abrirUsuarios() {
  if (typeof podeSairDoSetor === "function" && !(await podeSairDoSetor())) return;
  telaAtual = "usuarios"; setorAberto = null;
  if (typeof cSetor !== "undefined") cSetor = null;
  $("#barra-total").hidden = true; document.body.classList.remove("com-barra");
  $("#abas").hidden = true; $("#pil-loja").hidden = true;
  $("#pg-titulo").textContent = "Usuários";
  $("#pg-sub").textContent = "Logins do app web de pedidos: perfil, loja, senha e acesso.";
  renderNav();
  $("#conteudo").innerHTML = '<div class="vazio">Carregando...</div>';
  try { usuariosPed = await rpc("pedidos_web_usuarios_lista", { p_token: sessao.token }); }
  catch (e) { $("#conteudo").innerHTML = `<div class="cartao"><p class="erro">${esc(e.message)}</p></div>`; return; }
  renderUsuarios();
}

function renderUsuarios() {
  const d = usuariosPed;
  const quando = (v) => (v ? new Date(v).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo", day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }).replace(",", "") : "—");
  $("#conteudo").innerHTML = `<div class="pagina">
    <div class="cartao"><h3>Logins<small>${d.usuarios.length} usuários</small></h3>
      <p class="dica" style="margin-bottom:12px">Trocar a senha, o perfil, a loja ou desativar derruba na hora quem estiver logado com aquele usuário. A senha fica guardada só como código (hash), ninguém consegue ver.</p>
      <div class="barra-filtros"><button type="button" class="btn btn-primario" data-us-ped="novo">+ Novo login</button></div>
      <div class="tabela-wrap"><table class="us-ped"><thead><tr><th>Usuário</th><th>Perfil</th><th>Loja / setores</th><th>Situação</th><th>Senha</th><th>Último login</th><th></th></tr></thead><tbody>
      ${d.usuarios.map((u, i) => `<tr class="${u.ativo ? "" : "inativo"}">
        <td><b>${esc(u.usuario)}</b>${u.usuario === sessao.usuario ? ' <span class="tag ok">você</span>' : ""}${u.online ? ' <span class="tag manual" title="Tem sessão aberta (entrou nas últimas 18 horas e não saiu)">sessão aberta</span>' : ""}</td>
        <td>${esc(rotuloPerfil(u))}</td><td>${u.loja ? esc(nomeLoja(u.loja)) : u.setores?.length ? u.setores.map((s) => `<span class="tag sem">${esc(tituloSetor(s))}</span>`).join(" ") : u.perfil === "ceasa" ? "Folhagem, FLV Normal e Ofertas" : "Todos os setores"}</td>
        <td>${u.ativo ? '<span class="tag ok">Ativo</span>' : '<span class="tag fora">Desativado</span>'}</td>
        <td>${ROTULO_SENHA[u.senha] || ""}</td><td>${esc(quando(u.ultimo_login))}</td>
        <td><button type="button" class="btn btn-secundario btn-peq" data-us-ped="${i}">Editar</button></td></tr>`).join("")}
      </tbody></table></div></div>
    <div class="cartao"><h3>Senha geral das lojas<small>${d.senha_geral ? "definida" : "não definida"}</small></h3>
      <p class="dica" style="margin-bottom:12px">Vale para toda loja que <b>não</b> tem senha própria. Para uma loja ter senha só dela, edite o login da loja e digite uma senha.</p>
      <form id="form-geral" class="linha-geral" autocomplete="off">
        <label class="campo"><span>Nova senha geral (mínimo 6)</span><input name="senha" type="text" class="senha" autocomplete="off" autocapitalize="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" data-form-type="other"></label>
        <button type="submit" class="btn btn-primario">Trocar senha geral</button>
      </form></div></div>`;
  $$("[data-us-ped]").forEach((b) => (b.onclick = () => formUsuarioPed(b.dataset.usPed === "novo" ? null : d.usuarios[+b.dataset.usPed])));
  $("#form-geral").onsubmit = async (e) => {
    e.preventDefault();
    const v = e.target.senha.value;
    if (v.length < 6) return toast("A senha precisa ter pelo menos 6 caracteres.", "erro");
    try { await rpc("pedidos_web_senha_geral", { p_token: sessao.token, p_senha: v }); toast("Senha geral das lojas trocada.", "ok"); abrirUsuarios(); }
    catch (err) { toast(err.message, "erro"); }
  };
}

function formUsuarioPed(u) {
  const novo = !u;
  abrirJanela(`<h3>${novo ? "Novo login" : "Editar " + esc(u.usuario)}</h3>
    <form id="form-us-ped" class="form-us" autocomplete="off">
      <label class="campo"><span>Nome do usuário (aparece na lista do login)</span><input name="usuario" maxlength="40" value="${esc(u?.usuario || "")}"${novo ? "" : " readonly"} placeholder="ex.: Loja 05 - Padaria, Compras"></label>
      <label class="campo"><span>Perfil</span><select name="perfil">${["loja", "admin", "ceasa"].map((p) => `<option value="${p}"${(u?.perfil || "loja") === p ? " selected" : ""}>${ROTULO_PERFIL_PED[p]}</option>`).join("")}</select></label>
      <fieldset class="campo setores-us" id="campo-setores-us"><span>Acesso nos setores</span>
        <label class="check-us"><input type="radio" name="acesso" value="total"${u?.setores?.length ? "" : " checked"}> Acesso total (todos os setores + tela de Usuários)</label>
        <label class="check-us"><input type="radio" name="acesso" value="resp"${u?.setores?.length ? " checked" : ""}> Responsável só pelos setores marcados (todas as lojas, catálogo e produtos)</label>
        <div class="setores-grade">${SETORES.map((s) => `<label class="check-us"><input type="checkbox" name="setor" value="${esc(s.nome)}"${u?.setores?.includes(s.nome) ? " checked" : ""}> ${esc(s.titulo)}</label>`).join("")}</div>
      </fieldset>
      <label class="campo" id="campo-loja-us"><span>Loja</span><select name="loja">${LOJAS.map((l, i) => `<option value="${i + 1}"${u?.loja === i + 1 ? " selected" : ""}>${l}</option>`).join("")}</select></label>
      <label class="campo"><span>${novo ? "Senha (mínimo 6)" : "Nova senha (em branco = não muda)"}</span>
        <input name="senha" type="text" class="senha" autocomplete="off" autocapitalize="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" data-form-type="other"></label>
      <label class="check-us" id="campo-geral-us"><input type="checkbox" name="usar_geral"${novo || u?.senha === "geral" ? " checked" : ""}> Usar a senha geral das lojas (sem senha própria)</label>
      <label class="check-us"><input type="checkbox" name="ativo"${u?.ativo === false ? "" : " checked"}> Login ativo (aparece na tela de entrada)</label>
      <p class="dica">Administrador e Ceasa veem a central; o perfil Loja só faz o pedido da loja escolhida. Responsável de setor vê todas as lojas, mas só os setores dele, e não abre esta tela.</p>
      <div class="acoes"><button type="button" class="btn btn-secundario" data-fechar>Cancelar</button><button type="submit" class="btn btn-primario" id="us-ped-salvar">Salvar</button></div>
    </form>`);
  const f = $("#form-us-ped");
  const ajustar = () => {
    const loja = f.perfil.value === "loja";
    $("#campo-loja-us").hidden = !loja; $("#campo-geral-us").hidden = !loja;
    if (!loja) f.usar_geral.checked = false;
    $("#campo-setores-us").hidden = f.perfil.value !== "admin";
    $(".setores-grade", f).classList.toggle("desligado", f.acesso.value !== "resp");
  };
  f.perfil.onchange = ajustar; f.acesso.forEach((r) => (r.onchange = ajustar)); ajustar();
  f.senha.oninput = () => { if (f.senha.value) f.usar_geral.checked = false; };
  f.onsubmit = async (e) => {
    e.preventDefault();
    const btn = $("#us-ped-salvar");
    const setores = f.perfil.value === "admin" && f.acesso.value === "resp" ? $$('input[name="setor"]:checked', f).map((c) => c.value) : null;
    if (setores && !setores.length) return toast("Marque pelo menos um setor (ou escolha acesso total).", "erro");
    btn.disabled = true;
    try {
      await rpc("pedidos_web_usuario_salvar", { p_token: sessao.token, p_usuario: f.usuario.value, p_perfil: f.perfil.value,
        p_loja: f.perfil.value === "loja" ? +f.loja.value : null, p_ativo: f.ativo.checked, p_senha: f.senha.value, p_usar_geral: f.usar_geral.checked, p_setores: setores });
      fecharJanela();
      toast(novo ? "Login criado." : "Login atualizado.", "ok");
      abrirUsuarios();
    } catch (err) { toast(err.message, "erro"); btn.disabled = false; }
  };
}

Object.assign(MENSAGENS, {
  SEM_PERMISSAO: "Só o Administrador de acesso total pode mexer nos usuários.",
  SETOR_NAO_PERMITIDO: "Seu login não tem acesso a este setor.",
  NOME_INVALIDO: "O nome do usuário precisa ter de 3 a 40 letras.",
  NOME_RESERVADO: "Esse nome é reservado. Escolha outro.",
  PERFIL_INVALIDO: "Escolha o perfil.",
  NAO_PODE_SI_MESMO: "Você não pode desativar nem tirar o perfil de Administrador do seu próprio login.",
  SENHA_CURTA: "A senha precisa ter pelo menos 6 caracteres.",
  SENHA_OBRIGATORIA: "Defina uma senha para o login novo (ou marque “usar a senha geral das lojas”).",
});
