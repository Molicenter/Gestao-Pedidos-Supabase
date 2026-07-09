import streamlit as st
import base64
import os

# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURAÇÃO DA PÁGINA (Única para todo o portal)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão Pedidos - Molicenter",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────────────────────────────────────
# 2. IMPORTAÇÃO DOS MÓDULOS (Enxuta: apenas a tela genérica unificada)
# ─────────────────────────────────────────────────────────────────────────────
import tela_pedidos_generica # <-- Apenas este arquivo cuidará de todos os setores!

# ─────────────────────────────────────────────────────────────────────────────
# 3. LINKS E VARIÁVEIS DE IMAGENS EXTERNAS
# ─────────────────────────────────────────────────────────────────────────────
IMG_FOLHAGEM = "https://images.unsplash.com/photo-1574316071802-0d684efa7bf5?w=400"
IMG_FLV      = "https://images.unsplash.com/photo-1610348725531-843dff563e2c?w=400"
IMG_ORIENTAL = "https://images.unsplash.com/photo-1579871494447-9811cf80d66c?w=400"
IMG_ACOUGUE  = "https://images.unsplash.com/photo-1544025162-d76694265947?w=400"
IMG_PADARIA  = "https://images.unsplash.com/photo-1509440159596-0249088772ff?w=400"

@st.cache_data
def imagem_para_b64(caminho):
    if not caminho: return ""
    if caminho.startswith("http"): return caminho
    try:
        with open(caminho, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            ext = caminho.split('.')[-1].lower()
            mime = "image/png" if ext == "png" else "image/jpeg"
            return f"data:{mime};base64,{encoded}"
    except:
        return ""

# ─────────────────────────────────────────────────────────────────────────────
# 4. VARIÁVEIS DE SESSÃO GLOBAIS
# ─────────────────────────────────────────────────────────────────────────────
if 'usuario_logado' not in st.session_state:
    st.session_state['usuario_logado'] = None

if 'modulo_ativo' not in st.session_state:
    st.session_state['modulo_ativo'] = 'Home'

# ─────────────────────────────────────────────────────────────────────────────
# 5. CSS AVANÇADO (TEMA CLARO MOLICENTER + FORÇAR MENU ESQUERDO)
# 🎨 O fundo e as cores base agora vêm do .streamlit/config.toml (tema claro).
#    Este CSS cuida apenas dos componentes customizados (banner, cards, login).
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Controle de Largura e Zoom (Cravado para mostrar tudo sem rolar) ── */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 1rem !important;
    max-width: 95% !important;
    zoom: 0.90 !important; /* Ajuste fino: compacto para os cards, mas legível no login */
}

/* Ocultar Menu e Rodapé, MAS DEIXAR O CABEÇALHO TRANSPARENTE */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {background-color: transparent !important;} 

/* ── MÁGICA: FORÇAR O MENU ESQUERDO A APARECER PARA AS LOJAS ── */
/* Isso anula as travas escondidas dentro dos outros 9 arquivos dos setores */
html body [data-testid="collapsedControl"] {
    display: flex !important;
    visibility: visible !important;
    z-index: 99999 !important;
}
html body section[data-testid="stSidebar"] {
    display: block !important;
}

/* ── Banner Superior Azul Escuro (identidade Molicenter — mantido) ── */
.banner-container {
    background: linear-gradient(135deg, #07263b 0%, #0e4a74 100%);
    padding: 12px 24px;
    border-radius: 8px;
    margin-bottom: 25px;
    display: flex;
    align-items: center;
    gap: 14px;
    box-shadow: 0 4px 15px rgba(11, 61, 99, 0.20);
    border: 1px solid rgba(11, 61, 99, 0.25);
}
.banner-logo {
    height: 40px;
    width: auto;
    object-fit: contain;
}
.banner-title {
    font-family: 'Segoe UI', Tahoma, sans-serif;
    font-size: 22px;
    font-weight: 800;
    color: #fff;
}

/* ── Container dos Cards (branco com borda azulada, hover azul Molicenter) ── */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #ffffff !important;
    border-radius: 8px !important;
    border: 1px solid #D5E0EA !important;
    transition: all 0.25s ease !important;
}
div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    border-color: #0093E9 !important;
    background-color: #F7FAFD !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] > div {
    padding: 12px !important;
    gap: 6px !important;
}

/* ── Imagem do Card ── */
.card-img-container {
    width: 100%;
    height: 120px; 
    border-radius: 4px;
    overflow: hidden;
    margin-bottom: 8px;
    background-color: #F2F6FA;
}
.card-img-container img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    opacity: 0.95;
    transition: opacity 0.3s;
}
div[data-testid="stVerticalBlockBorderWrapper"]:hover .card-img-container img {
    opacity: 1.0;
}

/* ── Botões de Título dos Cards (azul Molicenter sobre branco) ── */
div[data-testid="stVerticalBlockBorderWrapper"] button:not([kind="primary"]) {
    background-color: #ffffff !important;
    background: #ffffff !important;
    border: 1px solid #C4D4E0 !important;
    border-radius: 4px !important;
    width: 100% !important;
    min-height: 38px !important;
    padding: 6px 12px !important;
    margin-top: 5px !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] button:not([kind="primary"]) * {
    color: #0B3D63 !important; /* TEXTO AZUL MOLICENTER */
    font-weight: 800 !important;
    font-size: 14px !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] button:not([kind="primary"]):hover {
    background-color: #EAF3FB !important;
    background: #EAF3FB !important;
    border-color: #0093E9 !important;
}

/* ── Formatação de Texto de Horários ── */
.texto-horario {
    font-size: 12px;
    color: #44576B;
    line-height: 1.4;
    font-weight: 500;
    min-height: 36px;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    margin-top: 8px;
}

/* ── Títulos das Linhas (Setores) — azul com detalhe magenta ── */
.linha-titulo-sec {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #0B4F7A;
    margin-bottom: 8px;
    margin-top: 15px;
    font-weight: 700;
    border-left: 3px solid #E5007D;
    padding-left: 8px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 6. TELA DE LOGIN ÚNICA DO PORTAL (MAIOR E QUADRADA)
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state['usuario_logado'] is None:
    st.write(<br><br><br>"", unsafe_allow_html=True)"
    
    # Aumentando a largura da coluna central para 1.8 (deixa a caixa mais imponente)
    _, col2, _ = st.columns([1, 1, 1])
    
    with col2:
        with st.container(border=True):
            # Espaço extra interno para deixar "quadrado"
            st.write("<br>", unsafe_allow_html=True)
            
            # --- INÍCIO DA MUDANÇA: Sub-colunas para alinhar Título e Logo ---
            # Proporção 1:4:1 para manter o texto centralizado na caixa
            _, title_col, logo_col = st.columns([1, 4, 1])
            
            with title_col:
                st.markdown("""
                    <div style='text-align:center;'>
                        <h2 style='margin-bottom:0; color:#0B3D63;'>Portal de Pedidos</h2>
                        <p style='color:#5F7387;font-size:14px;'>Acesso Unificado — Molicenter</p>
                    </div>
                """, unsafe_allow_html=True)
                
            with logo_col:
                # Opcional: Adicionar um pequeno espaço para alinhar verticalmente com o texto
                st.write("") 
                try:
                    st.image("passaro_logo.png", width=60)
                except Exception:
                    # Centralizando o emoji para caso a imagem falhe
                    st.markdown("<div style='text-align:center; font-size:30px;'>🥬</div>", unsafe_allow_html=True)
            # --- FIM DA MUDANÇA ---
        
            st.divider()
            
            LOJAS_LOGIN = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]
            usuarios_permitidos = ["Selecione...", "Administrador"] + LOJAS_LOGIN
            usuario_selecionado = st.selectbox("👤 Usuário de acesso:", usuarios_permitidos)
            senha_digitada = st.text_input("🔑 Senha de acesso:", type="password", autocomplete="off")

            # ─────────────────────────────────────────────────────────────────
            # 🔧 CORREÇÃO DO AUTOFILL DA SENHA
            # Quando o navegador preenche a senha salva, ele NÃO dispara o evento
            # que o Streamlit escuta — então o Python recebia a senha vazia e o
            # botão "Entrar" não fazia nada (só funcionava ao redigitar).
            # Este script força esse evento uma única vez por campo, fazendo o
            # Streamlit registrar a senha preenchida automaticamente.
            # ─────────────────────────────────────────────────────────────────
            st.components.v1.html(
                """
                <script>
                (function(){
                  const win = window.parent, doc = win.document;
                  const t = setInterval(function(){
                    doc.querySelectorAll('input[type="password"]').forEach(function(inp){
                      if (inp.value && !inp.dataset.__synced) {
                        inp.dataset.__synced = '1';
                        const setter = Object.getOwnPropertyDescriptor(
                            win.HTMLInputElement.prototype, 'value').set;
                        setter.call(inp, inp.value);                 // valor preenchido pelo navegador
                        inp.dispatchEvent(new Event('input', { bubbles: true })); // Streamlit registra
                      }
                    });
                  }, 250);
                  setTimeout(function(){ clearInterval(t); }, 5000); // para após ~5s
                })();
                </script>
                """,
                height=0,
            )
            
            st.write("<br>", unsafe_allow_html=True)

            if st.button("Entrar no Sistema", type="primary", use_container_width=True):
                if usuario_selecionado == "Selecione...":
                    st.error("⚠️ Por favor, selecione um usuário.")
                elif usuario_selecionado == "Administrador" and senha_digitada == "moli0000":
                    st.session_state['usuario_logado'] = usuario_selecionado
                    st.rerun()
                elif usuario_selecionado in LOJAS_LOGIN and senha_digitada == "moli1234":
                    st.session_state['usuario_logado'] = usuario_selecionado
                    st.rerun()
                elif senha_digitada:
                    st.error("⚠️ Senha incorreta. Tente novamente.")
                else:
                    # Rede de segurança: senha veio vazia (ex.: autofill não captado)
                    st.warning("⚠️ Digite a senha. Se ela foi preenchida automaticamente pelo navegador, clique no campo e tecle uma vez antes de entrar.")
            
            st.write("<br>", unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# 7. FUNÇÕES PARA GERAR A VITRINE DE CARDS
# ─────────────────────────────────────────────────────────────────────────────
def criar_card(titulo, subtitulo, caminho_imagem, emoji_fallback, chave_modulo):
    img_src = imagem_para_b64(caminho_imagem)
    
    with st.container(border=True):
        if img_src:
            st.markdown(f"""
            <div class="card-img-container">
                <img src="{img_src}">
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="card-img-container" style="display:flex; justify-content:center; align-items:center;">
                <span style="font-size: 40px;">{emoji_fallback}</span>
            </div>
            """, unsafe_allow_html=True)
        
        # Botão branco com texto azul Molicenter
        if st.button(titulo, key=f"btn_{chave_modulo}", use_container_width=True):
            st.session_state['modulo_ativo'] = chave_modulo
            st.rerun()
            
        st.markdown(f'<div class="texto-horario">{subtitulo}</div>', unsafe_allow_html=True)

def renderizar_dashboard():
    logo_src = imagem_para_b64("passaro_logo.png")
    img_tag = f'<img src="{logo_src}" class="banner-logo" alt="Logo">' if logo_src else '<span style="font-size:28px">🛒</span>'
    
    # 1. Pegamos a loja/usuário que está logado
    loja_logada = st.session_state.get('usuario_logado', '')
    
    # 2. Injetamos a variável {loja_logada} no banner
    st.markdown(f"""
    <div class="banner-container">
        {img_tag}
        <div class="banner-title">
            Gestão Pedidos - Molicenter 
            <span style="font-size: 18px; font-weight: 500; color: #a5d8ff;">&nbsp; ---> Visão: {loja_logada}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # --- LINHA 1 ---
    st.markdown('<div class="linha-titulo-sec">🥦 Setor Hortifruti (FLV)</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4, gap="medium")
    with c1: criar_card("Folhagem", "Seg a Sáb até 12:00hrs", IMG_FOLHAGEM, "🥬", "flv_folhagem")
    with c2: criar_card("FLV Normal", "Terças-feira até 17:00hrs<br>Quintas-feira até 14:00hrs", IMG_FLV, "🍎", "flv_normal")
    with c3: criar_card("FLV Ofertas", "Quintas-feiras até 14:00hrs", IMG_FLV, "🏷️", "flv_ofertas")
    with c4: criar_card("FLV Oriental", "Quintas-feiras até 14:00hrs", IMG_ORIENTAL, "🍣", "flv_oriental")

    # --- LINHA 2 ---
    st.markdown('<div class="linha-titulo-sec">🥩 Setor Açougue e Aves</div>', unsafe_allow_html=True)
    c1, c2, c3, _ = st.columns(4, gap="medium")
    with c1: criar_card("Pioneiro + BF + Paraná", "Seg a Sex até 11:00hrs", "Pioneiros.jpg", "🍗", "acougue_especiais")
    with c2: criar_card("Açougue Adriano", "Quartas-feira até 15:00hrs<br>Sábado até 15:00hrs", IMG_ACOUGUE, "🔪", "acougue_total")
    with c3: criar_card("Peças Açougue - Manoel", "Seg/Qua/Sex - Arap. 15:00h<br>Ter/Qui/Sáb - Maringá 15:00h", IMG_ACOUGUE, "🥩", "acougue_pecas")

    # --- LINHA 3 ---
    st.markdown('<div class="linha-titulo-sec">📦 Outros Setores e Logística</div>', unsafe_allow_html=True)
    c1, c2, c3, _ = st.columns(4, gap="medium")
    with c1: criar_card("Embalagens", "Sexta-feira até as 17:30hrs", "Embalagens.jpg", "🥡", "embalagem")
    with c2: criar_card("Padaria e Confeitaria", "Sábado", IMG_PADARIA, "🥖", "padaria_confeitaria")
    with c3: criar_card("Matéria Prima", "Até Sábado", "materiaprima.jpg", "🌾", "materia_prima")

    # ─────────────────────────────────────────────
    # RODAPÉ
    # ─────────────────────────────────────────────
    st.write("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align:center; margin-top:20px; padding-bottom: 20px; color:#5F7387; font-size:14px; font-weight: 500;">
        Molicenter Supermercados © 2026 — Painel Web de Pedidos Centralizados
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 8. ROTEADOR DE TELAS (Refatorado para Supabase)
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state['modulo_ativo'] == 'Home':
    renderizar_dashboard()
else:
    # Barra lateral padrão para navegação de retorno
    with st.sidebar:
        st.write(f"👤 Utilizador: **{st.session_state['usuario_logado']}**")
        st.divider()
        if st.button("⬅️ Voltar ao Painel Central", use_container_width=True, type="primary"):
            st.session_state['modulo_ativo'] = 'Home'
            st.rerun()
        if st.button("🚪 Sair do Sistema", use_container_width=True):
            st.session_state['usuario_logado'] = None
            st.session_state['modulo_ativo'] = 'Home'
            st.rerun()

    # Dicionário que mapeia o clique do botão ao nome exato do Setor no Supabase
    MAPA_SETORES_SUPABASE = {
        'flv_folhagem': "Folhagem",
        'flv_normal': "FLV Normal",
        'flv_ofertas': "FLV Ofertas",
        'flv_oriental': "FLV Oriental",
        'acougue_especiais': "Açougue Especiais",
        'acougue_total': "Açougue Adriano",
        'acougue_pecas': "Peças Açougue - Manoel",
        'embalagem': "Embalagens",
        'padaria_confeitaria': "Padaria e Confeitaria",
        'materia_prima': "Matéria Prima"
    }

    modulo_atual = st.session_state['modulo_ativo']
    
    if modulo_atual in MAPA_SETORES_SUPABASE:
        # Chama a tela única passando o nome do setor por parâmetro
        tela_pedidos_generica.iniciar_tela(setor=MAPA_SETORES_SUPABASE[modulo_atual])
