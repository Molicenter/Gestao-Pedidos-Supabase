import streamlit as st
import pandas as pd
import io
from streamlit_gsheets import GSheetsConnection

# ─────────────────────────────────────────────
# CONSTANTES E PRODUTOS INICIAIS
# ─────────────────────────────────────────────
LOJAS = ["Loja 01", "Loja 02", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]

produtos_iniciais = [
    {"Código": 521798, "Descrição": "Ampan Azuki C/6 Satsumaya",                     "Código Barra": "7897327901294", "Marca": "Satsumaya"},
    {"Código": 520504, "Descrição": "Kamaboko 200g Agronippo",                       "Código Barra": "7896293804165", "Marca": "Agronippo"},
    {"Código": 521150, "Descrição": "Kamaboko 200g Kai-Ho Agronippo",                "Código Barra": "7896293804110", "Marca": "Agronippo"},
    {"Código": 521974, "Descrição": "Lokyozuke Cebolinha 200g Marunaka Cons",        "Código Barra": "7897231100134", "Marca": "Agronippo"},
    {"Código": 533623, "Descrição": "Massa Gobomaki Kai-Ho 200g",                    "Código Barra": "7896293804134", "Marca": "Agronippo"},
    {"Código": 53662,  "Descrição": "Massa Konnyaku 350g C/Alga",                    "Código Barra": "7896293805100", "Marca": "Agronippo"},
    {"Código": 524768, "Descrição": "Massa Shirataki 200g Agronippo Noodles Konjac", "Código Barra": "7898944092266", "Marca": "Agronippo"},
    {"Código": 583497, "Descrição": "Massa Shirataki 200g Hyde Alimentos Noodles",   "Código Barra": "7898944092211", "Marca": "Hyde Alimentos"},
    {"Código": 577362, "Descrição": "Mirinzuke 200g Conserva De Nabo",               "Código Barra": "7896101500265", "Marca": "Agronippo"},
    {"Código": 520911, "Descrição": "Narutomakii Kai.Ho 200g",                       "Código Barra": "7896293804158", "Marca": "Agronippo"},
    {"Código": 141820, "Descrição": "Nippo Kyoka Natto 100g",                        "Código Barra": "7896293805001", "Marca": "Agronippo"},
    {"Código": 524713, "Descrição": "Nippo Shirataki 200g",                          "Código Barra": "7896293805117", "Marca": "Agronippo"},
    {"Código": 139940, "Descrição": "Shogazuke 245g Beni Shoga Gengibre Ralado",    "Código Barra": "7897231100042", "Marca": "Agronippo"},
    {"Código": 55406,  "Descrição": "Sushi Ague 110g Agronippo",                     "Código Barra": "7896293803014", "Marca": "Agronippo"},
    {"Código": 608981, "Descrição": "Takuan 200g Haruko Amarelo",                    "Código Barra": "0798190064482", "Marca": "Haruko"},
    {"Código": 521965, "Descrição": "Takuan 500g Haruko Amarelo",                    "Código Barra": "0798190022024", "Marca": "Haruko"},
    {"Código": 100221, "Descrição": "Takuwan 200g Takaki Pequeno",                   "Código Barra": "7896101500272", "Marca": "Ceasa Box"},
    {"Código": 176507, "Descrição": "Takuwan 500g Takaki Grande",                    "Código Barra": "7896101500234", "Marca": "Ceasa Box"},
    {"Código": 530112, "Descrição": "Tempura 200g Kai Ho",                           "Código Barra": "7896293804127", "Marca": "Agronippo"},
    {"Código": 53679,  "Descrição": "Tikuwa 200g Agronippo",                         "Código Barra": "7896293804103", "Marca": "Agronippo"},
    {"Código": 577380, "Descrição": "Tiocenzuke 245 conserva Pic Nabo Acelga",       "Código Barra": "7897231100219", "Marca": "Agronippo"},
    {"Código": 524722, "Descrição": "Tofu 1kg Agronippo Nigari Momen",               "Código Barra": "7896293808156", "Marca": "Agronippo"},
    {"Código": 524731, "Descrição": "Tofu 500g Agronippo Tradicional",               "Código Barra": "7896293802048", "Marca": "Agronippo"},
]

# ─────────────────────────────────────────────
# CONEXÃO GOOGLE SHEETS & FUNÇÕES DE DADOS
# ─────────────────────────────────────────────
conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=15)
def carregar_banco_oriental():
    df_prod = conn.read(worksheet="Produtos_Oriental")
    df_ped = conn.read(worksheet="Pedidos_Oriental")

    mudou_algo = False

    # Inicializa Produtos se estiver vazio
    if df_prod.empty or "Código" not in df_prod.columns:
        df_prod = pd.DataFrame(produtos_iniciais)
        for loja in LOJAS: df_prod[loja] = True
        conn.update(worksheet="Produtos_Oriental", data=df_prod)
        mudou_algo = True

    # Inicializa Pedidos se estiver vazio
    if df_ped.empty or "Código" not in df_ped.columns:
        df_ped = pd.DataFrame(columns=["Código"] + LOJAS)
        df_ped["Código"] = df_prod["Código"]
        df_ped[LOJAS] = 0
        conn.update(worksheet="Pedidos_Oriental", data=df_ped)
        mudou_algo = True

    # --- Sincronização Automática (Se adicionarem produtos no catálogo) ---
    novos_ped = df_prod[~df_prod["Código"].isin(df_ped["Código"])]["Código"]
    if not novos_ped.empty:
        df_n_ped = pd.DataFrame({"Código": novos_ped})
        df_n_ped[LOJAS] = 0
        df_ped = pd.concat([df_ped, df_n_ped], ignore_index=True)
        conn.update(worksheet="Pedidos_Oriental", data=df_ped)
        mudou_algo = True

    # --- BLINDAGEM DE TIPOS (Evita o StreamlitAPIException) ---
    if "Código" in df_prod.columns:
        df_prod["Código"] = pd.to_numeric(df_prod["Código"], errors='coerce').fillna(0).astype(int)
    if "Descrição" in df_prod.columns:
        df_prod["Descrição"] = df_prod["Descrição"].fillna("").astype(str)
    if "Código Barra" in df_prod.columns:
        df_prod["Código Barra"] = df_prod["Código Barra"].astype(str).str.replace(r'\.0$', '', regex=True).replace("nan", "")
    if "Marca" in df_prod.columns:
        df_prod["Marca"] = df_prod["Marca"].fillna("").astype(str)

    for loja in LOJAS:
        if loja in df_ped.columns: df_ped[loja] = pd.to_numeric(df_ped[loja], errors='coerce').fillna(0).astype(int)
        if loja in df_prod.columns: df_prod[loja] = df_prod[loja].fillna(False).astype(bool)

    if mudou_algo:
        st.cache_data.clear()

    return df_prod, df_ped

# ─────────────────────────────────────────────────────────────
# MÓDULO PRINCIPAL - A FUNÇÃO QUE O MAIN.PY VAI CHAMAR
# ─────────────────────────────────────────────────────────────
def iniciar_tela():
    
    if 'reset_counter_flv_oriental' not in st.session_state:
        st.session_state['reset_counter_flv_oriental'] = 0

    usuario_atual = st.session_state.get('usuario_logado', None)
    acesso_total  = (usuario_atual == "Administrador")

    df_produtos, df_pedidos = carregar_banco_oriental()

    # CSS Global e de Impressão encapsulado
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;700&display=swap');

    :root {
        --bg-main:        #0d1117;
        --bg-card:        #161b22;
        --bg-sidebar:     #0d1117;
        --green-dark:     #1a3a2a;
        --green-mid:      #1f4d35;
        --green-accent:   #2ea043;
        --green-bright:   #3fb950;
        --green-glow:     rgba(46,160,67,.25);
        --text-primary:   #e6edf3;
        --text-muted:     #7d8590;
        --text-header:    #cae8cb;
        --border:         #21262d;
        --border-active:  #2ea043;
        --row-hover:      rgba(46,160,67,.08);
        --row-selected:   rgba(46,160,67,.18);
        --danger:         #da3633;
        --warning:        #d29922;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--green-mid) 0%, var(--green-accent) 100%) !important;
        color: #fff !important; border: 1px solid var(--green-accent) !important;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 4px 18px var(--green-glow) !important;
        border-color: var(--green-bright) !important;
    }

    .stTextInput input:focus, .stSelectbox > div > div:focus-within {
        border-color: var(--green-accent) !important;
        box-shadow: 0 0 0 3px var(--green-glow) !important;
    }

    [data-testid="stDataEditor"] [data-testid="glideDataEditor"] .gdg-header-cell,
    [data-testid="stDataEditor"] .dvn-stack .gdg-header {
        background-color: var(--green-dark) !important;
        color: var(--text-header) !important;
    }

    [data-testid="stDataEditor"] { border: 1px solid var(--green-mid) !important; }
    [data-testid="stDataEditor"] .gdg-cell.gdg-selected,
    [data-testid="stDataEditor"] .gdg-cell[data-state="focused"],
    [data-testid="stDataEditor"] .gdg-cell[aria-selected="true"] {
        background-color: var(--row-selected) !important;
        outline: 2px solid var(--green-accent) !important;
    }
    [data-testid="stDataEditor"] .gdg-row:hover .gdg-cell { background-color: var(--row-hover) !important; }

    div[data-testid="stVerticalBlockBorderWrapper"]:hover { border-color: var(--green-mid) !important; }
    [data-testid="stMetricValue"] { color: var(--green-bright) !important; }

    .stAlert[data-baseweb="notification"] {
        background-color: rgba(46,160,67,.1) !important;
        border: 1px solid var(--green-mid) !important;
        color: var(--text-header) !important;
    }

    .loja-badge {
        display: inline-block; background: var(--green-dark); border: 1px solid var(--green-mid);
        color: var(--green-bright); font-size: 12px; font-weight: 700; padding: 3px 10px;
        border-radius: 20px; margin-bottom: 6px; font-family: 'IBM Plex Mono', monospace;
    }

    .page-header {
        background: linear-gradient(90deg, var(--green-dark) 0%, #0d2018 100%);
        border: 1px solid var(--green-mid); border-radius: 10px; padding: 14px 20px;
        margin-bottom: 22px; display: flex; align-items: center; gap: 12px;
    }
    .page-header-icon { font-size: 26px; }
    .page-header-title { font-size: 20px; font-weight: 700; color: var(--text-header); font-family: 'IBM Plex Sans', sans-serif; }
    .page-header-sub { font-size: 12px; color: var(--text-muted); margin-top: 2px; }

    .topbar-loja {
        background: linear-gradient(90deg, var(--green-dark) 0%, #0d2018 100%);
        border: 1px solid var(--green-mid); border-radius: 10px; padding: 10px 18px;
        margin-bottom: 18px; display: flex; align-items: center; justify-content: space-between;
    }
    .topbar-left { display: flex; align-items: center; gap: 12px; }
    .topbar-title { font-size: 18px; font-weight: 700; color: var(--text-header); }
    .topbar-sub { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
    </style>
    """, unsafe_allow_html=True)

    if not acesso_total:
        st.markdown("""
        <style>
            section[data-testid="stSidebar"] { display: none !important; }
            [data-testid="collapsedControl"]  { display: none !important; }
            .main .block-container { max-width: 100% !important; padding-left: 2.5rem !important; padding-right: 2.5rem !important; }
        </style>
        """, unsafe_allow_html=True)

    def page_header(icon: str, title: str, subtitle: str = ""):
        sub_html = f'<div class="page-header-sub">{subtitle}</div>' if subtitle else ""
        st.markdown(f"""
        <div class="page-header">
            <span class="page-header-icon">{icon}</span>
            <div>
                <div class="page-header-title">{title}</div>
                {sub_html}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # DIÁLOGO DE CONFIRMAÇÃO PARA ZERAR
    # ─────────────────────────────────────────────
    @st.dialog("🚨 Confirmação Necessária")
    def modal_zerar_pedidos_oriental():
        st.markdown("Tem certeza que deseja **zerar todos os pedidos** de todas as lojas?")
        st.markdown("⚠️ *Esta ação limpará as quantidades direto no Google Sheets e não poderá ser desfeita.*")
        st.write("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Não, cancelar", use_container_width=True):
                st.rerun()
        with c2:
            if st.button("✔️ Sim, zerar tudo", type="primary", use_container_width=True):
                st.session_state['reset_counter_flv_oriental'] += 1
                _, df_ped = carregar_banco_oriental()
                df_ped[LOJAS] = 0
                conn.update(worksheet="Pedidos_Oriental", data=df_ped)
                st.cache_data.clear()
                st.rerun()

    # ─────────────────────────────────────────────
    # SIDEBAR INTERNA
    # ─────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### Módulo: FLV Oriental")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação:", [
                "Separação e Fechamento",
                "Visão das Lojas",
                "Catálogo de Produtos"
            ])
        else:
            perfil_navegacao = "Visão das Lojas"

        st.divider()
        total_preenchidos = (df_pedidos[LOJAS] > 0).any(axis=1).sum()
        st.metric("Itens c/ pedido", total_preenchidos, help="Itens que têm ao menos 1 quantidade preenchida")
        st.divider()

    # ─────────────────────────────────────────────
    # ROTA 1: SEPARAÇÃO E FECHAMENTO
    # ─────────────────────────────────────────────
    if perfil_navegacao == "Separação e Fechamento":
        page_header("📊", "Separação e Fechamento", "Consolidado geral — edite quantidades diretamente na tabela")

        with st.container(border=True):
            df_base  = df_produtos[["Código","Descrição","Código Barra","Marca"]]
            df_final = pd.merge(df_base, df_pedidos, on="Código")
            df_final["TOTAL GERAL"] = df_final[LOJAS].sum(axis=1)

            col_cfg = {
                "Código":       st.column_config.NumberColumn(width=80,  format="%d",   disabled=True),
                "Descrição":    st.column_config.TextColumn(             disabled=True),
                "Código Barra": st.column_config.TextColumn("Cód. Barras", width=130,   disabled=True),
                "Marca":        st.column_config.TextColumn(width=110,   disabled=True),
                "TOTAL GERAL":  st.column_config.NumberColumn("TOTAL ▶", width=90,  format="%d", disabled=True),
            }
            for loja in LOJAS:
                col_cfg[loja] = st.column_config.NumberColumn(loja, format="%d", min_value=0, step=1)

            df_editado_admin = st.data_editor(
                df_final, hide_index=True, use_container_width=True,
                height=580, column_config=col_cfg,
                key=f"admin_editor_{st.session_state['reset_counter_flv_oriental']}"
            )

            st.divider()
            col_salvar, col_csv, col_excel, col_limpa, _ = st.columns([2.5, 1.5, 1.5, 2, 2.5])

            with col_salvar:
                if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
                    _, df_ped_fresco = carregar_banco_oriental()
                    for loja in LOJAS:
                        df_ped_fresco[loja] = df_editado_admin[loja]
                    conn.update(worksheet="Pedidos_Oriental", data=df_ped_fresco)
                    st.cache_data.clear()
                    st.success("✅ Ajustes salvos com sucesso na nuvem!")
                    st.rerun()

            with col_csv:
                csv = df_editado_admin.to_csv(index=False).encode("utf-8")
                st.download_button(label="⬇️ CSV", data=csv, file_name="separacao_semanal_horti.csv", mime="text/csv", use_container_width=True)
                
            with col_excel:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_editado_admin.to_excel(writer, index=False, sheet_name='Pedidos')
                st.download_button(label="⬇️ Excel", data=buffer.getvalue(), file_name="separacao_semanal_horti.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            with col_limpa:
                if st.button("🚨 Zerar Pedidos", use_container_width=True):
                    modal_zerar_pedidos_oriental()

    # ─────────────────────────────────────────────
    # ROTA 2: VISÃO DAS LOJAS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS) if acesso_total else usuario_atual

        col_info, col_logout = st.columns([8, 2])
        with col_info:
            st.markdown(f"""
            <div class="topbar-loja">
                <div class="topbar-left">
                    <span style="font-size:22px">📋</span>
                    <div>
                        <div class="topbar-title">{loja_selecionada} — FLV Oriental</div>
                        <div class="topbar-sub">Preencha as quantidades necessárias e salve o pedido da semana</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_logout:
            st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
            if st.button("🏠 Voltar ao Início", use_container_width=True):
                st.session_state['modulo_ativo'] = 'Home'
                st.rerun()

        df_visiveis = df_produtos[df_produtos[loja_selecionada] == True]
        df_loja = pd.merge(
            df_visiveis[["Código","Descrição", "Código Barra", "Marca"]],
            df_pedidos[["Código", loja_selecionada]],
            on="Código"
        )

        mapa_banco_erp = {"Loja 01": "001", "Loja 02": "002", "Loja 03": "003", "Loja 04": "004", "Loja 05": "005", "Loja 06": "006", "Loja 07": "007", "Loja 08": "008"}
        cod_empresa_banco = mapa_banco_erp.get(loja_selecionada, "001")

        try:
            conn_pg = st.connection("banco_erp", type="sql")
            query_erp = f"""
                SELECT cade_codigo, estoque
                FROM "python_estoque"
                WHERE cade_codempresa::text = '{cod_empresa_banco}'
            """
            df_erp = conn_pg.query(query_erp, ttl=300)

            if not df_erp.empty:
                df_erp["cade_codigo"] = pd.to_numeric(df_erp["cade_codigo"], errors='coerce').fillna(0).astype(int)
                df_erp = df_erp.rename(columns={"cade_codigo": "Código", "estoque": "Estoque"})
                
                df_loja["Código"] = pd.to_numeric(df_loja["Código"], errors='coerce').fillna(0).astype(int)
                df_loja = pd.merge(df_loja, df_erp[["Código", "Estoque"]], on="Código", how="left")
            else:
                df_loja["Estoque"] = 0

        except Exception as e:
            st.error(f"⚠️ Erro ao puxar dados do ERP: {e}")
            df_loja["Estoque"] = 0
            
        df_loja["Estoque"] = df_loja["Estoque"].fillna(0).astype(int)
        df_loja = df_loja[["Código", "Descrição", "Código Barra", "Marca", "Estoque", loja_selecionada]]

        with st.container(border=True):
            st.info("💡 **Dica:** Clique em uma célula da coluna **Qtde** para editar. Linhas com quantidade > 0 ficam destacadas automaticamente.")

            col_cfg_loja = {
                "Código":         st.column_config.NumberColumn(width=85,  format="%d", disabled=True),
                "Descrição":      st.column_config.TextColumn(width=280, disabled=True),
                "Código Barra":   st.column_config.TextColumn("Cód. Barras", width=120, disabled=True),
                "Marca":          st.column_config.TextColumn(width=100, disabled=True),
                "Estoque":        st.column_config.NumberColumn("📦 Estoque", width=90, disabled=True),
                loja_selecionada: st.column_config.NumberColumn("🛒 Qtde", width=100, min_value=0, step=1, help="Digite a quantidade desejada"),
            }

            _, col_tabela, _ = st.columns([1, 6, 1])

            with col_tabela:
                df_editado = st.data_editor(
                    df_loja, column_config=col_cfg_loja,
                    hide_index=True, use_container_width=True, height=520,
                    key=f"loja_editor_{st.session_state['reset_counter_flv_oriental']}"
                )

        itens_com_pedido = int((df_editado[loja_selecionada] > 0).sum())
        total_itens      = len(df_editado)
        total_unidades   = int(df_editado[loja_selecionada].sum())
        pct              = round(itens_com_pedido / total_itens * 100) if total_itens else 0

        st.divider()
        m1, m2, m3, _, col_btn = st.columns([1.5, 1.5, 1.5, 2, 3])

        with m1: st.metric("Itens preenchidos", f"{itens_com_pedido} / {total_itens}")
        with m2: st.metric("Total de unidades", total_unidades)
        with m3: st.metric("Cobertura", f"{pct}%")
        with col_btn:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("💾 Salvar Pedido da Semana", type="primary", use_container_width=True):
                _, df_ped_fresco = carregar_banco_oriental()
                for _, row in df_editado.iterrows():
                    mask_ped = df_ped_fresco["Código"] == row["Código"]
                    df_ped_fresco.loc[mask_ped, loja_selecionada] = row[loja_selecionada]
                
                conn.update(worksheet="Pedidos_Oriental", data=df_ped_fresco)
                st.cache_data.clear()
                st.success(f"✅ Pedido da {loja_selecionada} salvo na nuvem com sucesso!")

    # ─────────────────────────────────────────────
    # ROTA 3: CATÁLOGO DE PRODUTOS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Catálogo de Produtos":
        page_header("🏷️", "Catálogo de Produtos", "Gerencie itens disponíveis e permissões por loja")

        with st.container(border=True):
            st.caption("➕ Adicione produtos na última linha  •  🗑️ Selecione a linha e pressione **Delete** para remover  •  ✅ Checkboxes controlam visibilidade por loja")

            # --- LIMPEZA BLINDADA PARA EVITAR ERRO DO DATA EDITOR ---
            df_cat_edit = df_produtos.copy()
            df_cat_edit["Código"] = pd.to_numeric(df_cat_edit["Código"], errors='coerce').fillna(0).astype(int)
            cols_texto = ["Descrição", "Código Barra", "Marca"]
            for col in cols_texto:
                if col in df_cat_edit.columns:
                    df_cat_edit[col] = df_cat_edit[col].fillna("").astype(str)
            for loja in LOJAS:
                if loja in df_cat_edit.columns:
                    df_cat_edit[loja] = df_cat_edit[loja].fillna(False).astype(bool)
            # --------------------------------------------------------

            config_catalogo = {
                "Código":       st.column_config.NumberColumn("Cód. Interno", width=90,  required=True, min_value=1, format="%d"),
                "Descrição":    st.column_config.TextColumn("Descrição do Item", width=310, required=True),
                "Código Barra": st.column_config.TextColumn("Cód. Barras",   width=130, required=True),
                "Marca":        st.column_config.TextColumn("Fabricante",    width=120, required=True),
            }
            for loja in LOJAS:
                config_catalogo[loja] = st.column_config.CheckboxColumn(loja, default=True, width=70)

            df_cat_editado = st.data_editor(
                df_cat_edit,
                num_rows="dynamic",
                column_config=config_catalogo,
                hide_index=True,
                use_container_width=True,
                height=580,
                key=f"cat_oriental_{st.session_state['reset_counter_flv_oriental']}"
            )

            st.divider()

            col_atualizar, col_nomes, col_info = st.columns([2.5, 2.5, 3])
            
            with col_atualizar:
                if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
                    df_cat_editado["Código"] = pd.to_numeric(df_cat_editado["Código"], errors='coerce').fillna(0).astype(int)
                    conn.update(worksheet="Produtos_Oriental", data=df_cat_editado)
                    st.cache_data.clear()
                    st.success("✅ Catálogo e permissões atualizados para todas as lojas!")
                    st.rerun()
                    
            with col_nomes:
                if st.button("📥 Atualizar Nomes (ERP)", use_container_width=True):
                    with st.spinner("⏳ Buscando descrições atualizadas no banco de dados..."):
                        try:
                            conn_pg = st.connection("banco_erp", type="sql")
                            query = 'SELECT cade_codigo, cadp_descricao FROM "python_estoque"'
                            df_erp_nomes = conn_pg.query(query, ttl=300)
                            
                            if not df_erp_nomes.empty:
                                df_erp_nomes["cade_codigo"] = pd.to_numeric(df_erp_nomes["cade_codigo"], errors='coerce').fillna(0).astype(int)
                                df_erp_nomes = df_erp_nomes.drop_duplicates(subset=["cade_codigo"])
                                
                                df_update = df_cat_editado.copy()
                                df_update = df_update.merge(df_erp_nomes, left_on="Código", right_on="cade_codigo", how="left")
                                
                                mask = df_update["cadp_descricao"].notna() & (df_update["cadp_descricao"] != "")
                                df_update.loc[mask, "Descrição"] = df_update.loc[mask, "cadp_descricao"]
                                df_update = df_update.drop(columns=["cade_codigo", "cadp_descricao"])
                                
                                conn.update(worksheet="Produtos_Oriental", data=df_update)
                                st.cache_data.clear()
                                st.success("✅ Nomes atualizados com sucesso pelo ERP!")
                                st.rerun()
                            else:
                                st.warning("⚠️ Não foram encontrados dados na view do banco.")
                        except Exception as e:
                            st.error(f"⚠️ Erro ao puxar nomes do ERP: {e}")
                            
            with col_info:
                total_prods = len(df_cat_editado)
                st.info(f"📦 **{total_prods}** produtos cadastrados  •  **{len(LOJAS)}** lojas")
