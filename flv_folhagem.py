import streamlit as st
import pandas as pd
import io
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection

# ─────────────────────────────────────────────
# CONSTANTES E CONEXÕES GLOBAIS (Não executam ações visuais)
# ─────────────────────────────────────────────
LOJAS = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]
MAPA_LOJAS = {l: l for l in LOJAS}

conn = st.connection("gsheets", type=GSheetsConnection)
conn_pg = st.connection("banco_erp", type="sql")

# ─────────────────────────────────────────────
# FUNÇÕES DE BANCO DE DADOS (Apenas definições)
# ─────────────────────────────────────────────
@st.cache_data(ttl=60)
def buscar_estoque_pg(loja_nome, codigos):
    if not codigos:
        return pd.DataFrame(columns=["Código", "Estoque"])
    loja_id = int(loja_nome.split()[-1])
    loja_id_str = f"{loja_id:03d}" 
    cods_str = ", ".join(map(str, set(codigos)))
    query = f"""
        SELECT cade_codigo AS "Código", estoque AS "Estoque"
        FROM python_estoque WHERE cade_codempresa = '{loja_id_str}' AND cade_codigo IN ({cods_str})
    """
    try:
        return conn_pg.query(query)
    except:
        return pd.DataFrame({"Código": codigos, "Estoque": 0})

def buscar_produtos_pg(codigos):
    if not codigos:
        return pd.DataFrame(columns=["Código", "Descrição"])
    cods_str = ", ".join(map(str, set(codigos)))
    query = f"""SELECT DISTINCT cade_codigo AS "Código", cadp_descricao AS "Descrição" FROM python_estoque WHERE cade_codigo IN ({cods_str})"""
    try: 
        return conn_pg.query(query)
    except: 
        return pd.DataFrame()

@st.cache_data(ttl=15)
def carregar_catalogo_folhagem():
    df = conn.read(worksheet="Produtos_Folhagem", ttl=0, usecols=list(range(20)))
    if df.empty: return pd.DataFrame(columns=["Código", "Descrição", "Nome Personalizado", "Fornecedor"] + LOJAS)
    
    novas_colunas = {col: loja for col in df.columns for loja in LOJAS if loja.lower() in str(col).strip().lower()}
    df = df.rename(columns=novas_colunas)
    
    if "Nome Personalizado" not in df.columns: df["Nome Personalizado"] = ""
    df["Nome Personalizado"] = df["Nome Personalizado"].fillna("").astype(str)
    
    for col in LOJAS:
        if col not in df.columns: df[col] = False
        else:
            df[col] = df[col].apply(lambda x: x if isinstance(x, bool) else str(x).strip().upper() in ['TRUE', 'VERDADEIRO', '1', 'V', 'SIM', 'X'])
            
    if "Código" in df.columns:
        df["Código"] = pd.to_numeric(df["Código"], errors='coerce').fillna(0).astype(int)
    return df

@st.cache_data(ttl=15)
def carregar_pedidos():
    df_pedidos = conn.read(worksheet="Folhagem", ttl=0)
    df_cat = carregar_catalogo_folhagem()
    if not df_pedidos.empty:
        novas_cols = {col: loja for col in df_pedidos.columns for loja in LOJAS if loja.lower() in str(col).strip().lower()}
        df_pedidos = df_pedidos.rename(columns=novas_cols)
    if df_pedidos.empty or "Fornecedor" not in df_pedidos.columns:
        df_init = df_cat[["Código", "Descrição", "Fornecedor"]].copy()
        for loja in LOJAS: df_init[loja] = 0
        return df_init
    if "Código" in df_pedidos.columns:
        df_pedidos["Código"] = pd.to_numeric(df_pedidos["Código"], errors='coerce').fillna(0).astype(int)
    for loja in LOJAS:
        if loja in df_pedidos.columns: df_pedidos[loja] = pd.to_numeric(df_pedidos[loja], errors='coerce').fillna(0).astype(int)
    return df_pedidos

@st.cache_data(ttl=30)
def carregar_vendas_90d():
    try: return conn.read(worksheet="Folhagem_90d", ttl=0)
    except: return pd.DataFrame(columns=["loja", "codigo", "media_dia"])

def salvar_pedidos(df_to_save):
    try: conn.update(worksheet="Folhagem", data=df_to_save); st.cache_data.clear(); return True
    except: return False

def salvar_catalogo(df_to_save):
    try: conn.update(worksheet="Produtos_Folhagem", data=df_to_save); st.cache_data.clear(); return True
    except: return False

# ─────────────────────────────────────────────
# FUNCTION MASTER DO MÓDULO (Executa apenas quando chamada)
# ─────────────────────────────────────────────
def iniciar_tela():
    if 'reset_counter_folhagem' not in st.session_state:
        st.session_state['reset_counter_folhagem'] = 0

    usuario_atual = st.session_state['usuario_logado']
    acesso_total  = (usuario_atual == "Administrador")

    # Força a exibição da barra lateral dentro do módulo se for Admin
    if not acesso_total:
        st.markdown("<style>section[data-testid='stSidebar'] { display: none !important; } [data-testid='collapsedControl'] { display: none !important; }</style>", unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # SIDEBAR INTERNA DO MÓDULO
    # ─────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### Módulo: Folhagem")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação Interna:", [
                "Separação e Fechamento", "Visão das Lojas", "Visão Fornecedores (Resumo)", "Catálogo de Produtos"
            ], key="radio_folhagem_nav")
        else:
            perfil_navegacao = "Visão das Lojas"

        st.divider()
        df_ped = carregar_pedidos()
        total_preenchidos = (df_ped[LOJAS] > 0).any(axis=1).sum() if not df_ped.empty else 0
        st.metric("Itens c/ pedido", total_preenchidos)

    # ─────────────────────────────────────────────
    # DIÁLOGO DE CONFIRMAÇÃO PARA ZERAR
    # ─────────────────────────────────────────────
    @st.dialog("🚨 Confirmação Necessária")
    def modal_zerar_pedidos():
        st.markdown("Tem certeza que deseja **zerar todos os pedidos**?")
        if st.button("✔️ Sim, zerar tudo", type="primary", use_container_width=True):
            st.session_state['reset_counter_folhagem'] += 1
            df_main = carregar_pedidos()
            for loja in LOJAS:
                if loja in df_main.columns: df_main[loja] = 0
            if salvar_pedidos(df_main): st.rerun()

    # ─────────────────────────────────────────────
    # ROTA 1 — SEPARAÇÃO E FECHAMENTO
    # ─────────────────────────────────────────────
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown("<div class='topbar-loja'><div class='topbar-title'>📊 Separação e Fechamento — Folhagem</div></div>", unsafe_allow_html=True)
        with st.container(border=True):
            df_base = carregar_pedidos()
            if df_base.empty: st.warning("Base vazia."); st.stop()
            
            df_base["TOTAL GERAL"] = df_base[LOJAS].sum(axis=1)
            col_cfg = {"Fornecedor": st.column_config.TextColumn("Fornecedor", disabled=True), "Código": st.column_config.NumberColumn("Cód", disabled=True), "Descrição": st.column_config.TextColumn("Produto", disabled=True), "TOTAL GERAL": st.column_config.NumberColumn("TOTAL ▶️", disabled=True)}
            for loja in LOJAS: col_cfg[loja] = st.column_config.NumberColumn(loja, format="%d", min_value=0)
            
            # O dataframe editável no ecrã
            df_editado = st.data_editor(df_base[["Fornecedor", "Código", "Descrição"] + LOJAS + ["TOTAL GERAL"]], hide_index=True, use_container_width=True, height=500, column_config=col_cfg, key=f"sep_editor_{st.session_state['reset_counter_folhagem']}")
            
            st.divider()

            # ─── GERAÇÃO DO FICHEIRO EXCEL DA SEPARAÇÃO ───
            buffer_sep = io.BytesIO()
            df_excel_sep = df_editado.copy()
            
            # Limpeza dos ZEROS do Excel
            for col in LOJAS + ["TOTAL GERAL"]:
                if col in df_excel_sep.columns:
                    df_excel_sep[col] = df_excel_sep[col].apply(lambda x: None if pd.isna(x) or x == 0 else x)

            with pd.ExcelWriter(buffer_sep, engine='openpyxl') as writer:
                df_excel_sep.to_excel(writer, index=False, sheet_name="Separacao")
                worksheet_sep = writer.sheets["Separacao"]
                
                # Ajuste automático das larguras das colunas
                for idx, col_name in enumerate(df_excel_sep.columns):
                    max_content_len = df_excel_sep[col_name].fillna("").astype(str).str.len().max()
                    if pd.isna(max_content_len):
                        max_content_len = 0
                    max_len = max(max_content_len, len(str(col_name))) + 2
                    col_letter = chr(65 + idx)
                    worksheet_sep.column_dimensions[col_letter].width = max_len
                    
            excel_data_sep = buffer_sep.getvalue()

            # ─── BOTÕES ───
            c1, c2, c3 = st.columns([4, 3, 3])
            with c1:
                if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
                    if salvar_pedidos(df_editado.drop(columns=["TOTAL GERAL"])): st.success("Salvo!"); st.rerun()
            with c2:
                # O botão de imprimir foi substituído pelo de download do Excel
                st.download_button(
                    label="📊 Baixar Excel",
                    data=excel_data_sep,
                    file_name="separacao_fechamento_folhagem.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_excel_sep"
                )
            with c3:
                if st.button("🚨 Zerar Todos", use_container_width=True): modal_zerar_pedidos()

    # ─────────────────────────────────────────────
    # ROTA 2 — VISÃO DAS LOJAS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS) if acesso_total else usuario_atual
        st.markdown(f"<div class='topbar-loja'><div class='topbar-title'>🥬 {loja_selecionada} — Lançamento</div></div>", unsafe_allow_html=True)
        
        df_cat = carregar_catalogo_folhagem()
        df_cat_loja = df_cat[df_cat[loja_selecionada] == True].copy()
        if df_cat_loja.empty: st.warning("Nenhum produto habilitado."); st.stop()

        df_all = carregar_pedidos()
        df_all_clean = df_all.drop(columns=["Descrição"]) if "Descrição" in df_all.columns else df_all
        df_loja_view = pd.merge(df_cat_loja[["Fornecedor", "Código", "Descrição", "Nome Personalizado"]], df_all_clean[["Código", "Fornecedor", loja_selecionada]], on=["Código", "Fornecedor"], how="left")
        df_loja_view[loja_selecionada] = df_loja_view[loja_selecionada].fillna(0).astype(int)
        
        mask = (df_loja_view["Nome Personalizado"].notna()) & (df_loja_view["Nome Personalizado"].str.strip() != "")
        df_loja_view.loc[mask, "Descrição"] = df_loja_view.loc[mask, "Nome Personalizado"]
        
        df_estoque = buscar_estoque_pg(loja_selecionada, df_loja_view["Código"].unique().tolist())
        df_loja_view = pd.merge(df_loja_view, df_estoque, on="Código", how="left")
        df_loja_view["Estoque"] = df_loja_view["Estoque"].fillna(0).astype(int)
        df_loja_view = df_loja_view.rename(columns={loja_selecionada: "Qtde"})

        col_cfg_l = {"Fornecedor": st.column_config.TextColumn(disabled=True), "Código": st.column_config.NumberColumn(disabled=True), "Descrição": st.column_config.TextColumn(disabled=True), "Estoque": st.column_config.NumberColumn(disabled=True), "Qtde": st.column_config.NumberColumn(min_value=0, step=1)}
        df_editado_loja = st.data_editor(df_loja_view[["Fornecedor", "Código", "Descrição", "Estoque", "Qtde"]], column_config=col_cfg_l, hide_index=True, use_container_width=True, height=450, key=f"loja_folhagem_{st.session_state['reset_counter_folhagem']}")
        
        if st.button("💾 Salvar Pedido da Loja", type="primary", use_container_width=True):
            df_main = carregar_pedidos()
            for _, row in df_editado_loja.iterrows():
                m_row = (df_main["Fornecedor"] == row["Fornecedor"]) & (df_main["Código"] == row["Código"])
                if m_row.any(): 
                    df_main.loc[m_row, loja_selecionada] = row["Qtde"]
                    df_main.loc[m_row, "Descrição"] = row["Descrição"]
            if salvar_pedidos(df_main): st.success("Pedido gravado!"); st.rerun()

    # ─────────────────────────────────────────────
    # ROTA 3 — VISÃO FORNECEDORES
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Visão Fornecedores (Resumo)":
        st.markdown("<div class='topbar-loja'><div class='topbar-title'>🚚 Resumo por Fornecedor</div></div>", unsafe_allow_html=True)
        
        df_all = carregar_pedidos()
        df_cat = carregar_catalogo_folhagem()
        
        # Dicionário de tamanho fixo em pixels para garantir alinhamento perfeito
        col_cfg_resumo = {
            "Código": st.column_config.TextColumn("Código", width=60),
            "Descrição": st.column_config.TextColumn("Descrição", width=280),
            "TOTAL": st.column_config.TextColumn("TOTAL", width=70)
        }
        for l in LOJAS:
            col_cfg_resumo[l] = st.column_config.TextColumn(l, width=65)

        df_excel_list = []
        colunas_exibicao = ["Código", "Descrição"] + LOJAS + ["TOTAL"]

        for forn in df_cat["Fornecedor"].dropna().unique():
            df_forn = df_all[df_all["Fornecedor"] == forn].copy()
            if df_forn.empty: continue
            
            df_cat_forn = df_cat[df_cat["Fornecedor"] == forn]
            codigos_habilitados = df_cat_forn["Código"].unique()
            df_forn = df_forn[df_forn["Código"].isin(codigos_habilitados)].copy()
            
            if df_forn.empty: continue
                
            df_forn["TOTAL"] = df_forn[LOJAS].sum(axis=1)
            
            df_forn_export = df_forn[colunas_exibicao].copy()
            df_forn_export.insert(0, "Fornecedor", forn)
            df_excel_list.append(df_forn_export)
            
            df_display = df_forn[colunas_exibicao].copy()
            df_display["Código"] = df_display["Código"].astype(str)
            
            for col in LOJAS + ["TOTAL"]:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)
                df_display[col] = df_display[col].astype(str).replace({"0": ""})
            
            with st.container(border=True):
                st.markdown(f"##### Fornecedor: {forn}")
                
                st.dataframe(
                    df_display, 
                    hide_index=True, 
                    use_container_width=False, 
                    column_config=col_cfg_resumo,
                    key=f"f_{forn}"
                )

        buffer = io.BytesIO()
        if df_excel_list:
            df_final_excel = pd.concat(df_excel_list, ignore_index=True)
            
            for col in LOJAS + ["TOTAL"]:
                df_final_excel[col] = df_final_excel[col].apply(lambda x: None if pd.isna(x) or x == 0 else x)

            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final_excel.to_excel(writer, index=False, sheet_name="Resumo_Pedidos")
                worksheet = writer.sheets["Resumo_Pedidos"]
                
                for idx, col_name in enumerate(df_final_excel.columns):
                    max_content_len = df_final_excel[col_name].fillna("").astype(str).str.len().max()
                    if pd.isna(max_content_len):
                        max_content_len = 0
                    max_len = max(max_content_len, len(str(col_name))) + 2
                    col_letter = chr(65 + idx)
                    worksheet.column_dimensions[col_letter].width = max_len

            excel_data = buffer.getvalue()
        else:
            excel_data = None

        st.markdown("<br>", unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns([3, 4, 3])
        with c2:
            if excel_data:
                st.download_button(
                    label="📊 Baixar Excel Consolidado",
                    data=excel_data,
                    file_name="resumo_pedidos_folhagem.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_excel_forn"
                )
            else:
                st.button("📊 Baixar Excel Consolidado", disabled=True, use_container_width=True)

    # ─────────────────────────────────────────────
    # ROTA 4 — CATÁLOGO DE PRODUTOS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Catálogo de Produtos":
        st.markdown("<div class='topbar-loja'><div class='topbar-title'>🗂️ Catálogo de Produtos</div></div>", unsafe_allow_html=True)
        df_catalogo = carregar_catalogo_folhagem()
        col_cfg_c = {"Código": st.column_config.NumberColumn(), "Descrição": st.column_config.TextColumn(disabled=True)}
        for l in LOJAS: col_cfg_c[l] = st.column_config.CheckboxColumn()
        
        edited_cat = st.data_editor(df_catalogo[["Fornecedor", "Código", "Descrição", "Nome Personalizado"] + LOJAS], use_container_width=True, hide_index=True, height=400, num_rows="dynamic", column_config=col_cfg_c, key="cat_editor")
        if st.button("💾 Salvar Estrutura do Catálogo", type="primary"):
            if salvar_catalogo(edited_cat): st.success("Catálogo atualizado!"); st.rerun()
