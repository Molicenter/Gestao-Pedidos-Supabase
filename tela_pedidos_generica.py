import streamlit as st
import pandas as pd
import io
import time
from datetime import date
from supabase import create_client, Client

# ─────────────────────────────────────────────────────────────────────────────
# ⚙️ CONSTANTES E CONEXÕES GLOBAIS
# ─────────────────────────────────────────────────────────────────────────────
LOJAS_NOMES = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]

def obter_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

conn_pg = st.connection("banco_erp", type="sql")

# ─────────────────────────────────────────────────────────────────────────────
# 🔍 FUNÇÕES AUXILIARES DE CONSULTA E UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def buscar_estoque_erp(loja_nome, codigos):
    if not codigos: return pd.DataFrame(columns=["Código", "Estoque"])
    loja_id = int(loja_nome.split()[-1])
    loja_id_str = f"{loja_id:03d}" 
    cods_str = ", ".join(map(str, set(codigos)))
    query = f"""
        SELECT cade_codigo AS "Código", estoque AS "Estoque"
        FROM python_estoque WHERE cade_codempresa = '{loja_id_str}' AND cade_codigo IN ({cods_str})
    """
    try: return conn_pg.query(query)
    except: return pd.DataFrame({"Código": codigos, "Estoque": 0})

def gerar_excel_download(df: pd.DataFrame, nome_aba: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba[:30])
    return output.getvalue()

def injetar_botao_impressao():
    st.components.v1.html(
        """
        <button onclick="window.print()" style="
            width: 100%;
            background-color: #f0f2f6;
            color: #31333f;
            border: 1px solid rgba(49, 51, 63, 0.2);
            padding: 0.53rem 1rem;
            border-radius: 0.5rem;
            cursor: pointer;
            font-weight: 500;
            font-size: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            box-sizing: border-box;
            height: 38px;
        ">
            🖨️ Imprimir
        </button>
        """,
        height=42,
    )

def exibir_status_digitacao_lojas(df_pedidos_hoje):
    st.markdown("##### 🏪 Status de Digitação das Lojas (Hoje)")
    lojas_que_digitam = set()
    if not df_pedidos_hoje.empty:
        lojas_codigos = df_pedidos_hoje["loja"].unique()
        for c in lojas_codigos:
            lojas_que_digitam.add(f"Loja {c:02d}")
    
    cols = st.columns(8)
    for i, loja_nome in enumerate(LOJAS_NOMES):
        with cols[i]:
            if loja_nome in lojas_que_digitam:
                st.markdown(f"<div style='text-align:center; background-color:#d4edda; color:#155724; padding:5px; border-radius:5px; font-size:11px; font-weight:bold;'>{loja_nome}<br>✅ OK</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='text-align:center; background-color:#f8d7da; color:#721c24; padding:5px; border-radius:5px; font-size:11px; font-weight:bold;'>{loja_nome}<br>❌ Faltando</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 🧠 FUNÇÃO DIRETORA DO MÓDULO UNIFICADO
# ─────────────────────────────────────────────────────────────────────────────
def iniciar_tela(setor: str):
    supabase = obter_supabase()
    usuario_atual = st.session_state.get('usuario_logado', 'Loja 01')
    acesso_total = (usuario_atual == "Administrador")

    with st.sidebar:
        st.markdown(f"### Parâmetros: {setor}")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação Interna:", [
                "Separação e Fechamento", "Visão das Lojas", "Visão Fornecedores (Resumo)", "Catálogo de Produtos"
            ])
        else:
            perfil_navegacao = "Visão das Lojas"

    if acesso_total and perfil_navegacao in ["Separação e Fechamento", "Visão Fornecedores (Resumo)"]:
        with st.sidebar:
            st.markdown("---")
            st.markdown("⚠️ **Zona de Perigo (Admin)**")
            if st.button("🗑️ Limpar Pedidos de Hoje", type="secondary", use_container_width=True):
                with st.spinner("Limpando dados de hoje..."):
                    supabase.table("pedidos").delete().eq("setor", setor).eq("data_pedido", str(date.today())).execute()
                st.toast(f"Planilha de {setor} zerada com sucesso!", icon="🗑️")
                time.sleep(1)
                st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 1 — SEPARAÇÃO E FECHAMENTO (ADMIN)
    # ─────────────────────────────────────────────────────────────────────────
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown(f"## 📊 Separação e Fechamento — {setor}")
        
        resp_prod = supabase.table("produtos").select("codigo, descricao, fornecedor").eq("setor", setor).execute()
        resp_ped = supabase.table("pedidos").select("codigo_produto, loja, quantidade").eq("setor", setor).eq("data_pedido", str(date.today())).execute()
        
        df_prod = pd.DataFrame(resp_prod.data)
        df_ped = pd.DataFrame(resp_ped.data)
        
        if df_prod.empty:
            st.warning("Nenhum produto cadastrado para este setor.")
            return

        exibir_status_digitacao_lojas(df_ped)
        
        if not df_ped.empty:
            df_pivot = df_ped.pivot_table(index='codigo_produto', columns='loja', values='quantidade', aggfunc='sum').reset_index()
            for n in range(1, 9):
                if n in df_pivot.columns: df_pivot = df_pivot.rename(columns={n: f"Loja {n:02d}"})
            itens_com_pedido = df_pivot.shape[0]
        else:
            df_pivot = pd.DataFrame(columns=['codigo_produto'])
            itens_com_pedido = 0

        st.metric(label="📦 Total de Itens Solicitados Hoje", value=f"{itens_com_pedido} produtos")

        df_consolidado = pd.merge(df_prod, df_pivot, left_on='codigo', right_on='codigo_produto', how='left').drop(columns=['codigo_produto'])
        
        for loja in LOJAS_NOMES:
            df_consolidado[loja] = df_consolidado[loja].fillna(0).astype(int)

        df_consolidado["TOTAL GERAL"] = df_consolidado[LOJAS_NOMES].sum(axis=1)
        
        for loja in LOJAS_NOMES:
            df_consolidado[loja] = df_consolidado[loja].replace({0: ""})
        df_consolidado["TOTAL GERAL"] = df_consolidado["TOTAL GERAL"].replace({0: ""})

        df_consolidado = df_consolidado.rename(columns={'codigo': 'Código', 'descricao': 'Descrição', 'fornecedor': 'Fornecedor'})
        df_exibicao = df_consolidado[["Fornecedor", "Código", "Descrição"] + LOJAS_NOMES + ["TOTAL GERAL"]].sort_values(by='Descrição')

        col_cfg = {
            "Fornecedor": st.column_config.TextColumn(disabled=True, width=100), 
            "Código": st.column_config.NumberColumn(disabled=True, width=70, format="%d"), 
            "Descrição": st.column_config.TextColumn(disabled=True, width=180), 
            "TOTAL GERAL": st.column_config.TextColumn("TOTAL", disabled=True, width=70)
        }
        for loja in LOJAS_NOMES: 
            col_cfg[loja] = st.column_config.TextColumn(loja, width=65)
        
        df_editado = st.data_editor(df_exibicao, hide_index=True, use_container_width=True, height=500, column_config=col_cfg)
        
        c_salvar, c_excel, c_print = st.columns([2, 2, 1])
        with c_salvar:
            btn_salvar = st.button("💾 Salvar Ajustes Administrativos", type="primary", use_container_width=True)
        with c_excel:
            dados_excel = gerar_excel_download(df_editado, f"Fechamento {setor}")
            st.download_button(
                label="📊 Exportar Excel (.xlsx)",
                data=dados_excel,
                file_name=f"Separacao_Fechamento_{setor}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_print:
            injetar_botao_impressao()
            
        if btn_salvar:
            with st.spinner("Atualizando registros de pedidos..."):
                for loja_nome in LOJAS_NOMES:
                    n_loja = int(loja_nome.split()[-1])
                    supabase.table("pedidos").delete().eq("setor", setor).eq("loja", n_loja).eq("data_pedido", str(date.today())).execute()
                    
                    lista_insert = []
                    for _, r in df_editado.iterrows():
                        val_loja = r[loja_nome]
                        
                        # Tratamento blindado: evita erros de conversão com pd.NA e nulos
                        if pd.notna(val_loja) and not isinstance(val_loja, (type(pd.NA), type(None))):
                            val_str = str(val_loja).strip()
                            if val_str != "" and val_str != "0" and val_str != "0.0" and val_str.lower() != "<na>":
                                try:
                                    qtd_int = int(float(val_str.replace(',', '.')))
                                    if qtd_int > 0:
                                        lista_insert.append({
                                            "data_pedido": str(date.today()), "setor": setor, "loja": n_loja,
                                            "codigo_produto": int(r["Código"]), "quantidade": qtd_int, "usuario": usuario_atual
                                        })
                                except ValueError:
                                    pass
                                
                    if lista_insert: supabase.table("pedidos").insert(lista_insert).execute()
            st.success("Alterações consolidadas com sucesso!"); st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 2 — VISÃO DAS LOJAS (DIGITAÇÃO DE PEDIDOS)
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS_NOMES) if acesso_total else usuario_atual
        num_loja = int(loja_selecionada.split()[-1])

        st.markdown(f"## 🥬 Lançamento de Pedidos — {loja_selecionada}")
        
        resp_prod = supabase.table("produtos").select("*").eq("setor", setor).eq("ativo", True).execute()
        resp_perm = supabase.table("produtos_lojas").select("codigo_produto, disponivel").eq("loja", num_loja).eq("disponivel", True).execute()
        resp_med = supabase.table("medias_90d").select("codigo_produto, media_dia").eq("loja", num_loja).execute()
        resp_existente = supabase.table("pedidos").select("codigo_produto, quantidade, observacao").eq("setor", setor).eq("loja", num_loja).eq("data_pedido", str(date.today())).execute()

        df_prod = pd.DataFrame(resp_prod.data)
        df_perm = pd.DataFrame(resp_perm.data)
        df_med = pd.DataFrame(resp_med.data)
        df_existente = pd.DataFrame(resp_existente.data)

        if df_prod.empty or df_perm.empty:
            st.warning("Nenhum produto liberado para esta loja neste setor.")
            return

        df_loja = pd.merge(df_prod, df_perm, left_on='codigo', right_on='codigo_produto', how='inner')
        df_loja = pd.merge(df_loja, df_med, on='codigo_produto', how='left')
        df_loja['media_dia'] = df_loja['media_dia'].fillna(0.0)

        if not df_existente.empty:
            df_loja = pd.merge(df_loja, df_existente, on='codigo_produto', how='left')
            df_loja['quantidade'] = df_loja['quantidade'].fillna(0).astype(int)
            df_loja['observacao'] = df_loja['observacao'].fillna("")
            itens_digitados = df_existente[df_existente['quantidade'] > 0].shape[0]
        else:
            df_loja['quantidade'] = 0
            df_loja['observacao'] = ""
            itens_digitados = 0

        st.metric(label="📝 Seus Itens Preenchidos", value=f"{itens_digitados} produtos")

        df_estoque = buscar_estoque_erp(loja_selecionada, df_loja["codigo"].tolist())
        df_loja = pd.merge(df_loja, df_estoque, left_on='codigo', right_on='Código', how='left')
        df_loja["Estoque"] = df_loja["Estoque"].fillna(0).astype(int)

        df_loja['quantidade'] = df_loja['quantidade'].replace({0: ""})

        df_final_grid = pd.DataFrame({
            'Código': df_loja['codigo'], 'Fornecedor': df_loja['fornecedor'], 'Descrição': df_loja['descricao'],
            'Média (90d)': df_loja['media_dia'], 'Estoque ERP': df_loja['Estoque'],
            'Qtde Pedida': df_loja['quantidade'], 'Observação': df_loja['observacao']
        }).sort_values(by='Descrição')

        col_cfg_l = {
            "Fornecedor": st.column_config.TextColumn(disabled=True, width=120), 
            "Código": st.column_config.NumberColumn(disabled=True, format="%d", width=70), 
            "Descrição": st.column_config.TextColumn(disabled=True, width=250),
            "Estoque ERP": st.column_config.NumberColumn(disabled=True, format="%d", width=90), 
            "Média (90d)": st.column_config.NumberColumn(disabled=True, format="%.2f", width=90),
            "Qtde Pedida": st.column_config.TextColumn("Qtde Pedida", width=100), 
            "Observação": st.column_config.TextColumn("Observação", max_chars=100, width=180)
        }

        grid_editado = st.data_editor(df_final_grid, column_config=col_cfg_l, hide_index=True, use_container_width=True)

        c_salvar, c_excel, c_print = st.columns([2, 2, 1])
        with c_salvar:
            btn_salvar_loja = st.button("💾 Salvar Pedido Oficial", type="primary", use_container_width=True)
        with c_excel:
            dados_excel = gerar_excel_download(grid_editado, f"Pedido {loja_selecionada}")
            st.download_button(
                label="📊 Exportar Excel (.xlsx)",
                data=dados_excel,
                file_name=f"Pedido_{loja_selecionada}_{setor}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_print:
            injetar_botao_impressao()

        if btn_salvar_loja:
            with st.spinner("Gravando no banco de dados relacional..."):
                supabase.table("pedidos").delete().eq("setor", setor).eq("loja", num_loja).eq("data_pedido", str(date.today())).execute()
                
                lista_inserts = []
                for _, r in grid_editado.iterrows():
                    val_ped = r["Qtde Pedida"]
                    if pd.notna(val_ped) and not isinstance(val_ped, (type(pd.NA), type(None))):
                        val_str = str(val_ped).strip()
                        if val_str != "" and val_str != "0" and val_str != "0.0" and val_str.lower() != "<na>":
                            try:
                                qtd_int = int(float(val_str.replace(',', '.')))
                                if qtd_int > 0:
                                    lista_inserts.append({
                                        "data_pedido": str(date.today()), "setor": setor, "loja": num_loja, "codigo_produto": int(r["Código"]),
                                        "quantidade": qtd_int, "observacao": str(r["Observação"]).strip() if r["Observação"] else None, "usuario": usuario_atual
                                    })
                            except ValueError:
                                pass
                if lista_inserts:
                    supabase.table("pedidos").insert(lista_inserts).execute()
            st.success("Pedido gravado instantaneamente no Supabase!"); st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 3 — VISÃO FORNECEDORES (RESUMO)
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Visão Fornecedores (Resumo)":
        st.markdown(f"## 🚚 Resumo Consolidado por Fornecedor — {setor}")
        
        resp_prod = supabase.table("produtos").select("codigo, descricao, fornecedor").eq("setor", setor).execute()
        resp_ped = supabase.table("pedidos").select("codigo_produto, loja, quantidade").eq("setor", setor).eq("data_pedido", str(date.today())).execute()
        
        df_prod = pd.DataFrame(resp_prod.data)
        df_ped = pd.DataFrame(resp_ped.data)

        if df_ped.empty:
            st.info("Nenhum pedido realizado hoje para este setor até o momento.")
            return

        df_pivot = df_ped.pivot_table(index='codigo_produto', columns='loja', values='quantidade', aggfunc='sum').reset_index()
        for n in range(1, 9):
            if n in df_pivot.columns: df_pivot = df_pivot.rename(columns={n: f"Loja {n:02d}"})
            
        df_mestre = pd.merge(df_prod, df_pivot, left_on='codigo', right_on='codigo_produto', how='inner').drop(columns=['codigo_produto'])
        
        for l in LOJAS_NOMES: 
            if l not in df_mestre.columns: df_mestre[l] = ""
            else: df_mestre[l] = df_mestre[l].fillna(0).apply(lambda x: int(x) if x == int(x) else x).astype(str).replace({"0": "", "0.0": "", "nan": ""})

        for forn in df_mestre["fornecedor"].dropna().unique():
            df_forn_view = df_mestre[df_mestre["fornecedor"] == forn][["codigo", "descricao"] + [col for col in LOJAS_NOMES if col in df_mestre.columns]].rename(columns={'codigo': 'Código', 'descricao': 'Produto'})
            with st.container(border=True):
                st.markdown(f"##### Fornecedor: {forn}")
                st.dataframe(df_forn_view, hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c_excel, c_print = st.columns([4, 1])
        with c_excel:
            dados_excel = gerar_excel_download(df_mestre, f"Fornecedores {setor}")
            st.download_button(
                label="📊 Exportar Fornecedores para Excel (.xlsx)",
                data=dados_excel,
                file_name=f"Resumo_Fornecedores_{setor}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_print:
            injetar_botao_impressao()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 4 — CATÁLOGO DE PRODUTOS
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Catálogo de Produtos":
        st.markdown(f"## 🗂️ Gestão de Catálogo e Permissões por Loja — {setor}")

        resp_prod = supabase.table("produtos").select("*").eq("setor", setor).execute()
        resp_perm = supabase.table("produtos_lojas").select("*").execute()

        df_prod = pd.DataFrame(resp_prod.data)
        df_perm = pd.DataFrame(resp_perm.data)

        if df_prod.empty:
            st.warning("Nenhum produto cadastrado no mestre para este setor.")
            return

        if not df_perm.empty:
            df_perm_pivot = df_perm.pivot(index='codigo_produto', columns='loja', values='disponivel').reset_index()
            for n in range(1, 9):
                df_perm_pivot = df_perm_pivot.rename(columns={n: f"Loja {n:02d}"})
        else:
            df_perm_pivot = pd.DataFrame(columns=['codigo_produto'] + LOJAS_NOMES)

        df_cat_completo = pd.merge(df_prod, df_perm_pivot, left_on='codigo', right_on='codigo_produto', how='left').drop(columns=['codigo_produto'])
        for l in LOJAS_NOMES: df_cat_completo[l] = df_cat_completo[l].fillna(True).astype(bool)

        col_cfg_c = {
            "codigo": st.column_config.NumberColumn("Código", disabled=True, format="%d", width=70), 
            "descricao": st.column_config.TextColumn("Descrição do Produto", disabled=True, width=220), 
            "fornecedor": st.column_config.TextColumn("Fornecedor/Marca", width=120)
        }
        for l in LOJAS_NOMES: 
            col_cfg_c[l] = st.column_config.CheckboxColumn(l, width=65)

        edited_cat = st.data_editor(df_cat_completo[["fornecedor", "codigo", "descricao"] + LOJAS_NOMES], use_container_width=True, hide_index=True, column_config=col_cfg_c)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Salvar Matriz do Catálogo", type="primary", use_container_width=True):
            with st.spinner("Atualizando regras de disponibilidade por loja..."):
                lista_upserts_permissoes = []
                for _, row in edited_cat.iterrows():
                    cod_p = int(row["codigo"])
                    for num_loja in range(1, 9):
                        col_loja = f"Loja {num_loja:02d}"
                        lista_upserts_permissoes.append({
                            "codigo_produto": cod_p,
                            "loja": num_loja,
                            "disponivel": bool(row[col_loja])
                        })
                supabase.table("produtos_lojas").upsert(lista_upserts_permissoes, on_conflict="codigo_produto, loja").execute()
                st.success("Matriz de travas updated!"); st.rerun()
