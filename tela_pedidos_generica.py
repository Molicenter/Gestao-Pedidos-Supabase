import streamlit as st
import pandas as pd
import io
import time
from datetime import date
from supabase import create_client, Client

# ─────────────────────────────────────────────────────────────────────────────
# ⚙️ CONSTANTES, ESTILOS E CONEXÕES GLOBAIS
# ─────────────────────────────────────────────────────────────────────────────
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

LOJAS_NOMES = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]

def obter_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

conn_pg = st.connection("banco_erp", type="sql")

# ─────────────────────────────────────────────────────────────────────────────
# 🔍 FUNÇÕES AUXILIARES DE CONSULTA E UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────
def converter_para_int_seguro(valor) -> int:
    if pd.isna(valor) or valor is None:
        return 0
    val_str = str(valor).strip().replace(',', '.')
    if val_str == "" or val_str.lower() in ["<na>", "none", "nan", "null", "-"]:
        return 0
    try:
        qtd_float = float(val_str)
        return int(qtd_float)
    except ValueError:
        return 0

def buscar_estoque_erp(loja_nome, codigos_erp, setor):
    if not codigos_erp: 
        return pd.DataFrame(columns=["Código", "Estoque"])
        
    loja_id = int(loja_nome.split()[-1])
    loja_id_str = f"{loja_id:03d}" 
    cods_str = ", ".join(map(str, set(codigos_erp)))
    
    coluna_alvo = "estoqueemb" if setor == "Embalagem" else "estoque"
    
    query = f"""
        SELECT cade_codigo AS "Código", {coluna_alvo} AS "Estoque"
        FROM python_estoque WHERE cade_codempresa = '{loja_id_str}' AND cade_codigo IN ({cods_str})
    """
    try: 
        return conn_pg.query(query, ttl=30) 
    except Exception as e: 
        st.error(f"Erro ao buscar estoque: {e}")
        return pd.DataFrame({"Código": codigos_erp, "Estoque": 0})

# ─────────────────────────────────────────────────────────────────────────────
# 📊 MOTORES DE EXPORTAÇÃO EXCEL CUSTOMIZADOS (OPENPYXL)
# ─────────────────────────────────────────────────────────────────────────────
def gerar_excel_download(df: pd.DataFrame, nome_aba: str) -> bytes:
    df_export = df.copy()
    df_export = df_export.rename(columns={
        "Cód. ERP": "Código",
        "TOTAL GERAL": "Total",
        "Qtde Pedida": "Pedido",
        "Estoque ERP": "Estoque"
    })

    def limpar_valor_excel(v):
        if pd.isna(v): return None
        v_str = str(v).strip()
        if v_str in ["", "-", "0", "0.0"]: return None
        try:
            f = float(v_str)
            if f == 0: return None
            return int(f) if f.is_integer() else f
        except ValueError:
            return v 

    for col in df_export.columns:
        col_upper = str(col).upper()
        if any(x in col_upper for x in ["LOJA", "PEDIDO", "ESTOQUE", "TOTAL", "MÉDIA"]):
            df_export[col] = df_export[col].apply(limpar_valor_excel)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nome_aba[:30], startrow=1)
        worksheet = writer.sheets[nome_aba[:30]]

        fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        fill_green = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        
        font_header = Font(color="FFFFFF", bold=True)
        font_bold = Font(bold=True)
        
        border_thin = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                             top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
                             
        border_header = Border(left=Side(style='dashed', color='FFFFFF'), right=Side(style='dashed', color='FFFFFF'),
                               top=Side(style='dashed', color='FFFFFF'), bottom=Side(style='dashed', color='FFFFFF'))
        
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        hoje_str = date.today().strftime("%d/%m/%Y")
        worksheet["A1"] = f"Pedidos do dia {hoje_str}"
        worksheet["A1"].font = font_bold

        colunas_verdes = []
        col_total_idx = None
        cols_para_soma = []
        contador_loja = 0 

        for col_num, column_title in enumerate(df_export.columns, 1):
            col_name = str(column_title).upper()
            if "LOJA" in col_name:
                cols_para_soma.append(col_num)
                contador_loja += 1
                if contador_loja % 2 != 0:
                    colunas_verdes.append(col_num)
            elif any(x in col_name for x in ["PEDIDO", "TOTAL", "MÉDIA", "ESTOQUE"]):
                colunas_verdes.append(col_num)
            if "TOTAL" in col_name:
                col_total_idx = col_num

        for col_num, cell in enumerate(worksheet[2], 1):
            cell.fill = fill_header
            cell.font = font_header
            cell.border = border_header
            cell.alignment = align_center

        for row_num, row in enumerate(worksheet.iter_rows(min_row=3, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column), 3):
            for cell in row:
                cell.border = border_thin
                cell.font = font_bold
                nome_col_atual = str(df_export.columns[cell.column - 1]).upper()
                if any(x in nome_col_atual for x in ["FORNECEDOR", "DESCRIÇÃO", "PRODUTO"]):
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center
                if cell.column in colunas_verdes:
                    cell.fill = fill_green

        if col_total_idx and cols_para_soma:
            letra_total = get_column_letter(col_total_idx)
            letra_primeira = get_column_letter(min(cols_para_soma))
            letra_ultima = get_column_letter(max(cols_para_soma))
            for row_num in range(3, worksheet.max_row + 1):
                worksheet[f"{letra_total}{row_num}"].value = f'=IF(SUM({letra_primeira}{row_num}:{letra_ultima}{row_num})>0, SUM({letra_primeira}{row_num}:{letra_ultima}{row_num}), "")'

        for col_num, column_title in enumerate(df_export.columns, 1):
            letra = get_column_letter(col_num)
            col_name = str(column_title).upper()
            if "FORNECEDOR" in col_name: 
                worksheet.column_dimensions[letra].width = 18
            elif "DESCRI" in col_name or "PRODUTO" in col_name: 
                worksheet.column_dimensions[letra].width = 45
            elif "CÓDIGO" in col_name: 
                worksheet.column_dimensions[letra].width = 12
            elif "MÉDIA" in col_name or "ESTOQUE" in col_name: 
                worksheet.column_dimensions[letra].width = 12
            else: 
                worksheet.column_dimensions[letra].width = 9 

        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
        worksheet.page_setup.orientation = worksheet.ORIENTATION_PORTRAIT
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0 
        worksheet.page_margins.left = 1 / 2.54
        worksheet.page_margins.right = 1 / 2.54
        worksheet.page_margins.top = 1 / 2.54
        worksheet.page_margins.bottom = 1 / 2.54
        worksheet.page_margins.header = 0.5 / 2.54
        worksheet.page_margins.footer = 0.5 / 2.54

    return output.getvalue()

def gerar_excel_fornecedores(df: pd.DataFrame, nome_aba: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        worksheet = writer.book.create_sheet(nome_aba[:30])
        writer.book.active = worksheet
        
        fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        fill_green = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        font_bold = Font(bold=True)
        font_normal = Font(bold=False)
        border_thin = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                             top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
                             
        border_header = Border(left=Side(style='dashed', color='FFFFFF'), right=Side(style='dashed', color='FFFFFF'),
                               top=Side(style='dashed', color='FFFFFF'), bottom=Side(style='dashed', color='FFFFFF'))
                               
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        def limpar(v):
            if pd.isna(v): return ""
            v_str = str(v).strip()
            if v_str in ["", "-", "0", "0.0", "nan"]: return ""
            try:
                f = float(v_str)
                if f == 0: return ""
                return int(f) if f.is_integer() else f
            except ValueError:
                return v_str

        worksheet.column_dimensions['A'].width = 12  
        worksheet.column_dimensions['B'].width = 45  
        for col_idx in range(3, 12): 
            worksheet.column_dimensions[get_column_letter(col_idx)].width = 9

        hoje_str = date.today().strftime("%d/%m/%Y")
        worksheet["A1"] = "Tipo"
        worksheet["B1"] = "Descrição"
        worksheet["C1"] = f"Pedidos do dia {hoje_str}"
        worksheet.merge_cells("C1:K1") 
        
        for col_idx in range(1, 12):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = fill_header
            cell.font = font_header
            cell.border = border_header
            cell.alignment = align_center

        current_row = 2
        lojas_cols = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]
        
        forns = sorted(df["fornecedor"].dropna().unique())
        for forn in forns:
            df_f = df[df["fornecedor"] == forn]
            
            worksheet.cell(row=current_row, column=1).value = ""
            worksheet.cell(row=current_row, column=2).value = forn
            
            for i, loja in enumerate(lojas_cols): 
                worksheet.cell(row=current_row, column=3+i).value = loja
                
            worksheet.cell(row=current_row, column=11).value = "TOTAL"
            
            for col_idx in range(1, 12):
                cell = worksheet.cell(row=current_row, column=col_idx)
                cell.fill = fill_header
                cell.font = font_header
                cell.border = border_header
                cell.alignment = align_center
            
            current_row += 1
            
            for _, r in df_f.iterrows():
                c1 = worksheet.cell(row=current_row, column=1, value=limpar(r.get("Cód. ERP", "")))
                c1.alignment = align_center
                c1.border = border_thin
                c1.font = font_normal
                
                c2 = worksheet.cell(row=current_row, column=2, value=limpar(r.get("Produto", "")))
                c2.alignment = align_left
                c2.border = border_thin
                c2.font = font_normal
                
                for i, loja in enumerate(lojas_cols):
                    c_loja = worksheet.cell(row=current_row, column=3+i, value=limpar(r.get(loja, "")))
                    c_loja.alignment = align_center
                    c_loja.border = border_thin
                    c_loja.font = font_bold
                    if i % 2 == 0: 
                        c_loja.fill = fill_green
                
                c_total = worksheet.cell(row=current_row, column=11)
                c_total.value = f'=IF(SUM(C{current_row}:J{current_row})>0, SUM(C{current_row}:J{current_row}), "")'
                c_total.alignment = align_center
                c_total.border = border_thin
                c_total.font = font_bold
                c_total.fill = fill_green
                current_row += 1
                
        if 'Sheet' in writer.book.sheetnames: 
            writer.book.remove(writer.book['Sheet'])

        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
        worksheet.page_setup.orientation = worksheet.ORIENTATION_PORTRAIT
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0 
        worksheet.page_margins.left = 1 / 2.54
        worksheet.page_margins.right = 1 / 2.54
        worksheet.page_margins.top = 1 / 2.54
        worksheet.page_margins.bottom = 1 / 2.54
        worksheet.page_margins.header = 0.5 / 2.54
        worksheet.page_margins.footer = 0.5 / 2.54

    return output.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# 🖨️ OUTROS UTILS DA INTERFACE
# ─────────────────────────────────────────────────────────────────────────────
def injetar_botao_impressao():
    st.components.v1.html(
        """
        <button onclick="window.parent.print()" style="
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
    st.markdown("<div class='no-print'>##### 🏪 Status de Digitação das Lojas (Hoje)</div>", unsafe_allow_html=True)
    lojas_que_digitam = set()
    if not df_pedidos_hoje.empty:
        lojas_codigos = df_pedidos_hoje["loja"].unique()
        for c in lojas_codigos: 
            lojas_que_digitam.add(f"Loja {c:02d}")
    
    cols = st.columns(8)
    for i, loja_nome in enumerate(LOJAS_NOMES):
        with cols[i]:
            if loja_nome in lojas_que_digitam:
                st.markdown(f"<div class='no-print' style='text-align:center; background-color:#d4edda; color:#155724; padding:5px; border-radius:5px; font-size:11px; font-weight:bold;'>{loja_nome}<br>✅ OK</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='no-print' style='text-align:center; background-color:#f8d7da; color:#721c24; padding:5px; border-radius:5px; font-size:11px; font-weight:bold;'>{loja_nome}<br>❌ Faltando</div>", unsafe_allow_html=True)
    st.markdown("<div class='no-print'><br></div>", unsafe_allow_html=True)

def buscar_permissoes_setor(supabase_client, codigos_setor, num_loja=None):
    if not codigos_setor: 
        return pd.DataFrame()
    dfs = []
    for i in range(0, len(codigos_setor), 200):
        lote = codigos_setor[i:i+200]
        query = supabase_client.table("produtos_lojas").select("codigo_produto, loja, disponivel").in_("codigo_produto", lote)
        if num_loja is not None: 
            query = query.eq("loja", num_loja)
        resp = query.execute()
        if resp.data: 
            dfs.append(pd.DataFrame(resp.data))
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────────────
# 🧠 FUNÇÃO DIRETORA DO MÓDULO UNIFICADO
# ─────────────────────────────────────────────────────────────────────────────
def iniciar_tela(setor: str):
    supabase = obter_supabase()
    usuario_atual = st.session_state.get('usuario_logado', 'Loja 01')
    acesso_total = (usuario_atual == "Administrador")

    st.markdown("""
        <style>
        div[data-testid="stComponentStack"] { width: 100% !important; }
        div[data-testid="stTable"] td { text-align: center !important; }
        
        [data-testid="stSidebar"] button[data-testid="stBaseButton-primary"],
        [data-testid="stSidebar"] button[kind="primary"] {
            background-color: #ff4b4b !important; 
            color: white !important; 
            border-color: #ff4b4b !important;
        }
        
        @media screen {
            .print-only { display: none !important; } 
        }
        
        /* 🔥 CSS OTMIZADO PARA PREENCHIMENTO 100% NA IMPRESSÃO 🔥 */
        @media print {
            @page {
                margin: 10mm;
            }

            /* Oculta tudo que não é o relatório e zera seu espaço */
            section[data-testid="stSidebar"], 
            div[data-testid="stSidebarNav"],
            header[data-testid="stHeader"], 
            footer, 
            [data-testid="stToolbar"],
            button, .stButton, 
            [data-testid="stMetric"], 
            [data-testid="stRadio"], 
            [data-testid="stSelectbox"], 
            [data-testid="stTextInput"], 
            [data-testid="stAlert"],
            [data-testid="stHorizontalBlock"], /* Esconde as colunas de filtros/botões de cima */
            h1, h2, h5, h6, 
            [data-testid="stDataEditor"], 
            [data-testid="stDataFrame"],
            .no-print {
                display: none !important;
                width: 0 !important;
                height: 0 !important;
                min-width: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            
            /* Remove os preenchimentos/bordas invisíveis dos st.container */
            [data-testid="stVerticalBlockBorderWrapper"] {
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            div[data-testid="stVerticalBlock"], .element-container {
                gap: 0 !important;
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
            }

            /* FORÇA O PREENCHIMENTO DO ESPAÇO DA ESQUERDA PARA 0 */
            /* DEPOIS */
            [data-testid="stMain"], .main,
            html, body, .stApp, 
            [data-testid="stAppViewContainer"],
            [data-testid="stAppViewBlockContainer"],
            .block-container {
                background-color: white !important;
                position: static !important;
                margin: 0 !important;
                padding: 0 !important;
                width: 100% !important;
                max-width: 100% !important;
                overflow: visible !important;
            }

            /* Puxa agressivamente o conteúdo pro topo */
            [data-testid="stMain"] > div:first-child,
            [data-testid="stAppViewBlockContainer"] > div:first-child {
                margin-top: -20px !important;
                padding-top: 0 !important;
            }
            
            .print-only {
                display: block !important;
                width: 100% !important;
            }
            
            .print-only h3 {
                font-size: 12pt !important;
                margin: 0 0 10px 0 !important;
                color: black !important;
            }
            
            /* Título do fornecedor grudado na tabela como um bloco contínuo */
            .print-only h4.supplier-header {
                font-size: 9.5pt !important;
                margin: 0 !important;
                padding: 6px 8px !important;
                background-color: #e0e0e0 !important;
                border: 1px solid black !important;
                border-bottom: none !important;
                color: black !important;
                text-align: left !important;
            }
            
            /* Tabela ultra-comprimida */
            table.print-table {
                width: 100% !important;
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                font-size: 8.5pt !important;
                margin-top: 0 !important;
                margin-bottom: 20px !important; /* 🔥 Aumentado o espaço entre fornecedores para não sobrepor 🔥 */
                page-break-inside: auto;
            }
            
            table.print-table tr {
                page-break-inside: avoid;
                page-break-after: auto;
            }
            
            table.print-table th, table.print-table td {
                border: 1px solid black !important;
                padding: 4px 6px !important; 
                color: black !important;
                line-height: 1.1 !important; 
                text-align: center;
            }
            
            table.print-table th {
                background-color: #f0f2f6 !important;
                font-weight: bold;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            
            table.print-table td:nth-child(2), table.print-table td:nth-child(3), table.print-table td:nth-child(4) {
                text-align: left;
            }
        }
        </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(f"### Parâmetros: {setor}")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação Interna:", [
                "Visão Fornecedores (Resumo)",
                "Separação e Fechamento", 
                "Visão das Lojas", 
                "Catálogo de Produtos"
            ])
        else:
            perfil_navegacao = "Visão das Lojas"

    # 🔥 INJEÇÃO PARA MODO PAISAGEM (LANDSCAPE) APENAS NA SEPARAÇÃO E FECHAMENTO 🔥
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown("""
            <style>
            @media print {
                @page { size: landscape; margin: 10mm; }
            }
            </style>
        """, unsafe_allow_html=True)

    if acesso_total:
        with st.sidebar:
            st.markdown("---")
            st.markdown("🔄 **Atualizar Médias (90d)**")
            
            dict_views = {
                "Média Semanal": "python_90dSEMANA", 
                "Diária": "python_90dDIARIA",
                "Seg-Ter": "python_90dSEGTER", 
                "Qua-Qui": "python_90dQUAQUI", 
                "Sex-Sab-Dom": "python_90dSEXSABDOM"
            }
            view_escolhida = st.selectbox("Selecione o período base:", list(dict_views.keys()), label_visibility="collapsed")
            
            if st.button("📥 Puxar Médias do ERP", type="secondary", use_container_width=True):
                view_sql = dict_views[view_escolhida]
                with st.spinner(f"Sincronizando {view_sql}..."):
                    try:
                        resp_prod = supabase.table("produtos").select("codigo, codigo_erp").eq("setor", setor).execute()
                        df_prod_map = pd.DataFrame(resp_prod.data)
                        
                        if df_prod_map.empty: 
                            st.warning("Nenhum produto cadastrado neste setor.")
                        else:
                            df_prod_map = df_prod_map.rename(columns={'codigo': 'codigo_pk_interna'})
                            if 'codigo_erp' not in df_prod_map.columns: 
                                df_prod_map['codigo_erp'] = df_prod_map['codigo_pk_interna']
                            df_prod_map['codigo_erp'] = df_prod_map['codigo_erp'].fillna(df_prod_map['codigo_pk_interna']).astype(int)
                            
                            codigos_erp_setor = df_prod_map['codigo_erp'].unique().tolist()
                            df_erp = conn_pg.query(f'SELECT * FROM "{view_sql}"', ttl=0)
                            
                            if not df_erp.empty:
                                c_loja, c_cod_erp, c_med = df_erp.columns[0], df_erp.columns[1], df_erp.columns[2]
                                df_erp_setor = df_erp[df_erp[c_cod_erp].isin(codigos_erp_setor)]
                                
                                if df_erp_setor.empty: 
                                    st.info("View vazia para estes produtos.")
                                else:
                                    df_merged = pd.merge(df_erp_setor, df_prod_map, left_on=c_cod_erp, right_on='codigo_erp', how='inner')
                                    codigos_pks = df_merged['codigo_pk_interna'].unique().tolist()
                                    
                                    for i in range(0, len(codigos_pks), 200): 
                                        supabase.table("medias_90d").delete().in_("codigo_produto", codigos_pks[i:i+200]).execute()
                                    
                                    lista_insert = []
                                    for _, row in df_merged.iterrows():
                                        lista_insert.append({
                                            "loja": int(row[c_loja]), 
                                            "codigo_produto": int(row['codigo_pk_interna']), 
                                            "media_dia": float(row[c_med]) if pd.notna(row[c_med]) else 0.0
                                        })
                                        
                                    for i in range(0, len(lista_insert), 1000): 
                                        supabase.table("medias_90d").insert(lista_insert[i:i+1000]).execute()
                                        
                                    st.success(f"Médias ({view_escolhida}) sincronizadas!")
                                    time.sleep(1.5)
                                    st.rerun()
                            else: 
                                st.warning("View do ERP retornou vazia.")
                    except Exception as e: 
                        st.error(f"Erro na sincronização: {e}")

            st.markdown("---")
            if 'confirmar_limpeza' not in st.session_state: 
                st.session_state['confirmar_limpeza'] = False
                
            if not st.session_state['confirmar_limpeza']:
                if st.button("🗑️ Limpar Pedidos", type="primary", use_container_width=True): 
                    st.session_state['confirmar_limpeza'] = True
                    st.rerun()
            else:
                st.warning("⚠️ **Confirma exclusão?**")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✔️ Sim", type="primary", use_container_width=True):
                        with st.spinner("Limpando..."): 
                            supabase.table("pedidos").delete().eq("setor", setor).eq("data_pedido", str(date.today())).execute()
                        st.session_state['confirmar_limpeza'] = False
                        st.rerun()
                with c2:
                    if st.button("❌ Não", use_container_width=True): 
                        st.session_state['confirmar_limpeza'] = False
                        st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 1 — SEPARAÇÃO E FECHAMENTO
    # ─────────────────────────────────────────────────────────────────────────
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown(f"<div class='no-print'><h2>📊 Separação e Fechamento — {setor}</h2></div>", unsafe_allow_html=True)
        
        resp_prod = supabase.table("produtos").select("codigo, codigo_erp, descricao, fornecedor, nome_personalizado").eq("setor", setor).execute()
        resp_ped = supabase.table("pedidos").select("codigo_produto, loja, quantidade").eq("setor", setor).eq("data_pedido", str(date.today())).execute()
        
        df_prod = pd.DataFrame(resp_prod.data)
        df_ped = pd.DataFrame(resp_ped.data)
        
        if df_prod.empty: 
            st.warning("Nenhum produto cadastrado para este setor.")
            return
            
        if 'codigo_erp' not in df_prod.columns: 
            df_prod['codigo_erp'] = df_prod['codigo']
            
        df_prod['codigo_erp'] = df_prod['codigo_erp'].fillna(df_prod['codigo']).astype(int)

        codigos_setor = df_prod['codigo'].tolist()
        df_perm_all = buscar_permissoes_setor(supabase, codigos_setor)
        df_prod['descricao'] = df_prod['nome_personalizado'].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else None).fillna(df_prod['descricao'])

        exibir_status_digitacao_lojas(df_ped)
        
        if not df_ped.empty:
            df_pivot = df_ped.pivot_table(index='codigo_produto', columns='loja', values='quantidade', aggfunc='sum').reset_index()
            for n in range(1, 9):
                if n in df_pivot.columns: 
                    df_pivot = df_pivot.rename(columns={n: f"Loja {n:02d}"})
        else: 
            df_pivot = pd.DataFrame(columns=['codigo_produto'])

        df_consolidado = pd.merge(df_prod, df_pivot, left_on='codigo', right_on='codigo_produto', how='left')
        
        if not df_perm_all.empty:
            df_perm_all['loja_nome'] = df_perm_all['loja'].apply(lambda x: f"Loja {int(x):02d}_perm")
            df_perm_pivot = df_perm_all.pivot_table(index='codigo_produto', columns='loja_nome', values='disponivel', aggfunc='last').reset_index()
        else: 
            df_perm_pivot = pd.DataFrame(columns=['codigo_produto'])

        df_consolidado = pd.merge(df_consolidado, df_perm_pivot, left_on='codigo', right_on='codigo_produto', how='left')
        
        perm_cols = [f"Loja {n:02d}_perm" for n in range(1, 9)]
        for col in perm_cols:
            if col not in df_consolidado.columns: 
                df_consolidado[col] = True
                
        df_consolidado[perm_cols] = df_consolidado[perm_cols].fillna(True)
        mask_active = df_consolidado[perm_cols].any(axis=1)
        df_consolidado = df_consolidado[mask_active]

        for loja in LOJAS_NOMES:
            if loja not in df_consolidado.columns: 
                df_consolidado[loja] = 0.0
            df_consolidado[loja] = df_consolidado[loja].fillna(0).astype(int)

        df_consolidado["TOTAL_NUM"] = df_consolidado[LOJAS_NOMES].sum(axis=1)

        st.markdown("<div class='no-print'><br></div>", unsafe_allow_html=True)
        col_filtro, col_metric = st.columns([3, 1])
        with col_filtro:
            opcoes_forn = ["Todos"] + sorted([str(f) for f in df_consolidado['fornecedor'].unique() if pd.notna(f) and str(f).strip() != ""])
            filtro_selecionado = st.radio("🔍 Filtrar Exibição por Setor:", options=opcoes_forn, horizontal=True)

        if filtro_selecionado != "Todos": 
            df_consolidado = df_consolidado[df_consolidado['fornecedor'] == filtro_selecionado]

        st.metric(label="📦 Itens c/ pedido", value=f"{df_consolidado[df_consolidado['TOTAL_NUM'] > 0].shape[0]} produtos")

        df_consolidado["TOTAL GERAL"] = df_consolidado["TOTAL_NUM"].replace({0: ""})
        for loja in LOJAS_NOMES:
            df_consolidado[loja] = df_consolidado[loja].replace({0: ""}).astype(str).replace({"0": "", "0.0": "", "nan": ""})
            perm_col = f"{loja}_perm"
            if perm_col in df_consolidado.columns: 
                df_consolidado.loc[df_consolidado[perm_col] != True, loja] = "-"

        df_consolidado = df_consolidado.rename(columns={'codigo_erp': 'Cód. ERP', 'descricao': 'Descrição', 'fornecedor': 'Fornecedor'})
        df_exibicao = df_consolidado[["Fornecedor", "codigo", "Cód. ERP", "Descrição"] + LOJAS_NOMES + ["TOTAL GERAL"]].sort_values(by=['Fornecedor', 'Descrição'])

        col_cfg = {
            "codigo": None, 
            "Cód. ERP": st.column_config.NumberColumn("Cód. ERP", disabled=True, format="%d", width=80), 
            "Fornecedor": st.column_config.TextColumn(disabled=True, width=110), 
            "Descrição": st.column_config.TextColumn(disabled=True, width=200), 
            "TOTAL GERAL": st.column_config.TextColumn("TOTAL", disabled=True, width=70)
        }
        for loja in LOJAS_NOMES: 
            col_cfg[loja] = st.column_config.TextColumn(loja, width=75, disabled=False)
        
        df_editado = st.data_editor(df_exibicao, hide_index=True, use_container_width=True, height=500, column_config=col_cfg)
        
        html_table = df_exibicao.drop(columns=['codigo'], errors='ignore').fillna('').to_html(index=False, classes="print-table")
        st.markdown(f'<div class="print-only"><h3>📊 Separação e Fechamento — {setor} ({filtro_selecionado})</h3>{html_table}</div>', unsafe_allow_html=True)

        c_salvar, c_excel, c_print = st.columns([2, 2, 1])
        with c_salvar: 
            btn_salvar = st.button("💾 Salvar Ajustes Administrativos", type="primary", use_container_width=True)
        with c_excel:
            df_export = df_editado.drop(columns=['codigo'], errors='ignore')
            st.download_button("📊 Exportar Excel", data=gerar_excel_download(df_export, f"Fechamento {setor}"), file_name=f"Separacao_Fechamento_{setor}.xlsx", use_container_width=True)
        with c_print: 
            injetar_botao_impressao()
            
        if btn_salvar:
            with st.spinner("Atualizando registros..."):
                cods = df_editado["codigo"].tolist()
                if cods:
                    for loja_nome in LOJAS_NOMES:
                        n_loja = int(loja_nome.split()[-1])
                        supabase.table("pedidos").delete().eq("setor", setor).eq("loja", n_loja).eq("data_pedido", str(date.today())).in_("codigo_produto", cods).execute()
                        
                        lista_ins = []
                        for _, r in df_editado.iterrows():
                            q = converter_para_int_seguro(r[loja_nome])
                            if q > 0: 
                                lista_ins.append({"data_pedido": str(date.today()), "setor": setor, "loja": n_loja, "codigo_produto": int(r["codigo"]), "quantidade": q, "usuario": usuario_atual})
                        
                        if lista_ins: 
                            supabase.table("pedidos").insert(lista_ins).execute()
            st.success("Alterações consolidadas!")
            st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 2 — VISÃO DAS LOJAS
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS_NOMES) if acesso_total else usuario_atual
        num_loja = int(loja_selecionada.split()[-1])

        st.markdown(f"<div class='no-print'><h2>🥬 Lançamento de Pedidos — {loja_selecionada}</h2></div>", unsafe_allow_html=True)
        
        resp_prod = supabase.table("produtos").select("codigo, codigo_erp, descricao, fornecedor, nome_personalizado").eq("setor", setor).eq("ativo", True).execute()
        df_prod = pd.DataFrame(resp_prod.data)

        if df_prod.empty: 
            st.warning("Nenhum produto cadastrado para este setor.")
            return

        if 'codigo_erp' not in df_prod.columns: 
            df_prod['codigo_erp'] = df_prod['codigo']
            
        df_prod['codigo_erp'] = df_prod['codigo_erp'].fillna(df_prod['codigo']).astype(int)

        codigos_setor = df_prod['codigo'].tolist()
        df_perm = buscar_permissoes_setor(supabase, codigos_setor, num_loja)
        
        resp_med = supabase.table("medias_90d").select("codigo_produto, media_dia").eq("loja", num_loja).execute()
        resp_exis = supabase.table("pedidos").select("codigo_produto, quantidade").eq("setor", setor).eq("loja", num_loja).eq("data_pedido", str(date.today())).execute()

        df_med = pd.DataFrame(resp_med.data)
        df_exis = pd.DataFrame(resp_exis.data)

        df_prod['descricao'] = df_prod['nome_personalizado'].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else None).fillna(df_prod['descricao'])

        df_loja = pd.merge(df_prod, df_perm, left_on='codigo', right_on='codigo_produto', how='left')
        if 'disponivel' not in df_loja.columns: 
            df_loja['disponivel'] = True
        else: 
            df_loja['disponivel'] = df_loja['disponivel'].fillna(True)
            
        df_loja = df_loja[df_loja['disponivel'] == True]

        if df_loja.empty: 
            st.warning("Nenhum produto liberado para esta loja neste setor.")
            return

        df_loja = pd.merge(df_loja, df_med, on='codigo_produto', how='left')
        df_loja['media_dia'] = df_loja['media_dia'].fillna(0.0)

        if not df_exis.empty:
            df_loja = pd.merge(df_loja, df_exis, on='codigo_produto', how='left')
            df_loja['quantidade'] = df_loja['quantidade'].fillna(0).astype(int)
            itens_digitados = df_exis[df_exis['quantidade'] > 0].shape[0]
        else:
            df_loja['quantidade'] = 0
            itens_digitados = 0

        st.metric(label="📝 Seus Itens Preenchidos", value=f"{itens_digitados} produtos")

        codigos_busca_erp = df_loja["codigo_erp"].dropna().astype(int).unique().tolist()
        df_estoque = buscar_estoque_erp(loja_selecionada, codigos_busca_erp, setor)
        df_loja = pd.merge(df_loja, df_estoque, left_on='codigo_erp', right_on='Código', how='left')
        df_loja["Estoque"] = df_loja["Estoque"].fillna(0).astype(int)
        
        df_loja['quantidade'] = df_loja['quantidade'].replace({0: ""}).astype(str).replace({"0": "", "0.0": "", "nan": ""})

        df_final_grid = pd.DataFrame({
            'codigo': df_loja['codigo'], 
            'Cód. ERP': df_loja['codigo_erp'],
            'Fornecedor': df_loja['fornecedor'], 
            'Descrição': df_loja['descricao'],
            'Média (90d)': df_loja['media_dia'], 
            'Estoque ERP': df_loja['Estoque'],
            'Qtde Pedida': df_loja['quantidade']
        }).sort_values(by=['Fornecedor', 'Descrição'])

        texto_busca = st.text_input("🔍 Buscar Produto (por Código ou Nome):")
        if texto_busca:
            tb = texto_busca.lower().strip()
            mask = df_final_grid['Descrição'].str.lower().str.contains(tb, na=False) | df_final_grid['Cód. ERP'].astype(str).str.contains(tb, na=False)
            df_filtrado = df_final_grid[mask]
        else:
            df_filtrado = df_final_grid

        col_cfg_l = {
            "codigo": None,
            "Cód. ERP": st.column_config.NumberColumn(disabled=True, format="%d", width=80), 
            "Fornecedor": st.column_config.TextColumn(disabled=True, width=120), 
            "Descrição": st.column_config.TextColumn(disabled=True, width=250),
            "Estoque ERP": st.column_config.NumberColumn(disabled=True, format="%d", width=90), 
            "Média (90d)": st.column_config.NumberColumn(disabled=True, format="%.2f", width=90),
            "Qtde Pedida": st.column_config.TextColumn("Qtde", width=100)
        }

        grid_editado = st.data_editor(df_filtrado, column_config=col_cfg_l, hide_index=True, use_container_width=True, key=f"grid_loja_{num_loja}")

        html_table = df_filtrado.drop(columns=['codigo'], errors='ignore').fillna('').to_html(index=False, classes="print-table")
        st.markdown(f'<div class="print-only"><h3>🥬 Pedido Oficial — {loja_selecionada}</h3>{html_table}</div>', unsafe_allow_html=True)

        c_salvar, c_excel, c_print = st.columns([2, 2, 1])
        with c_salvar: 
            btn_salvar_loja = st.button("💾 Salvar Pedido Oficial", type="primary", use_container_width=True)
        with c_excel:
            df_export = grid_editado.drop(columns=['codigo'], errors='ignore')
            st.download_button("📊 Exportar Excel", data=gerar_excel_download(df_export, f"Pedido"), file_name=f"Pedido_{loja_selecionada}.xlsx", use_container_width=True)
        with c_print: 
            injetar_botao_impressao()

        if btn_salvar_loja:
            with st.spinner("Gravando pedido..."):
                cods_tela = grid_editado["codigo"].tolist()
                if cods_tela:
                    supabase.table("pedidos").delete().eq("setor", setor).eq("loja", num_loja).eq("data_pedido", str(date.today())).in_("codigo_produto", cods_tela).execute()
                    
                    lista_ins = []
                    for _, r in grid_editado.iterrows():
                        q = converter_para_int_seguro(r["Qtde Pedida"])
                        if q > 0: 
                            lista_ins.append({"data_pedido": str(date.today()), "setor": setor, "loja": num_loja, "codigo_produto": int(r["codigo"]), "quantidade": q, "usuario": usuario_atual})
                    
                    if lista_ins: 
                        supabase.table("pedidos").insert(lista_ins).execute()
            st.success("Gravado!")
            st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 3 — VISÃO FORNECEDORES (RESUMO EDITÁVEL BANDEIRADO)
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Visão Fornecedores (Resumo)":
        st.markdown(f"<div class='no-print'><h2>🚚 Resumo Consolidado por Fornecedor — {setor}</h2></div>", unsafe_allow_html=True)
        
        resp_prod = supabase.table("produtos").select("codigo, codigo_erp, descricao, fornecedor, nome_personalizado").eq("setor", setor).execute()
        resp_ped = supabase.table("pedidos").select("codigo_produto, loja, quantidade").eq("setor", setor).eq("data_pedido", str(date.today())).execute()
        
        df_prod = pd.DataFrame(resp_prod.data)
        df_ped = pd.DataFrame(resp_ped.data)

        if df_prod.empty: 
            st.info("Nenhum produto cadastrado.")
            return

        if 'codigo_erp' not in df_prod.columns: 
            df_prod['codigo_erp'] = df_prod['codigo']
            
        df_prod['codigo_erp'] = df_prod['codigo_erp'].fillna(df_prod['codigo']).astype(int)

        codigos_setor = df_prod['codigo'].tolist()
        df_perm_all = buscar_permissoes_setor(supabase, codigos_setor)
        df_prod['descricao'] = df_prod['nome_personalizado'].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else None).fillna(df_prod['descricao'])

        if not df_ped.empty:
            df_pivot = df_ped.pivot_table(index='codigo_produto', columns='loja', values='quantidade', aggfunc='sum').reset_index()
            for n in range(1, 9):
                if n in df_pivot.columns: 
                    df_pivot = df_pivot.rename(columns={n: f"Loja {n:02d}"})
        else: 
            df_pivot = pd.DataFrame(columns=['codigo_produto'])
            
        df_mestre = pd.merge(df_prod, df_pivot, left_on='codigo', right_on='codigo_produto', how='left')
        
        for l in LOJAS_NOMES:
            if l not in df_mestre.columns: 
                df_mestre[l] = 0.0
            df_mestre[l] = df_mestre[l].fillna(0).astype(int)

        if not df_perm_all.empty:
            df_perm_all['loja_nome'] = df_perm_all['loja'].apply(lambda x: f"Loja {int(x):02d}_perm")
            df_perm_pivot = df_perm_all.pivot_table(index='codigo_produto', columns='loja_nome', values='disponivel', aggfunc='last').reset_index()
        else: 
            df_perm_pivot = pd.DataFrame(columns=['codigo_produto'])

        df_mestre = pd.merge(df_mestre, df_perm_pivot, left_on='codigo', right_on='codigo_produto', how='left')
        
        perm_cols = [f"Loja {n:02d}_perm" for n in range(1, 9)]
        for col in perm_cols:
            if col not in df_mestre.columns: 
                df_mestre[col] = True
                
        df_mestre[perm_cols] = df_mestre[perm_cols].fillna(True)
        mask_active = df_mestre[perm_cols].any(axis=1)
        df_mestre = df_mestre[mask_active]

        df_mestre["TOTAL_NUM"] = df_mestre[LOJAS_NOMES].sum(axis=1)
        df_mestre["TOTAL GERAL"] = df_mestre["TOTAL_NUM"].replace({0: ""})

        for l in LOJAS_NOMES:
            df_mestre[l] = df_mestre[l].replace({0: ""}).astype(str).replace({"0": "", "0.0": "", "nan": ""})
            perm_col = f"{l}_perm"
            if perm_col in df_mestre.columns: 
                df_mestre.loc[df_mestre[perm_col] != True, l] = "-"

        all_edited_frames = []

        for forn in sorted(df_mestre["fornecedor"].dropna().unique()):
            df_forn_bruto = df_mestre[df_mestre["fornecedor"] == forn]
            lojas_ativas = [l for l in LOJAS_NOMES if not (df_forn_bruto[l] == "-").all()]
            df_forn_view = df_forn_bruto[["codigo", "codigo_erp", "descricao"] + lojas_ativas + ["TOTAL GERAL"]].rename(columns={'codigo_erp': 'Cód. ERP', 'descricao': 'Produto'}).sort_values(by='Produto')
            
            with st.container(border=True):
                st.markdown(f"<div class='no-print'><h5>Fornecedor: {forn}</h5></div>", unsafe_allow_html=True)
                col_cfg_f = {
                    "codigo": None,
                    "Cód. ERP": st.column_config.NumberColumn(disabled=True, width=80, format="%d"),
                    "Produto": st.column_config.TextColumn(disabled=True, width=250),
                    "TOTAL GERAL": st.column_config.TextColumn("TOTAL", disabled=True, width=70)
                }
                for l in lojas_ativas: 
                    col_cfg_f[l] = st.column_config.TextColumn(l, width=85, disabled=False)
                    
                edit_df = st.data_editor(df_forn_view, hide_index=True, use_container_width=False, column_config=col_cfg_f, key=f"editor_forn_{forn}")
                
                html_table = df_forn_view.drop(columns=['codigo'], errors='ignore').fillna('').to_html(index=False, classes="print-table")
                st.markdown(f'<div class="print-only"><h4 class="supplier-header">🚚 Fornecedor: {forn}</h4>{html_table}</div>', unsafe_allow_html=True)
                
                for l in LOJAS_NOMES:
                    if l not in edit_df.columns: 
                        edit_df[l] = "-"
                
                edit_df["fornecedor"] = forn
                all_edited_frames.append(edit_df)

        st.markdown("<div class='no-print'><br></div>", unsafe_allow_html=True)
        
        if all_edited_frames:
            df_forn_editado_full = pd.concat(all_edited_frames, ignore_index=True)
            c_salvar, c_excel, c_print = st.columns([2, 2, 1])
            with c_salvar: 
                btn_salvar_forn = st.button("💾 Salvar Ajustes do Resumo", type="primary", use_container_width=True)
            with c_excel:
                df_export = df_forn_editado_full.drop(columns=['codigo'], errors='ignore')
                st.download_button("📊 Exportar Fornecedores", data=gerar_excel_fornecedores(df_export, f"Fornecedores"), file_name=f"Resumo_Fornecedores_{setor}.xlsx", use_container_width=True)
            with c_print: 
                injetar_botao_impressao()

            if btn_salvar_forn:
                with st.spinner("Atualizando registros..."):
                    cods = df_forn_editado_full["codigo"].tolist()
                    if cods:
                        for loja_nome in LOJAS_NOMES:
                            n_loja = int(loja_nome.split()[-1])
                            supabase.table("pedidos").delete().eq("setor", setor).eq("loja", n_loja).eq("data_pedido", str(date.today())).in_("codigo_produto", cods).execute()
                            
                            lista_ins = []
                            for _, r in df_forn_editado_full.iterrows():
                                q = converter_para_int_seguro(r[loja_nome])
                                if q > 0: 
                                    lista_ins.append({"data_pedido": str(date.today()), "setor": setor, "loja": n_loja, "codigo_produto": int(r["codigo"]), "quantidade": q, "usuario": usuario_atual})
                            
                            if lista_ins: 
                                supabase.table("pedidos").insert(lista_ins).execute()
                st.success("Alterações consolidadas!")
                st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # ROTA 4 — CATÁLOGO DE PRODUTOS
    # ─────────────────────────────────────────────────────────────────────────
    elif perfil_navegacao == "Catálogo de Produtos":
        st.markdown(f"<div class='no-print'><h2>🗂️ Gestão de Catálogo e Permissões por Loja — {setor}</h2></div>", unsafe_allow_html=True)

        resp_prod = supabase.table("produtos").select("*").eq("setor", setor).execute()
        df_prod = pd.DataFrame(resp_prod.data)

        if df_prod.empty: 
            df_prod = pd.DataFrame(columns=['codigo', 'codigo_erp', 'descricao', 'fornecedor', 'nome_personalizado', 'setor', 'ativo'])
        else:
            if 'codigo_erp' not in df_prod.columns: 
                df_prod['codigo_erp'] = df_prod['codigo']
            df_prod['codigo_erp'] = df_prod['codigo_erp'].fillna(df_prod['codigo']).astype(int)

        codigos_setor = df_prod['codigo'].tolist() if not df_prod.empty else []
        df_perm = buscar_permissoes_setor(supabase, codigos_setor)

        if not df_perm.empty:
            df_perm['loja_nome'] = df_perm['loja'].apply(lambda x: f"Loja {int(x):02d}")
            df_perm_pivot = df_perm.pivot_table(index='codigo_produto', columns='loja_nome', values='disponivel', aggfunc='last').reset_index()
        else: 
            df_perm_pivot = pd.DataFrame(columns=['codigo_produto'] + LOJAS_NOMES)

        if not df_prod.empty: 
            df_cat_completo = pd.merge(df_prod, df_perm_pivot, left_on='codigo', right_on='codigo_produto', how='left').drop(columns=['codigo_produto'], errors='ignore')
        else: 
            df_cat_completo = pd.DataFrame(columns=['codigo', 'codigo_erp', 'descricao', 'fornecedor', 'nome_personalizado'] + LOJAS_NOMES)
        
        for l in LOJAS_NOMES: 
            if l not in df_cat_completo.columns: 
                df_cat_completo[l] = True
            else: 
                df_cat_completo[l] = df_cat_completo[l].fillna(True).astype(bool)
        
        if 'nome_personalizado' not in df_cat_completo.columns: 
            df_cat_completo['nome_personalizado'] = ""
        else: 
            df_cat_completo['nome_personalizado'] = df_cat_completo['nome_personalizado'].fillna("")

        if not df_cat_completo.empty: 
            df_cat_completo = df_cat_completo.sort_values(by=['fornecedor', 'descricao']).reset_index(drop=True)

        col_cfg_c = {
            "codigo": None,
            "codigo_erp": st.column_config.NumberColumn("Cód. ERP", format="%d", width=80), 
            "descricao": st.column_config.TextColumn("Nome Prime", width=180), 
            "nome_personalizado": st.column_config.TextColumn("Nome Manual", width=160),
            "fornecedor": st.column_config.TextColumn("Fornecedor", width=130)
        }
        for l in LOJAS_NOMES: 
            col_cfg_c[l] = st.column_config.CheckboxColumn(l)

        cols_exibicao = ["fornecedor", "codigo", "codigo_erp", "descricao", "nome_personalizado"] + LOJAS_NOMES
        edited_cat = st.data_editor(df_cat_completo[cols_exibicao], use_container_width=True, hide_index=True, column_config=col_cfg_c, num_rows="dynamic", key="catalogo_editor")

        html_table = df_cat_completo[cols_exibicao].drop(columns=['codigo'], errors='ignore').fillna('').to_html(index=False, classes="print-table")
        st.markdown(f'<div class="print-only"><h3>🗂️ Catálogo Geral — {setor}</h3>{html_table}</div>', unsafe_allow_html=True)

        st.markdown("<div class='no-print'><br></div>", unsafe_allow_html=True)
        col_btn_salvar, col_btn_erp = st.columns(2)
        
        with col_btn_salvar: 
            btn_salvar = st.button("💾 Salvar Matriz do Catálogo", type="primary", use_container_width=True)
        with col_btn_erp: 
            btn_puxar_erp = st.button("📥 Puxar Nomes do ERP", use_container_width=True)

        if btn_salvar:
            state = st.session_state.get("catalogo_editor")
            with st.spinner("Automação Duplo-Código processando..."):
                try:
                    resp_all = supabase.table("produtos").select("codigo").execute()
                    codigos_globais = [p["codigo"] for p in resp_all.data] if resp_all.data else []
                    codigos_conhecidos = set(df_cat_completo['codigo'].dropna().astype(int).tolist()) if not df_cat_completo.empty else set()
                    
                    mapa_conflitos = {}

                    if state and state.get("deleted_rows"):
                        for idx in state["deleted_rows"]:
                            cod_p = int(df_cat_completo.iloc[idx]["codigo"])
                            supabase.table("produtos_lojas").delete().eq("codigo_produto", cod_p).execute()
                            supabase.table("pedidos").delete().eq("codigo_produto", cod_p).execute()
                            supabase.table("medias_90d").delete().eq("codigo_produto", cod_p).execute()
                            supabase.table("produtos").delete().eq("codigo", cod_p).execute()
                            if cod_p in codigos_globais: codigos_globais.remove(cod_p)
                            if cod_p in codigos_conhecidos: codigos_conhecidos.remove(cod_p)

                    if state and state.get("edited_rows"):
                        for idx_str, changes in state["edited_rows"].items():
                            idx = int(idx_str)
                            cod_p_original = int(df_cat_completo.iloc[idx]["codigo"])
                            prod_changes = {}
                            if "descricao" in changes: 
                                prod_changes["descricao"] = str(changes["descricao"])
                            if "fornecedor" in changes: 
                                prod_changes["fornecedor"] = str(changes["fornecedor"])
                            if "nome_personalizado" in changes: 
                                val_np = changes["nome_personalizado"]
                                prod_changes["nome_personalizado"] = str(val_np).strip() if pd.notna(val_np) and str(val_np).strip() != "" else None
                            if "codigo_erp" in changes:
                                try: prod_changes["codigo_erp"] = int(changes["codigo_erp"])
                                except: pass
                            if prod_changes: 
                                supabase.table("produtos").update(prod_changes).eq("codigo", cod_p_original).execute()

                    for _, row in edited_cat.iterrows():
                        c_pk = row.get("codigo")
                        c_erp = row.get("codigo_erp") 
                        if pd.isna(c_pk) or str(c_pk).strip() == "":
                            if pd.isna(c_erp) or str(c_erp).strip() == "": continue
                            cod_erp_digitado = int(c_erp)
                            cod_final = cod_erp_digitado
                            forn_add = str(row.get("fornecedor", "Box")).strip()
                            desc_add = str(row.get("descricao", "Novo Produto")).strip()
                            np_add = str(row.get("nome_personalizado", "")).strip() if pd.notna(row.get("nome_personalizado")) and str(row.get("nome_personalizado")).strip() != "" else None
                            if cod_final in codigos_globais:
                                base_str = str(cod_final)
                                for i in range(1, 100):
                                    tent = int(f"{base_str}{i:02d}")
                                    if tent not in codigos_globais:
                                        cod_final = tent
                                        break
                                st.toast(f"🤖 Gerado código interno invisível {cod_final} para o item {cod_erp_digitado} do ERP.", icon="✨")
                            mapa_conflitos[(cod_erp_digitado, forn_add)] = cod_final
                            supabase.table("produtos").insert({
                                "codigo": cod_final, "codigo_erp": cod_erp_digitado,
                                "descricao": desc_add, "fornecedor": forn_add, 
                                "nome_personalizado": np_add, "setor": setor, "ativo": True
                            }).execute()
                            codigos_globais.append(cod_final); codigos_conhecidos.add(cod_final)

                    lista_perms_geral = []
                    codigos_processados_perms = set()
                    
                    for _, row in edited_cat.iterrows():
                        c_pk = row.get("codigo")
                        c_erp = row.get("codigo_erp")
                        f_tela = str(row.get("fornecedor", "")).strip()
                        if pd.notna(c_pk) and str(c_pk).strip() != "": c_final = int(c_pk)
                        else: c_final = mapa_conflitos.get((int(c_erp), f_tela), None)
                        if not c_final: continue
                        
                        codigos_processados_perms.add(c_final)
                        for num_loja in range(1, 9):
                            val_loja = bool(row.get(f"Loja {num_loja:02d}", True))
                            lista_perms_geral.append({"codigo_produto": c_final, "loja": num_loja, "disponivel": val_loja})

                    codigos_lista = list(codigos_processados_perms)
                    if codigos_lista:
                        for i in range(0, len(codigos_lista), 200): supabase.table("produtos_lojas").delete().in_("codigo_produto", codigos_lista[i:i+200]).execute()
                        for i in range(0, len(lista_perms_geral), 1000): supabase.table("produtos_lojas").insert(lista_perms_geral[i:i+1000]).execute()

                    st.success("✅ Automação concluída!"); time.sleep(1.5); st.rerun()
                except Exception as e: st.error(f"⚠️ Erro processando: {e}")

        if btn_puxar_erp:
            with st.spinner("Buscando nomes oficias usando Cód. ERP..."):
                try:
                    cods_erp = [int(c) for c in edited_cat["codigo_erp"].tolist() if pd.notna(c) and str(c).strip() != ""]
                    if not cods_erp: st.warning("Nenhum código encontrado.")
                    else:
                        cods_str = ", ".join(map(str, set(cods_erp)))
                        query_nomes = f"SELECT cod, descricao FROM python_ajuste_cadastro WHERE cod IN ({cods_str})"
                        df_nomes = conn_pg.query(query_nomes, ttl=0)

                        if not df_nomes.empty:
                            for _, row in df_nomes.iterrows():
                                cod_oficial = int(row["cod"])
                                desc_erp = str(row["descricao"])
                                supabase.table("produtos").update({"descricao": desc_erp}).eq("codigo_erp", cod_oficial).execute()
                            st.success("✅ Nomes atualizados em todos os fornecedores!"); time.sleep(1); st.rerun()
                        else: st.info("Nenhum nome encontrado.")
                except Exception as e:
                    if "No database configured" in str(e) or "missing" in str(e).lower(): st.error("⚠️ Aviso: Credenciais do PostgreSQL não configuradas ou inacessíveis.")
                    else: st.error(f"⚠️ Erro ao buscar nomes no banco ERP: {e}")
