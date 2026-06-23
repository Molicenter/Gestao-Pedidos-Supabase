import streamlit as st
import pandas as pd
import io
import requests
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime
import pytz

# ─────────────────────────────────────────────
# CONSTANTES E PRODUTOS INICIAIS DA MATRIZ
# ─────────────────────────────────────────────
LOJAS = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]
MAPA_LOJAS = {l: l for l in LOJAS}

produtos_iniciais = [
    {"Fornecedor": "BOVINOS", "Código": 128537, "Descrição Oficial": "Bisteca Bov Kg S/File Congelada", "Nome Personalizado": ""},
    {"Fornecedor": "BOVINOS", "Código": 280989, "Descrição Oficial": "Contra File Bov Kg Agra Promoção", "Nome Personalizado": ""},
    {"Fornecedor": "BOVINOS", "Código": 115454, "Descrição Oficial": "Costela Minga Bov Kg Cong", "Nome Personalizado": ""},
    {"Fornecedor": "BOVINOS", "Código": 130815, "Descrição Oficial": "Coxao Mole Bov Kg Agra Promoção", "Nome Personalizado": ""},
    {"Fornecedor": "BOVINOS", "Código": 14717, "Descrição Oficial": "Mocoto Bov Kg Bdj", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 7696, "Descrição Oficial": "Ling Tosc Kg Frimesa", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 214463, "Descrição Oficial": "Ling Suina Kg Sadia", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 344755, "Descrição Oficial": "Ling Pernil Kg Frimesa", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 377869, "Descrição Oficial": "Ling Suina Kg Seara", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 609449, "Descrição Oficial": "Ling Suina 1kg Mimosa", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 611718, "Descrição Oficial": "Ling Frango 700g Lar Cong", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 624314, "Descrição Oficial": "Ling Churrasco Kg Alegra", "Nome Personalizado": ""},
    {"Fornecedor": "LINGUIÇAS", "Código": 641551, "Descrição Oficial": "Ling Frango 700g Copacol Fina Cong", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 352505, "Descrição Oficial": "Ancho Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 597588, "Descrição Oficial": "Aranha Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 597612, "Descrição Oficial": "Bananinha Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 597597, "Descrição Oficial": "Contra File Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 440457, "Descrição Oficial": "Cupim Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 181156, "Descrição Oficial": "Fraldinha Bov Kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "COCAMAR", "Código": 597630, "Descrição Oficial": "Picanha Bov Fatiada kg Cocamar", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 115612, "Descrição Oficial": "PORCO TIPO EXPOTAÇÃO", "Nome Personalizado": "PORCO TIPO EXPORTAÇÃO"},
    {"Fornecedor": "SUINO", "Código": 35170, "Descrição Oficial": "Bisteca Suina Kg Copa Congelado", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 7483, "Descrição Oficial": "Bisteca Suina Kg Lombo Congelado", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 7492, "Descrição Oficial": "Costela Suina Kg S/ Couro Band", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 16544, "Descrição Oficial": "Paleta Suina Kg C/Couro", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 16926, "Descrição Oficial": "Panceta Suina Kg", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 14960, "Descrição Oficial": "Pernil Suino Kg C/Couro", "Nome Personalizado": ""},
    {"Fornecedor": "SUINO", "Código": 7517, "Descrição Oficial": "Toucinho Suino Kg", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 318273, "Descrição Oficial": "File Sassami 1kg Copacol Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 341776, "Descrição Oficial": "File Sassami 800g Copacol \"1a1\"", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 341785, "Descrição Oficial": "Coxa Sobr Passarinho 800g Copacol \"1a1\"", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 415071, "Descrição Oficial": "File Sassami 1kg Nat Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 415080, "Descrição Oficial": "Coxinha Asa 1kg Nat", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 420338, "Descrição Oficial": "File Coxa Sobrecoxa 1kg Nat Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 420347, "Descrição Oficial": "Meio Asa 1kg Nat Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 421373, "Descrição Oficial": "Coxinha Asa 1kg Seara Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 455389, "Descrição Oficial": "Moela Frango 1kg Perdigao Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 475967, "Descrição Oficial": "Coxinha Asa 1kg Cancao Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 491664, "Descrição Oficial": "File Sassami 750g Pioneiro Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 492450, "Descrição Oficial": "Sobrecoxa 1kg Seara Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 493644, "Descrição Oficial": "File Peito 800g Copacol Bifes Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 544823, "Descrição Oficial": "File Coxa Sobr 800g Copacol S/Pele Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 567374, "Descrição Oficial": "Figado Frango 500g Levo Bdj Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 587703, "Descrição Oficial": "Moela Frango 700g Lar Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 587712, "Descrição Oficial": "Figado Frango 700g Lar Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 595872, "Descrição Oficial": "File Sassami 1kg Seara Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 597296, "Descrição Oficial": "Frango Passarinho 700g Lar Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 608699, "Descrição Oficial": "Sobrecoxa 1kg Nat Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 614061, "Descrição Oficial": "File Peito 1kg Nat Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 640279, "Descrição Oficial": "Coracao Frango 500g Nat Cong Band", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 647485, "Descrição Oficial": "Meio Asa Temp 500g Copacol Mestre Cong", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 650197, "Descrição Oficial": "Coxinha Asa 750g Pioneiro Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 657260, "Descrição Oficial": "File Sassami 1kg Cancao Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 657288, "Descrição Oficial": "Meio Asa 800g Cancao Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 662688, "Descrição Oficial": "Meio Asa 1kg Jagua", "Nome Personalizado": ""},
    {"Fornecedor": "CONGELADOS - PADRÃO E IQF", "Código": 662721, "Descrição Oficial": "Meio Asa 750g Pioneiro Iqf", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 109824, "Descrição Oficial": "Coxa Sobrecoxa Kg", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 7553, "Descrição Oficial": "Coxa Sobrecoxa Kg Copacol", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 111582, "Descrição Oficial": "Coxa Sobrecoxa Kg Dorsal", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 140982, "Descrição Oficial": "Dorso Kg", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 109833, "Descrição Oficial": "File Peito Kg", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 46015, "Descrição Oficial": "File Peito Kg Copacol", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 99547, "Descrição Oficial": "Frango Caipira Kg Nho Bento", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 359100, "Descrição Oficial": "Frango Kg S/Miudos Pioneiro", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 79761, "Descrição Oficial": "Frango Kg", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 527048, "Descrição Oficial": "Frango Kg Seara", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 7562, "Descrição Oficial": "Peito Frango Kg Copacol", "Nome Personalizado": ""},
    {"Fornecedor": "DIVERSOS", "Código": 139773, "Descrição Oficial": "Pe Frango Kg Bdj", "Nome Personalizado": ""}
]

# ─────────────────────────────────────────────
# FUNÇÕES DE BANCO DE DADOS E AUXILIARES
# ─────────────────────────────────────────────
def gerar_excel_estilizado(df, sheet_name="Resumo"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color='6B1D1D', end_color='6B1D1D', fill_type='solid')
        alt_row_fill = PatternFill(start_color='F9F2F2', end_color='F9F2F2', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        border_style = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column)):
            for cell in row:
                cell.border = border_style
                if row_idx == 0:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    if row_idx % 2 == 0:
                        cell.fill = alt_row_fill
                    if cell.column_letter not in ['C']: 
                        cell.alignment = Alignment(horizontal='center', vertical='center')

        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_length + 2

        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = False
        worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.75
        worksheet.page_margins.bottom = 0.75

    return buffer.getvalue()

conn = st.connection("gsheets", type=GSheetsConnection)

WS_PRODUTOS = "Acougue_Produtos"
WS_PEDIDOS  = "Acougue_Pecas"
WS_MEDIA_90D = "Acougue_90d"  

def parse_bool(x):
    if isinstance(x, bool): return x
    if isinstance(x, (int, float)): return bool(x) and not pd.isna(x)
    return str(x).strip().upper() in ['TRUE', 'VERDADEIRO', '1', 'V', 'SIM', 'YES', 'T', 'X']

def notificar_telegram(mensagem):
    try:
        bot_token = st.secrets["telegram"]["bot_token"]
        chat_id = st.secrets["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        payload = {"chat_id": chat_id, "text": mensagem}
        
        resposta = requests.post(url, json=payload)
        
        if resposta.status_code == 200:
            return True
        else:
            st.error(f"🚨 Erro do Telegram: {resposta.text}")
            return False
            
    except KeyError:
        st.error("⚠️ Credenciais do Telegram não configuradas no secrets.toml.")
        return False
    except Exception as e:
        st.error(f"⚠️ Erro de conexão ao enviar o Telegram: {e}")
        return False

@st.cache_data(ttl=15)
def carregar_catalogo_acougue():
    try:
        df = conn.read(worksheet=WS_PRODUTOS, ttl=0, usecols=list(range(20)))
    except ValueError as e:
        if "Spreadsheet must be specified" in str(e):
            st.error("🚨 **Erro Crítico:** URL da Planilha não especificada nas configurações do Streamlit Cloud.")
            st.stop()
        else:
            raise e
    except Exception as e:
        st.error(f"Erro ao conectar na aba {WS_PRODUTOS}: {e}")
        st.stop()
        
    if df.empty or "Fornecedor" not in df.columns:
        df_init = pd.DataFrame(produtos_iniciais)
        for loja in LOJAS: df_init[loja] = True
        conn.update(worksheet=WS_PRODUTOS, data=df_init)
        df = df_init.copy()
        
    need_update = False
    
    if "Descrição" in df.columns and "Descrição Oficial" not in df.columns:
        df = df.rename(columns={"Descrição": "Descrição Oficial"})
        need_update = True
        
    if "Nome Personalizado" not in df.columns:
        df["Nome Personalizado"] = ""
        df.loc[df["Código"] == 115612, "Nome Personalizado"] = "PORCO TIPO EXPORTAÇÃO"
        need_update = True
        
    if need_update:
        conn.update(worksheet=WS_PRODUTOS, data=df.drop(columns=["Descrição"], errors="ignore"))
    
    novas_colunas = {}
    for col in df.columns:
        col_str = str(col).strip()
        for loja in LOJAS:
            if loja.lower() in col_str.lower():
                novas_colunas[col] = loja
    df = df.rename(columns=novas_colunas)
    
    for loja in LOJAS:
        if loja not in df.columns:
            df[loja] = False
        else:
            df[loja] = df[loja].apply(parse_bool)

    if "Código" in df.columns:
        df["Código"] = pd.to_numeric(df["Código"], errors='coerce').fillna(0).astype(int)
        
    def obter_nome_final(row):
        apelido = str(row.get("Nome Personalizado", "")).strip()
        if apelido and apelido.lower() != "nan":
            return apelido
        return str(row.get("Descrição Oficial", "")).strip()
        
    df["Descrição"] = df.apply(obter_nome_final, axis=1)
            
    return df

@st.cache_data(ttl=15)
def carregar_pedidos():
    try:
        df_pedidos = conn.read(worksheet=WS_PEDIDOS, ttl=0)
    except Exception as e:
        st.error(f"Erro ao conectar na aba {WS_PEDIDOS}: {e}")
        st.stop()

    df_cat = carregar_catalogo_acougue()
    
    if df_pedidos.empty or "Código" not in df_pedidos.columns:
        df_init = df_cat[["Fornecedor", "Código", "Descrição"]].copy()
        for loja in LOJAS:
            df_init[loja] = 0
        if not df_init.empty:
            conn.update(worksheet=WS_PEDIDOS, data=df_init)
        return df_init

    df_pedidos = df_pedidos.drop(columns=["Descrição", "Fornecedor"], errors="ignore")
    df_pedidos = pd.merge(df_cat[["Código", "Fornecedor", "Descrição"]], df_pedidos, on="Código", how="left")

    if "Código" in df_pedidos.columns:
        df_pedidos["Código"] = pd.to_numeric(df_pedidos["Código"], errors='coerce').fillna(0).astype(int)
        
    for loja in LOJAS:
        if loja in df_pedidos.columns:
            df_pedidos[loja] = pd.to_numeric(df_pedidos[loja], errors='coerce').fillna(0).astype(int)
            
    return df_pedidos

@st.cache_data(ttl=60)
def carregar_media_90d():
    try:
        df = conn.read(worksheet=WS_MEDIA_90D, ttl=0)
        if df.empty or "codigo" not in df.columns:
            return pd.DataFrame(columns=["loja", "codigo", "qtde_media_semanal"])
        return df
    except Exception as e:
        return pd.DataFrame(columns=["loja", "codigo", "qtde_media_semanal"])

def salvar_pedidos(df_to_save):
    conn.update(worksheet=WS_PEDIDOS, data=df_to_save)
    st.cache_data.clear()

def salvar_catalogo(df_to_save):
    df_clean = df_to_save.drop(columns=["Descrição"], errors="ignore")
    conn.update(worksheet=WS_PRODUTOS, data=df_clean)
    st.cache_data.clear()

# ─────────────────────────────────────────────────────────────
# MÓDULO PRINCIPAL - A FUNÇÃO QUE O MAIN.PY VAI CHAMAR
# ─────────────────────────────────────────────────────────────
def iniciar_tela():

    if 'reset_counter_acougue_adriano' not in st.session_state:
        st.session_state['reset_counter_acougue_adriano'] = 0

    usuario_atual = st.session_state.get('usuario_logado', None)
    acesso_total  = (usuario_atual == "Administrador")

    # CSS Global e de Impressão encapsulado
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;700&display=swap');

    :root {
        --bg-main:        #0d1117;
        --bg-card:        #161b22;
        --bg-sidebar:     #0d1117;
        --red-dark:       #3a1010;
        --red-mid:        #6b1d1d;
        --red-accent:     #b32d2d;
        --red-bright:     #e63939;
        --red-glow:       rgba(179, 45, 45, .25);
        --text-primary:   #e6edf3;
        --text-muted:     #7d8590;
        --text-header:    #ffcccc;
        --border:         #21262d;
        --border-active:  #b32d2d;
        --row-hover:      rgba(179, 45, 45, .08);
        --row-selected:   rgba(179, 45, 45, .18);
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--red-mid) 0%, var(--red-accent) 100%) !important;
        color: #fff !important; border: 1px solid var(--red-accent) !important;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 4px 18px var(--red-glow) !important;
    }
    
    .title-input input {
        color: var(--red-bright) !important;
    }
    .title-input input:focus { border: 1px dashed var(--red-accent) !important; }

    [data-testid="stDataEditor"] [data-testid="glideDataEditor"] .gdg-header-cell,
    [data-testid="stDataEditor"] .dvn-stack .gdg-header {
        background-color: var(--red-dark) !important;
        color: var(--text-header) !important;
    }

    [data-testid="stDataEditor"] {
        border: 1px solid var(--red-mid) !important;
    }
    [data-testid="stDataEditor"] .gdg-cell.gdg-selected,
    [data-testid="stDataEditor"] .gdg-cell[data-state="focused"],
    [data-testid="stDataEditor"] .gdg-cell[aria-selected="true"] {
        background-color: var(--row-selected) !important;
        outline: 2px solid var(--red-accent) !important;
    }
    [data-testid="stDataEditor"] .gdg-row:hover .gdg-cell { background-color: var(--row-hover) !important; }

    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        border-color: var(--red-mid) !important;
    }
    [data-testid="stMetricValue"] { color: var(--red-bright) !important; }

    .topbar-loja {
        background: linear-gradient(90deg, var(--red-dark) 0%, #1a0808 100%);
        border: 1px solid var(--red-mid);
        border-radius: 10px;
        padding: 10px 18px;
        margin-bottom: 18px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .topbar-left { display: flex; align-items: center; gap: 12px; }
    .topbar-title { font-size: 18px; font-weight: 700; color: var(--text-header); }
    .topbar-sub { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
    .erp-badge { background-color: #2ea043; color: white; padding: 2px 8px; border-radius: 12px; font-size: 10px; font-weight: 600; margin-left: 8px;}

    /* REGRAS DE IMPRESSÃO */
    @media print {
        @page { margin: 2mm 5mm; }
        .main .block-container { padding-top: 0 !important; margin-top: 0 !important; max-width: 100% !important; }
        h3 { page-break-after: avoid !important; break-after: avoid !important; }
        .stApp, .main, body, html { background-color: #ffffff !important; background-image: none !important; color: #000000 !important; padding: 0 !important; margin: 0 !important; }
        header, [data-testid="stSidebar"], [data-testid="stHeader"] { display: none !important; }
        [data-testid="stElementContainer"], [data-testid="stHorizontalBlock"], div[data-testid="stVerticalBlockBorderWrapper"] { display: none !important; }
        [data-testid="stElementContainer"]:has(#print-section) { display: block !important; width: 100% !important; }
        #print-section { display: block !important; width: 100% !important; }
        #print-section h2 { font-size: 13px !important; margin: 0 0 5px 0 !important; padding-bottom: 3px !important; border-bottom: 1px solid #000 !important; color: #000 !important; display: block !important; text-align: center !important; }
        #print-section h3 { font-size: 10px !important; border-bottom: none !important; margin-top: 10px !important; margin-bottom: 3px !important; color: #000 !important; }
        .print-container { width: 100% !important; display: block !important;}
        table.print-table { width: 100% !important; border-collapse: collapse !important; color: #000000 !important; font-family: 'IBM Plex Sans', sans-serif !important; line-height: 1.1 !important; display: table !important; table-layout: fixed !important; margin-bottom: 3px !important; }
        table.print-table th, table.print-table td { border: 1px solid #000000 !important; padding: 0px 2px !important; text-align: left; color: #000000 !important; background-color: #ffffff !important; white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }
        table.print-table th { background-color: #e0e0e0 !important; font-weight: bold !important; text-align: center !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        table.print-table tr { break-inside: avoid !important; page-break-inside: avoid !important; }
        table.print-sep { font-size: 8px !important; }
        table.print-sep th:nth-child(1), table.print-sep td:nth-child(1) { width: 14% !important; }
        table.print-sep th:nth-child(2), table.print-sep td:nth-child(2) { width: 6% !important; text-align: center !important; }
        table.print-sep th:nth-child(3), table.print-sep td:nth-child(3) { width: 28% !important; }
        table.print-sep th:nth-child(n+4):nth-child(-n+11), table.print-sep td:nth-child(n+4):nth-child(-n+11) { width: 5.5% !important; text-align: center !important; }
        table.print-sep th:nth-child(12), table.print-sep td:nth-child(12) { width: 8% !important; text-align: center !important; font-weight: bold !important; background-color: #f0f0f0 !important; }
        table.print-loja { font-size: 10px !important; }
        table.print-loja th:nth-child(1), table.print-loja td:nth-child(1) { width: 12% !important; }
        table.print-loja th:nth-child(2), table.print-loja td:nth-child(2) { width: 8% !important; text-align: center !important; }
        table.print-loja th:nth-child(3), table.print-loja td:nth-child(3) { width: 40% !important; }
        table.print-loja th:nth-child(4), table.print-loja td:nth-child(4) { width: 12% !important; text-align: center !important; }
        table.print-loja th:nth-child(5), table.print-loja td:nth-child(5) { width: 12% !important; text-align: center !important; }
        table.print-loja th:nth-child(6), table.print-loja td:nth-child(6) { width: 16% !important; text-align: center !important; font-weight: bold !important; background-color: #f0f0f0 !important;}
        table.print-forn { font-size: 8.5px !important; }
        table.print-forn th:nth-child(1), table.print-forn td:nth-child(1) { width: 8% !important; text-align: center !important; }
        table.print-forn th:nth-child(2), table.print-forn td:nth-child(2) { width: 34% !important; }
        table.print-forn th:nth-child(n+3):nth-child(-n+10), table.print-forn td:nth-child(n+3):nth-child(-n+10) { width: 6% !important; text-align: center !important; }
        table.print-forn th:nth-child(11), table.print-forn td:nth-child(11) { width: 10% !important; text-align: center !important; font-weight: bold !important; background-color: #f0f0f0 !important;}
    }
    @media screen { #print-section { display: none !important; } }
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

    # SIDEBAR INTERNA
    with st.sidebar:
        st.markdown(f"### Módulo: Açougue Adriano")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação:", [
                "Visão por Fornecedor (Resumo)",
                "Separação e Fechamento",
                "Visão das Lojas",
                "Catálogo de Produtos" 
            ])
        else:
            perfil_navegacao = "Visão das Lojas"

        st.divider()
        
        df_ped = carregar_pedidos()
        if not df_ped.empty and set(LOJAS).issubset(df_ped.columns):
            total_preenchidos = (df_ped[LOJAS] > 0).any(axis=1).sum()
        else:
            total_preenchidos = 0
            
        st.metric("Itens c/ pedido", total_preenchidos, help="Itens com ao menos 1 quantidade preenchida")
        
        st.divider()
        
        if st.button("🔄 Sincronizar Dados", use_container_width=True):
            st.cache_data.clear()
            st.session_state['reset_counter_acougue_adriano'] += 1
            st.rerun()

        if st.button("📊 Atualizar média 90d", use_container_width=True):
            try:
                conn_pg = st.connection("banco_erp", type="sql")
                query_media = 'SELECT loja, codigo, qtde_media_semanal FROM "python_90dSEMANA"'
                df_media = conn_pg.query(query_media, ttl=0)
                
                if not df_media.empty:
                    conn.update(worksheet=WS_MEDIA_90D, data=df_media)
                    st.success("✅ Média 90d sincronizada do ERP para o Sheets!")
                    st.cache_data.clear()
                else:
                    st.warning("Nenhum dado retornado da view.")
            except Exception as e:
                st.error(f"⚠️ Erro ao atualizar média 90d do PostgreSQL: {e}")

    # FUNÇÃO MODAL DE CONFIRMAÇÃO PARA ZERAR
    @st.dialog("🚨 Confirmação Necessária")
    def modal_zerar_pedidos():
        st.markdown("Tem certeza que deseja **zerar todos os pedidos** de todas as lojas?")
        st.write("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Não, cancelar", use_container_width=True):
                st.rerun()
        with c2:
            if st.button("✔️ Sim, zerar tudo", type="primary", use_container_width=True):
                st.session_state['reset_counter_acougue_adriano'] += 1
                df_main = carregar_pedidos()
                for loja in LOJAS:
                    if loja in df_main.columns:
                        df_main[loja] = 0
                salvar_pedidos(df_main)
                st.rerun()

    # ROTA 1 — SEPARAÇÃO E FECHAMENTO
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown("""
        <div class="page-header hide-print" style="background: linear-gradient(90deg, var(--red-dark) 0%, #1a0808 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">📊</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Separação e Fechamento — Açougue Final Adriano</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.container(border=True):
            df_base = carregar_pedidos()
            
            if df_base.empty:
                st.warning("A base de pedidos está vazia. Cadastre produtos no Catálogo primeiro.")
                st.stop()
                
            df_base["TOTAL GERAL"] = df_base[LOJAS].sum(axis=1)

            col_cfg = {
                "Fornecedor":  st.column_config.TextColumn("Fornecedor", disabled=True),
                "Código":      st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                "Descrição":   st.column_config.TextColumn("Produto", disabled=True),
                "TOTAL GERAL": st.column_config.NumberColumn("TOTAL ▶️", width=90, format="%d", disabled=True),
            }
            for loja, novo_nome in MAPA_LOJAS.items():
                col_cfg[loja] = st.column_config.NumberColumn(novo_nome, format="%d", min_value=0, step=1)

            cols_order = ["Fornecedor", "Código", "Descrição"] + LOJAS + ["TOTAL GERAL"]
            df_exibir = df_base[cols_order]

            df_editado = st.data_editor(
                df_exibir,
                hide_index=True,
                use_container_width=True,
                height=600,
                column_config=col_cfg,
                key=f"sep_editor_{st.session_state['reset_counter_acougue_adriano']}"
            )

            html_table = df_editado.to_html(index=False, classes=["print-table", "print-sep"])
            st.markdown(f"""<div id="print-section">
            <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
                Resumo de Separação — Açougue Final Adriano
            </h2>
            <div class="print-container">
            {html_table}
            </div>
            </div>""", unsafe_allow_html=True)

            st.divider()
            col_salvar, col_csv, col_excel, col_print, col_zerar = st.columns([2.5, 1.2, 1.2, 1.5, 2.5])

            with col_salvar:
                if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
                    df_to_save = df_editado.drop(columns=["TOTAL GERAL"])
                    salvar_pedidos(df_to_save)
                    st.success("✅ Pedidos salvos na nuvem com sucesso!")
                    st.rerun()

            with col_csv:
                df_csv = df_editado.copy().rename(columns=MAPA_LOJAS)
                csv = df_csv.to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ CSV", data=csv, file_name="separacao_acougue_adriano.csv", mime="text/csv", use_container_width=True)

            with col_excel:
                df_exp = df_editado.copy().rename(columns=MAPA_LOJAS)
                excel_data = gerar_excel_estilizado(df_exp, "Separação Acougue")
                st.download_button("⬇️ Excel", data=excel_data, file_name="separacao_acougue_adriano.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
                                   
            with col_print:
                if st.button("🖨️ Imprimir", use_container_width=True):
                    components.html("<script>window.parent.print();</script>", height=0)

            with col_zerar:
                if st.button("🚨 Zerar Todos os Pedidos", use_container_width=True, key="btn_zerar_sep"):
                    modal_zerar_pedidos()

    # ROTA 2 — VISÃO DAS LOJAS
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS) if acesso_total else usuario_atual

        col_info, col_logout = st.columns([8, 2])
        with col_info:
            id_loja = MAPA_LOJAS.get(loja_selecionada, loja_selecionada)
            st.markdown(f"""
            <div class="topbar-loja hide-print">
                <div class="topbar-left">
                    <span style="font-size:22px">🍖</span>
                    <div>
                        <div class="topbar-title">{loja_selecionada} — Açougue Final Adriano <span class="erp-badge">🟢 Conectado ao ERP</span></div>
                        <div class="topbar-sub">Preencha a quantidade de cada produto</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_logout:
            st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
            if st.button("🏠 Voltar ao Início", use_container_width=True):
                st.session_state['modulo_ativo'] = 'Home'
                st.rerun()

        df_cat = carregar_catalogo_acougue()
        df_cat_loja = df_cat[df_cat[loja_selecionada] == True].copy()
        
        if df_cat_loja.empty:
            st.warning(f"Nenhum produto habilitado para a {loja_selecionada} no momento.")
            st.stop()

        df_all = carregar_pedidos()
        df_loja_view = pd.merge(
            df_cat_loja[["Fornecedor", "Código", "Descrição"]],
            df_all[["Fornecedor", "Código", loja_selecionada]],
            on=["Fornecedor", "Código"],
            how="left"
        )
        df_loja_view[loja_selecionada] = df_loja_view[loja_selecionada].fillna(0).astype(int)
        df_loja_view = df_loja_view.rename(columns={loja_selecionada: "Qtde"})

        mapa_banco_erp = {
            "Loja 01": "001", "Loja 02": "002", "Loja 03": "003", 
            "Loja 04": "004", "Loja 05": "005", "Loja 06": "006", 
            "Loja 07": "007", "Loja 08": "008"
        }
        cod_empresa_banco = mapa_banco_erp.get(loja_selecionada, "001")

        # 1. Puxar Estoque do ERP
        try:
            conn_pg = st.connection("banco_erp", type="sql")
            
            query_erp = f"""
                SELECT cade_codempresa,
                       cade_codigo,
                       cadp_descricao,
                       estoque
                FROM "python_estoque"
                WHERE cade_codempresa::text = '{cod_empresa_banco}'
                ORDER BY cade_codempresa, cade_codigo
            """
            
            df_erp = conn_pg.query(query_erp, ttl=300)
            
            if not df_erp.empty:
                df_erp = df_erp.rename(columns={"cade_codigo": "Código", "estoque": "Estoque"})
                df_loja_view = pd.merge(df_loja_view, df_erp[["Código", "Estoque"]], on="Código", how="left")
            else:
                df_loja_view["Estoque"] = 0
                
        except Exception as e:
            st.error(f"⚠️ Erro ao puxar dados do ERP PostgreSQL: {e}")
            df_loja_view["Estoque"] = 0

        # 2. Puxar Média 90d do Google Sheets
        df_media_all = carregar_media_90d()
        if not df_media_all.empty and "loja" in df_media_all.columns:
            df_media_loja = df_media_all[df_media_all["loja"].astype(str).str.zfill(3) == cod_empresa_banco].copy()
            df_media_loja = df_media_loja.rename(columns={"codigo": "Código", "qtde_media_semanal": "Média 90d"})
            df_media_loja = df_media_loja.drop_duplicates(subset=["Código"])
            df_media_loja["Código"] = pd.to_numeric(df_media_loja["Código"], errors='coerce').fillna(0).astype(int)
        else:
            df_media_loja = pd.DataFrame(columns=["Código", "Média 90d"])

        df_loja_view = pd.merge(df_loja_view, df_media_loja[["Código", "Média 90d"]], on="Código", how="left")

        # Trata Nulos
        df_loja_view["Estoque"] = df_loja_view["Estoque"].fillna(0).astype(int)
        df_loja_view["Média 90d"] = pd.to_numeric(df_loja_view["Média 90d"], errors='coerce').fillna(0.0).round(2)
        df_loja_view["Qtde"] = df_loja_view["Qtde"].fillna(0).astype(int)
        
        # Ordena colunas
        df_loja_view = df_loja_view[["Fornecedor", "Código", "Descrição", "Estoque", "Média 90d", "Qtde"]]

        with st.container(border=True):
            st.info("💡 **Dica:** O **Estoque** foi preenchido automaticamente com base no sistema ERP. Você pode preencher apenas a **Qtde** do pedido.")

            col_cfg_loja = {
                "Fornecedor": st.column_config.TextColumn("Fornecedor", width=150, disabled=True),
                "Código":     st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                "Descrição":  st.column_config.TextColumn("Produto", width=250, disabled=True),
                "Estoque":    st.column_config.NumberColumn("📦 Estoque", width=100, format="%d", disabled=True),
                "Média 90d":  st.column_config.NumberColumn("📊 Média 90d", width=100, format="%.2f", disabled=True),
                "Qtde":       st.column_config.NumberColumn("🛒 Qtde", width=120, min_value=0, step=1),
            }

            with st.container():
                df_editado = st.data_editor(
                    df_loja_view,
                    column_config=col_cfg_loja,
                    hide_index=True,
                    use_container_width=True,
                    height=520,
                    key=f"loja_acougue_adriano_{st.session_state['reset_counter_acougue_adriano']}"
                )

        df_imprimir = df_editado.copy()
        df_imprimir = df_imprimir.rename(columns={"Estoque": "Est.", "Média 90d": "Méd.90d", "Qtde": "Ped."})
        html_table_loja = df_imprimir.to_html(index=False, classes=["print-table", "print-loja"])
        
        st.markdown(f"""<div id="print-section">
        <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
            Resumo do Pedido — {loja_selecionada}
        </h2>
        <div class="print-container">
        {html_table_loja}
        </div>
        </div>""", unsafe_allow_html=True)

        itens_com_pedido = int((df_editado["Qtde"] > 0).sum())
        total_itens      = len(df_editado)
        total_unidades   = int(df_editado["Qtde"].sum())
        pct              = round(itens_com_pedido / total_itens * 100) if total_itens else 0

        st.divider()
        m1, m2, m3, col_print, col_aviso, col_btn = st.columns([1.5, 1.5, 1.5, 1.3, 2.4, 2.8])
        with m1: st.metric("Itens preenchidos", f"{itens_com_pedido} / {total_itens}")
        with m2: st.metric("Total de unidades", total_unidades)
        with m3: st.metric("Cobertura", f"{pct}%")
        
        with col_print:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("🖨️ Imprimir", use_container_width=True):
                components.html("<script>window.parent.print();</script>", height=0)

        with col_aviso:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("🚫 Sem Pedido Hoje", use_container_width=True):
                df_main = carregar_pedidos()
                df_main[loja_selecionada] = 0
                salvar_pedidos(df_main)
                
                msg_aviso = f"🚨 *AVISO - Açougue Final*\nA {loja_selecionada} informou que *NÃO* fará pedido de peças nesta semana."
                enviado = notificar_telegram(msg_aviso)
                
                if enviado:
                    st.success("✅ Supervisor avisado e pedido zerado!")
                else:
                    st.warning("⚠️ O pedido foi zerado, mas falhou ao enviar o Telegram.")

        with col_btn:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("💾 Salvar Pedido da Semana", type="primary", use_container_width=True):
                df_main = carregar_pedidos()
                
                for _, row in df_editado.iterrows():
                    mask = (
                        (df_main["Fornecedor"] == row["Fornecedor"]) &
                        (df_main["Código"] == row["Código"])
                    )
                    if mask.any():
                        df_main.loc[mask, loja_selecionada] = row["Qtde"]
                    else:
                        nova_linha = {"Fornecedor": row["Fornecedor"], "Código": row["Código"], "Descrição": row["Descrição"]}
                        for l in LOJAS: nova_linha[l] = 0
                        nova_linha[loja_selecionada] = row["Qtde"]
                        df_main = pd.concat([df_main, pd.DataFrame([nova_linha])], ignore_index=True)
                
                salvar_pedidos(df_main)
                st.success(f"✅ Pedido da {loja_selecionada} enviado para a nuvem com sucesso!")

    # ROTA 3 — VISÃO POR FORNECEDOR / RESUMO (Admin)
    elif perfil_navegacao == "Visão por Fornecedor (Resumo)":
        fuso_br = pytz.timezone('America/Sao_Paulo')
        agora = datetime.now(fuso_br).strftime('%d/%m/%Y %H:%M')

        st.markdown("""
        <div class="hide-print" style="background: linear-gradient(90deg, var(--red-dark) 0%, #1a0808 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">🍖</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Visão por Fornecedor — Açougue Final Adriano</div>
                <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">Resumo consolidado agrupado pelas categorias/fornecedores de produtos</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="hide-print" style="text-align: right; color: #ffffff; font-size: 15px; margin-top: -10px; margin-bottom: 15px;">
            Gerado em: <strong>{agora}</strong>
        </div>
        """, unsafe_allow_html=True)

        df_all = carregar_pedidos()
        df_cat = carregar_catalogo_acougue()
        
        if df_all.empty or df_cat.empty:
            st.warning("Não há dados de pedidos ou catálogo preenchidos.")
            st.stop()

        nomes_forn = df_cat["Fornecedor"].dropna().unique().tolist()
        
        html_print_content = ""

        for i in range(0, len(nomes_forn), 1):
            cols = st.columns(1, gap="small")
            for j, fornecedor in enumerate(nomes_forn[i:i+1]):

                df_forn = df_all[df_all["Fornecedor"] == fornecedor].copy()
                colunas_presentes = LOJAS
                df_forn = df_forn[["Código", "Descrição"] + colunas_presentes].copy()
                df_forn["TOTAL"] = df_forn[colunas_presentes].sum(axis=1)
                
                lojas_renomeadas = {l: MAPA_LOJAS[l] for l in colunas_presentes}
                df_forn = df_forn.rename(columns=lojas_renomeadas)

                lojas_cols_renomeadas = [MAPA_LOJAS[l] for l in colunas_presentes]

                col_cfg_forn = {
                    "Código":    st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                    "Descrição": st.column_config.TextColumn("Produto", disabled=False),
                    "TOTAL":     st.column_config.NumberColumn("TOTAL", format="%d", disabled=True),
                }
                for c in lojas_cols_renomeadas:
                    col_cfg_forn[c] = st.column_config.NumberColumn(c, format="%d", disabled=False, min_value=0)

                altura = (len(df_forn) * 35) + 42 

                with cols[j]:
                    with st.container(border=True):
                        st.markdown('<div class="title-input">', unsafe_allow_html=True)
                        st.text_input(
                            "Fornecedor",
                            value=f"🍖 {fornecedor}",
                            label_visibility="collapsed",
                            key=f"title_forn_{fornecedor}_{st.session_state['reset_counter_acougue_adriano']}"
                        )
                        st.markdown('</div>', unsafe_allow_html=True)

                        cols_order_forn = ["Código", "Descrição"] + lojas_cols_renomeadas + ["TOTAL"]
                        df_forn_edit = st.data_editor(
                            df_forn[cols_order_forn],
                            hide_index=True,
                            use_container_width=True,
                            column_config=col_cfg_forn,
                            height=altura,
                            num_rows="fixed",
                            key=f"forn_acougue_adriano_{fornecedor}_{st.session_state['reset_counter_acougue_adriano']}"
                        )

                        total_geral = int(df_forn_edit["TOTAL"].sum()) if "TOTAL" in df_forn_edit.columns else 0
                        st.markdown(f"""
                            <div style="text-align:right; font-weight:700; margin-top:6px; color:var(--red-bright); font-size:15px;">
                                Total Geral: {total_geral} unidades
                            </div>
                        """, unsafe_allow_html=True)
                        
                        html_table = df_forn_edit.to_html(index=False, classes=["print-table", "print-forn"])
                        html_print_content += f"<h3 style='color: black; margin-top: 10px; margin-bottom: 4px;'>🍖 {fornecedor}</h3>\n"
                        html_print_content += f"{html_table}\n"
                        html_print_content += f"<div style='text-align:right; font-weight:bold; font-size:11px; margin-top:3px; margin-bottom: 8px; color: black;'>Total do Fornecedor: {total_geral} unidades</div>\n"

        st.write("<br>", unsafe_allow_html=True)

        st.markdown(f"""<div id="print-section">
        <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
            Visão por Fornecedor (Resumo) — Açougue Final Adriano &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span style="font-size: 11px; font-weight: bold; color: #000000;">{agora}</span>
        </h2>
        <div class="print-container">
        {html_print_content}
        </div>
        </div>""", unsafe_allow_html=True)
        
        st.divider()

        col_csv, col_excel, col_print, col_zerar = st.columns([1.5, 1.5, 2, 2.5])

        df_export_forn = df_all[["Fornecedor", "Código", "Descrição"] + LOJAS].copy()
        df_export_forn["TOTAL"] = df_export_forn[LOJAS].sum(axis=1)
        df_export_forn = df_export_forn.rename(columns=MAPA_LOJAS)

        with col_csv:
            csv_data = df_export_forn.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ CSV", data=csv_data, file_name="visao_fornecedores.csv", mime="text/csv", use_container_width=True)

        with col_excel:
            excel_data = gerar_excel_estilizado(df_export_forn, "Visão Fornecedores")
            st.download_button("⬇️ Excel", data=excel_data, file_name="visao_fornecedores.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

        with col_print:
            if st.button("🖨️ Imprimir Resumo Geral", use_container_width=True):
                components.html("<script>window.parent.print();</script>", height=0)

        with col_zerar:
            if st.button("🚨 Zerar Todos os Pedidos", use_container_width=True, key="btn_zerar_forn"):
                modal_zerar_pedidos()

    # ROTA 4 — CATÁLOGO DE PRODUTOS
    elif perfil_navegacao == "Catálogo de Produtos":
        st.markdown("""
        <div class="page-header hide-print" style="background: linear-gradient(90deg, var(--red-dark) 0%, #1a0808 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">🗂️</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Catálogo de Peças e Nomes Oficiais</div>
                <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">Atualize nomes direto do ERP ou crie apelidos personalizados. Os apelidos terão prioridade em todo o sistema.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        df_catalogo = carregar_catalogo_acougue()
        df_editor_input = df_catalogo.drop(columns=["Descrição"], errors="ignore")

        # LIMPEZA DE DADOS OBRIGATÓRIA PARA EVITAR ERRO DE TIPOS
        # 1. Garante que o código seja número
        df_editor_input["Código"] = pd.to_numeric(df_editor_input["Código"], errors='coerce').fillna(0).astype(int)

        # 2. Garante que colunas de texto não tenham nulos (Evita o erro de tipo)
        cols_texto = ["Fornecedor", "Descrição Oficial", "Nome Personalizado"]
        for col in cols_texto:
            if col in df_editor_input.columns:
                df_editor_input[col] = df_editor_input[col].fillna("").astype(str)

        # 3. Garante que as colunas das lojas sejam estritamente Booleanas (Checkbox)
        for loja in LOJAS:
            if loja in df_editor_input.columns:
                df_editor_input[loja] = df_editor_input[loja].fillna(False).astype(bool)

        ordem_colunas = ["Fornecedor", "Código", "Descrição Oficial", "Nome Personalizado"] + LOJAS
        df_editor_input = df_editor_input[ordem_colunas]
        
        with st.container(border=True):
            
            opcoes_forn = ["BOVINOS", "LINGUIÇAS", "COCAMAR", "SUINO", "CONGELADOS - PADRÃO E IQF", "DIVERSOS"]
            
            col_cfg_cat = {
                "Fornecedor":         st.column_config.SelectboxColumn("Fornecedor", options=opcoes_forn, width=150, required=True),
                "Código":             st.column_config.NumberColumn("Cód.", width=80, format="%d", required=True),
                "Descrição Oficial":  st.column_config.TextColumn("Nome Oficial (ERP)", width=280, disabled=False),
                "Nome Personalizado": st.column_config.TextColumn("Nome Personalizado (Apelido)", width=230),
            }
            for loja in LOJAS:
                col_cfg_cat[loja] = st.column_config.CheckboxColumn(loja, default=False)
                
            edited_cat = st.data_editor(
                df_editor_input,
                use_container_width=True,
                hide_index=True,
                height=600,
                num_rows="dynamic",
                column_config=col_cfg_cat,
                key=f"editor_catalogo_acougue_adriano_{st.session_state['reset_counter_acougue_adriano']}"
            )
            
        st.divider()
        col_atualizar, col_sync, _ = st.columns([2.5, 3, 3])
        
        with col_atualizar:
            if st.button("💾 Salvar Catálogo", type="primary", use_container_width=True):
                salvar_catalogo(edited_cat)
                st.session_state['reset_counter_acougue_adriano'] += 1
                st.success("✅ Catálogo atualizado com sucesso!")
                st.rerun()
                
        with col_sync:
            if st.button("📥 Puxar Nomes do ERP", use_container_width=True):
                try:
                    conn_pg = st.connection("banco_erp", type="sql")
                    cods = tuple(edited_cat["Código"].tolist())
                    
                    if len(cods) == 1:
                        cods_str = f"({cods[0]})"
                    else:
                        cods_str = str(cods)
                        
                    query_nomes = f"SELECT cadp_codigo, cadp_descricao FROM cadprod WHERE cadp_codigo IN {cods_str}"
                    df_nomes = conn_pg.query(query_nomes, ttl=0)
                    
                    for _, row in df_nomes.iterrows():
                        mask = edited_cat["Código"] == row["cadp_codigo"]
                        edited_cat.loc[mask, "Descrição Oficial"] = row["cadp_descricao"]
                        
                    salvar_catalogo(edited_cat)
                    st.success("✅ Nomes Oficiais sincronizados com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Erro ao buscar nomes no banco ERP: {e}")
