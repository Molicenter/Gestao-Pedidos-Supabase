import streamlit as st
import pandas as pd
from supabase import create_client, Client

# Inicializa o cliente do Supabase
def obter_cliente() -> Client:
    return create_client(st.secrets["https://eemfqkxshqqlhdyofgoz.supabase.co"], st.secrets["eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVlbWZxa3hzaHFxbGhkeW9mZ296Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA3ODcwMjUsImV4cCI6MjA5NjM2MzAyNX0.u7F3h4gdJ0JRY8oNe9qfAXNPwL0AWUPEVt0ZWpEC6SY"])

def iniciar_tela(setor: str):
    supabase = obter_cliente()
    loja_logada = st.session_state.get('usuario_logado', 'Loja 01')
    
    # Extrai apenas o número da loja logada (Ex: "Loja 03" vira 3)
    try:
        num_loja = int(''.join(filter(str.isdigit, loja_logada)))
    except:
        num_loja = 1

    st.markdown(f"## 📦 Pedidos de {setor} — Visão: {loja_logada}")
    st.caption(f"Buscando catálogo e parâmetros de vendas em tempo real no Supabase...")

    # --- 1. BUSCA PRODUTOS DO SETOR QUE ESTÃO ATIVOS PARA A LOJA LOGADA ---
    try:
        # Puxa os produtos do setor ativo
        resp_produtos = supabase.table("produtos").select("*").eq("setor", setor).eq("ativo", True).execute()
        df_produtos = pd.DataFrame(resp_produtos.data)
        
        # Puxa as regras de disponibilidade das lojas
        resp_lojas = supabase.table("produtos_lojas").select("*").eq("loja", num_loja).execute()
        df_lojas = pd.DataFrame(resp_lojas.data)
        
        if not df_produtos.empty:
            # Filtra apenas os produtos que a loja atual tem permissão para pedir
            if not df_lojas.empty:
                codigos_liberados = df_lojas[df_lojas['disponivel'] == True]['codigo_produto'].tolist()
                df_filtrado = df_produtos[df_produtos['codigo'].isin(codigos_liberados)]
            else:
                df_filtrado = df_produtos.copy() # Se não houver restrição, mostra tudo
                
            st.write(f"✅ {len(df_filtrado)} produtos localizados para este setor.")
            
            # --- RENDERIZAÇÃO DO FORMULÁRIO DE PEDIDO ---
            # Aqui vai entrar a sua tabela de digitação de quantidades (vamos construir no próximo passo)
            st.dataframe(df_filtrado[['codigo', 'descricao', 'fornecedor']], use_container_width=True, hide_index=True)
            
        else:
            st.warning(f"Nenhum produto cadastrado para o setor '{setor}' no Supabase.")
            
    except Exception as e:
        st.error(f"Erro ao conectar com o catálogo do Supabase: {e}")
