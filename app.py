import streamlit as st
import pandas as pd
import datetime

# --- CONFIGURAÇÃO DA PLANILHA PÚBLICA ---
# Substitua abaixo pelo ID real da sua planilha (aquela letra miúda que fica na URL entre /d/ e /edit)
PLANILHA_ID = "COLE_O_ID_DA_ SUA_PLANILHA_AQUI"
url = f"https://docs.google.com/spreadsheets/d/{PLANILHA_ID}/gviz/tq?tqx=out:csv"

# --- FUNÇÕES DE DADOS ---
@st.cache_data(ttl=10)
def carregar_dados():
    try:
        df = pd.read_csv(url)
        return df
    except Exception:
        # Se a planilha estiver vazia, retorna a estrutura padrão
        return pd.DataFrame(columns=['id', 'motorista', 'destino', 'observacoes', 'ajudantes', 'data_carga', 'data_saida', 'data_entrega', 'status'])

# --- INTERFACE ---
st.set_page_config(page_title="Painel de Controle de Cargas", layout="wide")
st.title("🚚 Painel de Controle de Cargas - Equipe")

menu = st.radio("Menu", ["📋 Painel (Kanban)", "➕ Nova Carga"], horizontal=True)

df_cargas = carregar_dados()

if menu == "📋 Painel (Kanban)":
    st.subheader("Cargas em Aberto")
    if not df_cargas.empty:
        st.dataframe(df_cargas, use_container_width=True)
    else:
        st.info("Nenhuma carga cadastrada ainda na planilha.")

elif menu == "➕ Nova Carga":
    st.subheader("Cadastrar Nova Carga")
    with st.form("form_carga"):
        col1, col2 = st.columns(2)
        with col1:
            mot = st.text_input("Motorista")
            dest = st.text_input("Destino")
        with col2:
            saida = st.date_input("Data Saída")
            status = st.selectbox("Status", ["Aguardando Carregamento", "Em Trânsito", "Entregue"])
        
        if st.form_submit_button("Salvar Informações"):
            st.success("Para registrar novos dados permanentemente com salvamento automático na planilha, preencha direto na sua planilha do Google Sheets aberta no navegador!")
