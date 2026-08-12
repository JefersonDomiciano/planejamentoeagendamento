import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime

# --- CONFIGURAÇÃO DA CONEXÃO COM GOOGLE SHEETS ---
@st.cache_resource
def get_sheet():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Dicionário completo e estruturado exigido pelo Google Auth
    creds_info = {
        "type": "service_account",
        "project_id": "logistica-app",
        "private_key_id": "private_key",
        "private_key": st.secrets["gcp_private_key"].replace("\\n", "\n"),
        "client_email": st.secrets["gcp_service_account_email"],
        "client_id": "123456789",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/" + st.secrets["gcp_service_account_email"].replace("@", "%40")
    }
    
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
    client = gspread.authorize(creds)
    return client.open("Banco_Logistica").sheet1

# --- FUNÇÕES DE DADOS ---
def carregar_dados():
    sheet = get_sheet()
    dados = sheet.get_all_records()
    return pd.DataFrame(dados) if dados else pd.DataFrame(columns=['id', 'motorista', 'destino', 'observacoes', 'ajudantes', 'data_carga', 'data_saida', 'data_entrega', 'status'])

def salvar_nova_carga(carga_dict):
    sheet = get_sheet()
    sheet.append_row([
        carga_dict['id'], carga_dict['motorista'], carga_dict['destino'], 
        carga_dict['observacoes'], str(carga_dict['ajudantes']), 
        carga_dict['data_carga'], carga_dict['data_saida'], 
        carga_dict['data_entrega'], carga_dict['status']
    ])

# --- INTERFACE ---
st.set_page_config(page_title="Painel de Controle de Cargas", layout="wide")
st.title("🚚 Painel de Controle de Cargas - Equipe")

menu = st.radio("Menu", ["📋 Painel (Kanban)", "➕ Nova Carga"], horizontal=True)

df_cargas = carregar_dados()

if menu == "📋 Painel (Kanban)":
    st.subheader("Cargas em Aberto")
    st.dataframe(df_cargas, use_container_width=True)

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
        
        if st.form_submit_button("Salvar na Planilha"):
            novo_id = len(df_cargas) + 1
            nova_carga = {
                "id": novo_id, "motorista": mot, "destino": dest,
                "observacoes": "", "ajudantes": "",
                "data_carga": str(datetime.date.today()),
                "data_saida": str(saida), "data_entrega": "",
                "status": status
            }
            salvar_nova_carga(nova_carga)
            st.success("Carga salva com sucesso no Google Sheets!")
            st.rerun()
