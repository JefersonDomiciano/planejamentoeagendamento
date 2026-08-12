import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

# --- CONFIGURAÇÃO DA CONEXÃO COM GOOGLE SHEETS ---
@st.cache_resource
def get_sheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp"]), scope)
    client = gspread.authorize(creds)
    # Abre a planilha pelo nome exato no seu Drive
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
