import datetime
import io
import pandas as pd
import streamlit as st
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import firebase_admin
from firebase_admin import credentials, firestore

st.set_page_config(
    page_title="Gestão de Cargas - Logística",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 98%;}
        .kanban-header {text-align: center; background: #262730; color: #ffffff !important; padding: 10px; border-radius: 8px; font-weight: 600; font-size: 14px; border: 1px solid #464e5f; margin-bottom: 10px;}
    </style>
""",
    unsafe_allow_html=True,
)

# Inicialização com diagnóstico de erro
@st.cache_resource
def conectar_firebase():
    try:
        # Limpa instâncias anteriores para evitar erro de app já existente
        for app in list(firebase_admin._apps.values()):
            firebase_admin.delete_app(app)
            
        if "firebase" not in st.secrets:
            return None
            
        cred_dict = dict(st.secrets["firebase"])
        
        # Corrige quebras de linha na chave privada
        if "private_key" in cred_dict:
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")

        cred = credentials.Certificate(cred_dict)
        app = firebase_admin.initialize_app(cred)
        return firestore.client(app=app)
    except Exception as e:
        st.error(f"Erro na conexão Firebase: {e}")
        return None

db = conectar_firebase()

# Funções auxiliares
def carregar_dados(colecao):
    if db:
        try:
            return [doc.to_dict() for doc in db.collection(colecao).stream()]
        except Exception:
            return []
    return []

def salvar_item_firebase(colecao, doc_id, dados):
    if db:
        db.collection(colecao).document(str(doc_id)).set(dados)

def excluir_item_firebase(colecao, doc_id):
    if db:
        db.collection(colecao).document(str(doc_id)).delete()

# Interface Principal
st.title("🚚 Painel de Controle de Cargas")

if not db:
    st.error("⚠️ Firebase não conectado. Verifique se a seção [firebase] existe nos seus Secrets no Streamlit Cloud.")

menu = st.radio("Menu Principal", ["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros", "📊 Relatório"], horizontal=True)

# Lógica básica para evitar erro de listagem vazia
motoristas = carregar_dados("motoristas")
cargas = carregar_dados("cargas")

if menu == "📋 Painel (Kanban)":
    st.subheader("Visão Geral das Cargas")
    # Aqui entra o seu código de exibição do Kanban que você já possui
    if not cargas:
        st.info("Nenhuma carga encontrada no banco de dados.")

elif menu == "➕ Nova Carga":
    with st.form("nova_carga"):
        m = st.selectbox("Motorista", [mot.get("nome") for mot in motoristas])
        d = st.text_input("Destino")
        if st.form_submit_button("Salvar"):
            salvar_item_firebase("cargas", str(datetime.datetime.now().timestamp()), {"motorista": m, "destino": d, "status": "Aguardando Carregamento"})
            st.rerun()
