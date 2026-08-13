import datetime
import io
import pandas as pd
import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
        .kanban-header {
            text-align: center; background: linear-gradient(135deg, #21262d 0%, #161b22 100%);
            color: #ffffff !important; padding: 10px; border-radius: 8px; font-weight: 600;
            font-size: 14px; border: 1px solid #30363d; margin-bottom: 10px;
        }
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] {
            background-color: #161b22 !important; border: 1px solid #30363d !important;
            border-radius: 8px !important; padding: 12px 14px !important; margin-bottom: 10px !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        try:
            cred = credentials.Certificate("serviceAccountKey.json")
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Erro ao inicializar Firebase (Verifique a chave JSON): {e}")
            return None
    return firestore.client()

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

db = init_firebase()
usar_firebase = db is not None

if not usar_firebase:
    st.error("⚠️ Firebase não conectado. Verifique a chave de acesso (serviceAccountKey.json).")
    # Inicia estado local vazio se falhar
    if "cargas" not in st.session_state: st.session_state.cargas = []
    if "motoristas" not in st.session_state: st.session_state.motoristas = []
    if "ajudantes" not in st.session_state: st.session_state.ajudantes = []

def carregar_dados(colecao):
    if usar_firebase:
        try:
            docs = db.collection(colecao).get()
            return [doc.to_dict() for doc in docs]
        except Exception as e:
            st.warning(f"Erro ao buscar '{colecao}': {e}")
            return []
    return st.session_state.get(colecao, [])

def salvar_dado(colecao, dados, doc_id):
    if usar_firebase:
        try:
            db.collection(colecao).document(str(doc_id)).set(dados)
        except Exception as e:
            st.error(f"Erro ao salvar no Firebase: {e}")

def adicionar_dado(colecao, dados, doc_id=None):
    if usar_firebase:
        try:
            if doc_id: db.collection(colecao).document(str(doc_id)).set(dados)
            else: db.collection(colecao).add(dados)
        except Exception as e:
            st.error(f"Erro ao adicionar: {e}")

# Funções auxiliares mantidas para funcionamento do painel
def formatar_data_br(data_str):
    if not data_str: return ""
    try: return datetime.date.fromisoformat(str(data_str)).strftime('%d/%m/%Y')
    except: return str(data_str)

# Carregamento dos dados
motoristas_raw = carregar_dados("motoristas")
ajudantes_raw = carregar_dados("ajudantes")
cargas_lista = carregar_dados("cargas")

motoristas_lista = [m.get("nome", "") for m in motoristas_raw]
ajudantes_lista = [a.get("nome", "") for a in ajudantes_raw]

# Criação das abas
tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros", "📊 Relatório"])

with tab1:
    st.subheader("Visão Geral das Cargas")
    # Lógica Kanban omitida para brevidade, insira a lógica anterior aqui
    st.write(f"Total de cargas carregadas: {len(cargas_lista)}")

with tab2:
    st.subheader("Cadastrar Nova Carga")
    # Formulário...

with tab3:
    st.subheader("Gerenciamento de Equipe")
    # Listas...

with tab4:
    st.subheader("Relatório")
    # Exportação...
