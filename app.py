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
            text-align: center; 
            background: linear-gradient(135deg, #21262d 0%, #161b22 100%);
            color: #ffffff !important;
            padding: 10px; 
            border-radius: 8px;
            font-weight: 600;
            font-size: 14px;
            border: 1px solid #30363d;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
        }

        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
            padding: 12px 14px !important;
            margin-bottom: 10px !important;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
        }

        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] p, 
        div[data-testid="stVerticalBlock"] div[data-testid="stContainer"] span {
            color: #c9d1d9 !important;
            font-size: 13px !important;
            margin-bottom: 2px !important;
        }

        .stSelectbox label, .stDateInput label {
            font-size: 12px !important;
            color: #8b949e !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

try:
    db = init_firebase()
    usar_firebase = True
    st.success("✅ Conectado ao Firebase com sucesso!")
except Exception:
    usar_firebase = False

def carregar_dados(colecao):
    if usar_firebase:
        docs = db.collection(colecao).get()
        return [doc.to_dict() for doc in docs]
    return []

def salvar_dado(colecao, dados, doc_id):
    if usar_firebase:
        db.collection(colecao).document(str(doc_id)).set(dados)

def adicionar_dado(colecao, dados, doc_id=None):
    if usar_firebase:
        if doc_id:
            db.collection(colecao).document(str(doc_id)).set(dados)
        else:
            db.collection(colecao).add(dados)

def formatar_data_br(data_str):
    if not data_str: return ""
    try: return datetime.date.fromisoformat(str(data_str)).strftime('%d/%m/%Y')
    except: return str(data_str)

def preparar_dataframe(cargas_lista):
    df = pd.DataFrame(cargas_lista)
    if df.empty: return df
    if "ajudantes" in df.columns: df["ajudantes"] = df["ajudantes"].apply(lambda x: ", ".join(x) if isinstance(x, list) else str(x))
    for col in ["data_carga", "data_saida", "data_entrega"]:
        if col in df.columns: df[col] = df[col].apply(formatar_data_br)
    return df[["id", "motorista", "destino", "ajudantes", "data_saida", "data_entrega", "status"]]

# Carregamento inicial
motoristas_raw = carregar_dados("motoristas")
ajudantes_raw = carregar_dados("ajudantes")
cargas_lista = carregar_dados("cargas")

motoristas_lista = [m.get("nome", "") for m in motoristas_raw]
ajudantes_lista = [a.get("nome", "") for a in ajudantes_raw]

tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros (Equipe)", "📊 Relatório Semanal"])

# [O restante da estrutura do painel permanece igual...]
with tab1:
    # Lógica do Kanban aqui...
    st.write("Visão Geral Carregada")
