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

        div[data-testid="stHorizontalBlock"] button {
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
            font-size: 16px !important;
            padding: 0px !important;
            min-height: unset !important;
        }
        div[data-testid="stHorizontalBlock"] button:hover {
            background-color: rgba(255, 255, 255, 0.05) !important;
            border: none !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- CONFIGURAÇÃO DA CONEXÃO FIREBASE (ATUALIZADA) ---
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        # Carrega o arquivo serviceAccountKey.json que agora existe no GitHub
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()

try:
    db = init_firebase()
    usar_firebase = True
except Exception as e:
    usar_firebase = False
    if "motoristas" not in st.session_state:
        st.session_state.motoristas = ["Carlos Silva", "João Pereira", "Maurício", "Cícero Taveira"]
    if "ajudantes" not in st.session_state:
        st.session_state.ajudantes = ["Pedrinho", "Lucas Souza"]
    if "cargas" not in st.session_state:
        st.session_state.cargas = []

# --- FUNÇÕES DE DADOS (MANTIDAS) ---
def carregar_dados(colecao):
    if usar_firebase:
        docs = db.collection(colecao).stream()
        return [doc.to_dict() for doc in docs]
    else:
        return st.session_state.get(colecao, [])

def salvar_dado(colecao, dados, doc_id):
    if usar_firebase:
        db.collection(colecao).document(str(doc_id)).set(dados)
    else:
        if colecao == "cargas":
            for i, c in enumerate(st.session_state.cargas):
                if c.get("id") == doc_id:
                    st.session_state.cargas[i] = dados
                    break

def adicionar_dado(colecao, dados, doc_id=None):
    if usar_firebase:
        if doc_id:
            db.collection(colecao).document(str(doc_id)).set(dados)
        else:
            db.collection(colecao).add(dados)
    else:
        st.session_state[colecao].append(dados)

def excluir_dado(colecao, campo_filtro, valor):
    if usar_firebase:
        docs = db.collection(colecao).where(campo_filtro, "==", valor).stream()
        for doc in docs:
            doc.reference.delete()
    else:
        if colecao == "cargas":
            st.session_state.cargas = [c for c in st.session_state.cargas if c["id"] != valor]
        else:
            st.session_state.motoristas = [m for m in st.session_state.motoristas if m != valor]

def formatar_data_br(data_str):
    if not data_str: return ""
    try: return datetime.date.fromisoformat(str(data_str)).strftime('%d/%m/%Y')
    except Exception: return str(data_str)

def preparar_dataframe(cargas_lista):
    df = pd.DataFrame(cargas_lista)
    if df.empty: return df
    if "ajudantes" in df.columns: df["ajudantes"] = df["ajudantes"].apply(lambda x: ", ".join(x) if isinstance(x, list) else str(x))
    for col in ["data_carga", "data_saida", "data_entrega"]:
        if col in df.columns: df[col] = df[col].apply(formatar_data_br)
    return df

# ... (Mantenha as funções gerar_excel_profissional e gerar_pdf originais aqui) ...

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

if not usar_firebase:
    st.warning("⚠️ Atenção: Rodando em modo local (Firebase não carregado).")

# ... (Mantenha todo o restante da lógica de Menu, Kanban e Cadastros original aqui) ...
