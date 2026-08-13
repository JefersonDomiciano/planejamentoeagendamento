import datetime
import os
import json
import pandas as pd
import streamlit as st
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

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

# Inicialização isolada do Firebase
@st.cache_resource
def obter_conexao():
    try:
        if not firebase_admin._apps:
            if "FIREBASE_CREDENTIALS_JSON" in os.environ:
                cred_dict = json.loads(os.environ.get("FIREBASE_CREDENTIALS_JSON"))
                cred = credentials.Certificate(cred_dict)
            else:
                cred = credentials.Certificate("serviceAccountKey.json")
            firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception:
        return None

# Funções protegidas com timeout simulado / tratamento seco para não travar
def buscar_colecao(colecao):
    db = obter_conexao()
    if db is None:
        return []
    try:
        docs = db.collection(colecao).get()
        return [doc.to_dict() for doc in docs]
    except Exception:
        return []

def salvar_no_firebase(colecao, dados, doc_id):
    db = obter_conexao()
    if db is None:
        st.error("Erro: Banco de dados desconectado.")
        return
    try:
        db.collection(colecao).document(str(doc_id)).set(dados)
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

# Inicializa dados vazios na sessão para carregar instantaneamente sem bloqueio
if "cargas" not in st.session_state:
    st.session_state["cargas"] = []
if "motoristas" not in st.session_state:
    st.session_state["motoristas"] = []
if "ajudantes" not in st.session_state:
    st.session_state["ajudantes"] = []

col_top1, col_top2 = st.columns([0.8, 0.2])
with col_top2:
    if st.button("🔄 Sincronizar Dados"):
        with st.spinner("Buscando dados..."):
            st.session_state["cargas"] = buscar_colecao("cargas")
            st.session_state["motoristas"] = buscar_colecao("motoristas")
            st.session_state["ajudantes"] = buscar_colecao("ajudantes")
        st.success("Sincronizado!")
        st.rerun()

cargas_lista = st.session_state["cargas"]
motoristas_lista = [m.get("nome", "") for m in st.session_state["motoristas"]]
ajudantes_lista = [a.get("nome", "") for a in st.session_state["ajudantes"]]

# Abas do Painel
tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros (Equipe)", "📊 Relatório Semanal"])

with tab1:
    st.subheader("Visão Geral das Cargas")
    if not cargas_lista:
        st.info("Clique no botão '🔄 Sincronizar Dados' no canto superior direito para carregar as informações.")
    else:
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        
        status_map = {
            "Pendente": col_k1,
            "Em Rota": col_k2,
            "Entregue": col_k3,
            "Cancelado": col_k4
        }
        
        with col_k1: st.markdown('<div class="kanban-header">📌 Pendente</div>', unsafe_allow_html=True)
        with col_k2: st.markdown('<div class="kanban-header">🚚 Em Rota</div>', unsafe_allow_html=True)
        with col_k3: st.markdown('<div class="kanban-header">✅ Entregue</div>', unsafe_allow_html=True)
        with col_k4: st.markdown('<div class="kanban-header">❌ Cancelado</div>', unsafe_allow_html=True)
        
        for carga in cargas_lista:
            status_atual = carga.get("status", "Pendente")
            col_alvo = status_map.get(status_atual, col_k1)
            
            with col_alvo:
                with st.container():
                    st.markdown(f"**ID:** {carga.get('id', 'N/A')}")
                    st.markdown(f"**Motorista:** {carga.get('motorista', 'Não informado')}")
                    st.markdown(f"**Destino:** {carga.get('destino', 'Não informado')}")
                    st.markdown(f"**Saída:** {carga.get('data_saida', '')}")
                    
                    novo_status = st.selectbox(
                        "Status", 
                        ["Pendente", "Em Rota", "Entregue", "Cancelado"], 
                        index=["Pendente", "Em Rota", "Entregue", "Cancelado"].index(status_atual) if status_atual in ["Pendente", "Em Rota", "Entregue", "Cancelado"] else 0,
                        key=f"status_{carga.get('id')}"
                    )
                    
                    if novo_status != status_atual:
                        carga["status"] = novo_status
                        salvar_no_firebase("cargas", carga, carga.get("id"))
                        st.rerun()

with tab2:
    st.subheader("Cadastrar Nova Carga")
    with st.form("form_nova_carga"):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            carga_id = st.text_input("ID / Número do Pedido ou Carga")
            motorista = st.selectbox("Motorista", motoristas_lista if motoristas_lista else ["Nenhum cadastrado (Sincronize)"])
            destino = st.text_input("Destino (Cidade / Bairro)")
        with col_f2:
            ajudantes_selecionados = st.multiselect("Ajudantes", ajudantes_lista)
            data_saida = st.date_input("Data de Saída", datetime.date.today())
            data_entrega = st.date_input("Previsão de Entrega", datetime.date.today())
            
        submitted = st.form_submit_button("Salvar Carga no Firebase")
        if submitted:
            if not carga_id:
                st.error("O ID da carga é obrigatório!")
            else:
                dados_carga = {
                    "id": carga_id,
                    "motorista": motorista,
                    "destino": destino,
                    "ajudantes": ajudantes_selecionados,
                    "data_saida": str(data_saida),
                    "data_entrega": str(data_entrega),
                    "status": "Pendente"
                }
                salvar_no_firebase("cargas", dados_carga, carga_id)
                st.success(f"Carga {carga_id} salva com sucesso!")

with tab3:
    st.subheader("Gerenciamento de Cadastros (Equipe)")
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("### Cadastrar Motorista")
        with st.form("form_motorista"):
            nome_mot = st.text_input("Nome do Motorista")
            if st.form_submit_button("Adicionar Motorista"):
                if nome_mot:
                    salvar_no_firebase("motoristas", {"nome": nome_mot}, nome_mot)
                    st.success(f"Motorista {nome_mot} adicionado!")
        
        st.markdown("#### Motoristas Carregados:")
        for m in motoristas_lista:
            st.text(f"• {m}")
            
    with col_c2:
        st.markdown("### Cadastrar Ajudante")
        with st.form("form_ajudante"):
            nome_aju = st.text_input("Nome do Ajudante")
            if st.form_submit_button("Adicionar Ajudante"):
                if nome_aju:
                    salvar_no_firebase("ajudantes", {"nome": nome_aju}, nome_aju)
                    st.success(f"Ajudante {nome_aju} adicionado!")
                    
        st.markdown("#### Ajudantes Carregados:")
        for a in ajudantes_lista:
            st.text(f"• {a}")

with tab4:
    st.subheader("Relatório Semanal e Exportação")
    if not cargas_lista:
        st.info("Sincronize os dados para gerar o relatório.")
    else:
        df_relatorio = pd.DataFrame(cargas_lista)
        st.dataframe(df_relatorio, use_container_width=True)
        st.download_button(
            label="Baixar Dados em CSV",
            data=df_relatorio.to_csv(index=False).encode('utf-8'),
            file_name="relatorio_cargas.csv",
            mime="text/csv",
        )
