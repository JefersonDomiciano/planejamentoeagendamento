import datetime
import json
import os
import pandas as pd
import streamlit as st

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

ARQUIVO_DADOS = "dados_logistica.json"

def carregar_dados_locais():
    if os.path.exists(ARQUIVO_DADOS):
        try:
            with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"cargas": [], "motoristas": [], "ajudantes": []}

def salvar_dados_locais(dados):
    try:
        with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"Erro ao salvar arquivo local: {e}")

if "db_local" not in st.session_state:
    st.session_state["db_local"] = carregar_dados_locais()

db = st.session_state["db_local"]

cargas_lista = db.get("cargas", [])
motoristas_lista = [m.get("nome", "") for m in db.get("motoristas", [])]
ajudantes_lista = [a.get("nome", "") for a in db.get("ajudantes", [])]

tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros (Equipe)", "📊 Relatório Semanal"])

with tab1:
    st.subheader("Visão Geral das Cargas")
    if not cargas_lista:
        st.info("Nenhuma carga cadastrada. Utilize a aba 'Nova Carga' para começar.")
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
        
        for idx, carga in enumerate(cargas_lista):
            status_atual = carga.get("status", "Pendente")
            col_alvo = status_map.get(status_atual, col_k1)
            
            with col_alvo:
                with st.container():
                    st.markdown(f"**Destino:** {carga.get('destino', 'Não informado')}")
                    st.markdown(f"**Motorista:** {carga.get('motorista', 'Não informado')}")
                    st.markdown(f"**Carregamento:** {carga.get('data_carregamento', '')}")
                    st.markdown(f"**Saída:** {carga.get('data_saida', '')}")
                    
                    novo_status = st.selectbox(
                        "Status", 
                        ["Pendente", "Em Rota", "Entregue", "Cancelado"], 
                        index=["Pendente", "Em Rota", "Entregue", "Cancelado"].index(status_atual) if status_atual in ["Pendente", "Em Rota", "Entregue", "Cancelado"] else 0,
                        key=f"status_{idx}"
                    )
                    
                    if novo_status != status_atual:
                        cargas_lista[idx]["status"] = novo_status
                        db["cargas"] = cargas_lista
                        salvar_dados_locais(db)
                        st.rerun()

with tab2:
    st.subheader("Cadastrar Nova Carga")
    with st.form("form_nova_carga", clear_on_submit=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            destino = st.text_input("Destino (Cidade / Bairro)")
            motorista = st.selectbox("Motorista", motoristas_lista if motoristas_lista else ["Nenhum cadastrado"])
            ajudantes_selecionados = st.multiselect("Ajudantes", ajudantes_lista)
        with col_f2:
            data_carregamento = st.date_input("Data do Carregamento", datetime.date.today())
            data_saida = st.date_input("Data de Saída", datetime.date.today())
            data_entrega = st.date_input("Previsão de Entrega", datetime.date.today())
            
        submitted = st.form_submit_button("Salvar Carga")
        if submitted:
            if not destino:
                st.error("O destino é obrigatório!")
            else:
                dados_carga = {
                    "id": str(len(cargas_lista) + 1),
                    "destino": destino,
                    "motorista": motorista,
                    "ajudantes": ajudantes_selecionados,
                    "data_carregamento": str(data_carregamento),
                    "data_saida": str(data_saida),
                    "data_entrega": str(data_entrega),
                    "status": "Pendente"
                }
                db["cargas"].append(dados_carga)
                salvar_dados_locais(db)
                st.success("Carga salva com sucesso!")
                st.rerun()

with tab3:
    st.subheader("Gerenciamento de Cadastros (Equipe)")
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("### Cadastrar Motorista")
        with st.form("form_motorista", clear_on_submit=True):
            nome_mot = st.text_input("Nome do Motorista")
            submitted_mot = st.form_submit_button("Adicionar Motorista")
            if submitted_mot:
                if nome_mot:
                    db["motoristas"].append({"nome": nome_mot})
                    salvar_dados_locais(db)
                    st.success(f"Motorista {nome_mot} adicionado!")
                    st.rerun()
        
        st.markdown("#### Motoristas Cadastrados:")
        for m in motoristas_lista:
            st.text(f"• {m}")
            
    with col_c2:
        st.markdown("### Cadastrar Ajudante")
        with st.form("form_ajudante", clear_on_submit=True):
            nome_aju = st.text_input("Nome do Ajudante")
            submitted_aju = st.form_submit_button("Adicionar Ajudante")
            if submitted_aju:
                if nome_aju:
                    db["ajudantes"].append({"nome": nome_aju})
                    salvar_dados_locais(db)
                    st.success(f"Ajudante {nome_aju} adicionado!")
                    st.rerun()
                    
        st.markdown("#### Ajudantes Cadastrados:")
        for a in ajudantes_lista:
            st.text(f"• {a}")

with tab4:
    st.subheader("Relatório Semanal e Exportação")
    if not cargas_lista:
        st.info("Sem dados para gerar relatório.")
    else:
        df_relatorio = pd.DataFrame(cargas_lista)
        st.dataframe(df_relatorio, use_container_width=True)
        st.download_button(
            label="Baixar Dados em CSV",
            data=df_relatorio.to_csv(index=False).encode('utf-8'),
            file_name="relatorio_cargas.csv",
            mime="text/csv",
        )
