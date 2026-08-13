import datetime
import json
import os
import requests
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
            font-size: 13px;
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

# Configuração do Firebase Firestore via REST API (Funciona perfeitamente no Render)
FIREBASE_PROJECT_ID = "planejamentoagendamento" # Substitua pelo ID do seu projeto no Firebase se for diferente

def formatar_data_br(data_str):
    if not data_str:
        return ""
    try:
        dt = datetime.datetime.strptime(str(data_str), "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return data_str

def carregar_colecao(colecao):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            documentos = []
            for doc in data.get("documents", []):
                doc_id = doc["name"].split("/")[-1]
                fields = doc.get("fields", {})
                item = {"id": doc_id}
                for k, v in fields.items():
                    item[k] = list(v.values())[0]
                documentos.append(item)
            return documentos
    except Exception:
        pass
    return []

def salvar_documento(colecao, doc_id, dados):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}?documentId={doc_id}"
    fields = {}
    for k, v in dados.items():
        if isinstance(v, list):
            fields[k] = {"arrayValue": {"values": [{"stringValue": str(x)} for x in v]}}
        else:
            fields[k] = {"stringValue": str(v)}
    try:
        requests.patch(f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}/{doc_id}", json={"fields": fields}, timeout=5)
    except Exception as e:
        st.error(f"Erro ao salvar no Firebase: {e}")

def deletar_documento(colecao, doc_id):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}/{doc_id}"
    try:
        requests.delete(url, timeout=5)
    except Exception as e:
        st.error(f"Erro ao excluir no Firebase: {e}")

# Inicializando Estado com Firebase
if "cargas" not in st.session_state:
    st.session_state["cargas"] = carregar_colecao("cargas")
if "motoristas" not in st.session_state:
    st.session_state["motoristas"] = carregar_colecao("motoristas")
if "ajudantes" not in st.session_state:
    st.session_state["ajudantes"] = carregar_colecao("ajudantes")

cargas_lista = st.session_state["cargas"]
motoristas_lista = [m.get("nome", "") for m in st.session_state["motoristas"] if m.get("nome")]
ajudantes_lista = [a.get("nome", "") for a in st.session_state["ajudantes"] if a.get("nome")]

if "editando_idx" not in st.session_state:
    st.session_state["editando_idx"] = None

tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros (Equipe)", "📊 Relatório Semanal"])

with tab1:
    st.subheader("Visão Geral das Cargas")
    
    if st.session_state["editando_idx"] is not None:
        idx_edit = st.session_state["editando_idx"]
        if idx_edit < len(cargas_lista):
            carga_edit = cargas_lista[idx_edit]
            st.markdown("---")
            st.info(f"✏️ Editando Carga para: **{carga_edit.get('destino')}**")
            
            with st.form("form_editar_carga"):
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    novo_destino = st.text_input("Destino", value=carga_edit.get("destino", ""))
                    idx_mot = motoristas_lista.index(carga_edit.get("motorista")) if carga_edit.get("motorista") in motoristas_lista else 0
                    novo_motorista = st.selectbox("Motorista", motoristas_lista if motoristas_lista else ["Nenhum cadastrado"], index=idx_mot)
                    
                    ajudantes_atuais = carga_edit.get("ajudantes", [])
                    if isinstance(ajudantes_atuais, str):
                        ajudantes_atuais = [a.strip() for a in ajudantes_atuais.split(",") if a.strip()]
                    novos_ajudantes = st.multiselect("Ajudantes", ajudantes_lista, default=[a for a in ajudantes_atuais if a in ajudantes_lista])
                
                with col_e2:
                    dt_car_val = datetime.datetime.strptime(carga_edit.get("data_carregamento"), "%Y-%m-%d").date() if carga_edit.get("data_carregamento") else None
                    dt_said_val = datetime.datetime.strptime(carga_edit.get("data_saida"), "%Y-%m-%d").date() if carga_edit.get("data_saida") else None
                    
                    nova_data_carregamento = st.date_input("Previsão / Data de Carregamento", value=dt_car_val if dt_car_val else datetime.date.today())
                    usar_carregamento = st.checkbox("Incluir Data de Carregamento", value=True if carga_edit.get("data_carregamento") else False)

                    nova_data_saida = st.date_input("Data de Saída", value=dt_said_val if dt_said_val else datetime.date.today())
                    usar_saida = st.checkbox("Incluir Data de Saída", value=True if carga_edit.get("data_saida") else False)
                    
                    status_opcoes = ["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"]
                    st_atual = carga_edit.get("status", "Previsão de Carregamento")
                    idx_st = status_opcoes.index(st_atual) if st_atual in status_opcoes else 0
                    novo_status = st.selectbox("Status", status_opcoes, index=idx_st)

                col_salvar, col_cancelar = st.columns(2)
                with col_salvar:
                    salvar_edicao = st.form_submit_button("💾 Salvar Alterações")
                with col_cancelar:
                    cancelar_edicao = st.form_submit_button("❌ Cancelar")

                if salvar_edicao:
                    doc_id = carga_edit.get("id")
                    dados_atualizados = {
                        "id": doc_id,
                        "destino": novo_destino,
                        "motorista": novo_motorista,
                        "ajudantes": novos_ajudantes,
                        "data_carregamento": str(nova_data_carregamento) if usar_carregamento else "",
                        "data_saida": str(nova_data_saida) if usar_saida else "",
                        "status": novo_status
                    }
                    salvar_documento("cargas", doc_id, dados_atualizados)
                    cargas_lista[idx_edit] = dados_atualizados
                    st.session_state["cargas"] = cargas_lista
                    st.session_state["editando_idx"] = None
                    st.success("Carga atualizada com sucesso!")
                    st.rerun()

                if cancelar_edicao:
                    st.session_state["editando_idx"] = None
                    st.rerun()
            st.markdown("---")

    if not cargas_lista:
        st.info("Nenhuma carga cadastrada.")
    else:
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        
        status_map = {
            "Previsão de Carregamento": col_k1,
            "Carregando": col_k2,
            "Em Rota": col_k3,
            "Concluído": col_k4
        }
        
        with col_k1: st.markdown('<div class="kanban-header">⏳ Previsão de Carregamento</div>', unsafe_allow_html=True)
        with col_k2: st.markdown('<div class="kanban-header">📦 Carregando</div>', unsafe_allow_html=True)
        with col_k3: st.markdown('<div class="kanban-header">🚚 Em Rota</div>', unsafe_allow_html=True)
        with col_k4: st.markdown('<div class="kanban-header">✅ Concluído</div>', unsafe_allow_html=True)
        
        for idx, carga in enumerate(cargas_lista):
            status_atual = carga.get("status", "Previsão de Carregamento")
            col_alvo = status_map.get(status_atual, col_k1)
            
            with col_alvo:
                with st.container():
                    st.markdown(f"**Destino:** {carga.get('destino', 'Não informado')}")
                    st.markdown(f"**Motorista:** {carga.get('motorista', 'Não informado')}")
                    
                    ajudantes_val = carga.get('ajudantes', [])
                    if isinstance(ajudantes_val, list):
                        ajudantes_str = ", ".join(ajudantes_val) if ajudantes_val else "Nenhum"
                    else:
                        ajudantes_str = str(ajudantes_val)
                    st.markdown(f"**Ajudantes:** {ajudantes_str}")
                    
                    if carga.get('data_carregamento'):
                        st.markdown(f"**Carregamento:** {formatar_data_br(carga.get('data_carregamento'))}")
                    if carga.get('data_saida'):
                        st.markdown(f"**Saída:** {formatar_data_br(carga.get('data_saida'))}")
                    
                    novo_status = st.selectbox(
                        "Status", 
                        ["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"], 
                        index=["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"].index(status_atual) if status_atual in ["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"] else 0,
                        key=f"status_select_{idx}"
                    )
                    
                    if novo_status != status_atual:
                        cargas_lista[idx]["status"] = novo_status
                        if novo_status == "Carregando" and not cargas_lista[idx].get("data_carregamento"):
                            cargas_lista[idx]["data_carregamento"] = str(datetime.date.today())
                        salvar_documento("cargas", carga.get("id"), cargas_lista[idx])
                        st.rerun()

                    bcol1, bcol2 = st.columns(2)
                    with bcol1:
                        if st.button("✏️ Editar", key=f"btn_edit_{idx}"):
                            st.session_state["editando_idx"] = idx
                            st.rerun()
                    with bcol2:
                        if st.button("🗑️ Excluir", key=f"btn_del_{idx}"):
                            deletar_documento("cargas", carga.get("id"))
                            cargas_lista.pop(idx)
                            st.session_state["cargas"] = cargas_lista
                            if st.session_state["editando_idx"] == idx:
                                st.session_state["editando_idx"] = None
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
            data_carregamento = st.date_input("Previsão / Data de Carregamento", value=None)
            data_saida = st.date_input("Data de Saída", value=None)
            
        submitted = st.form_submit_button("Salvar Carga")
        if submitted:
            if not destino:
                st.error("O destino é obrigatório!")
            else:
                novo_id = f"carga_{int(datetime.datetime.now().timestamp())}"
                dados_carga = {
                    "id": novo_id,
                    "destino": destino,
                    "motorista": motorista,
                    "ajudantes": ajudantes_selecionados,
                    "data_carregamento": str(data_carregamento) if data_carregamento else "",
                    "data_saida": str(data_saida) if data_saida else "",
                    "status": "Previsão de Carregamento"
                }
                salvar_documento("cargas", novo_id, dados_carga)
                cargas_lista.append(dados_carga)
                st.session_state["cargas"] = cargas_lista
                st.success("Carga salva com sucesso no banco de dados!")
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
                    doc_id = f"mot_{int(datetime.datetime.now().timestamp())}"
                    dados_mot = {"id": doc_id, "nome": nome_mot}
                    salvar_documento("motoristas", doc_id, dados_mot)
                    st.session_state["motoristas"].append(dados_mot)
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
                    doc_id = f"aju_{int(datetime.datetime.now().timestamp())}"
                    dados_aju = {"id": doc_id, "nome": nome_aju}
                    salvar_documento("ajudantes", doc_id, dados_aju)
                    st.session_state["ajudantes"].append(dados_aju)
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
