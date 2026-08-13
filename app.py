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

def formatar_data_br(data_str):
    if not data_str:
        return ""
    try:
        dt = datetime.datetime.strptime(str(data_str), "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return data_str

if "db_local" not in st.session_state:
    st.session_state["db_local"] = carregar_dados_locais()

db = st.session_state["db_local"]

cargas_lista = db.get("cargas", [])
motoristas_lista = [m.get("nome", "") for m in db.get("motoristas", [])]
ajudantes_lista = [a.get("nome", "") for a in db.get("ajudantes", [])]

# Controle de edição na sessão
if "editando_idx" not in st.session_state:
    st.session_state["editando_idx"] = None

tab1, tab2, tab3, tab4 = st.tabs(["📋 Painel (Kanban)", "➕ Nova Carga", "👥 Cadastros (Equipe)", "📊 Relatório Semanal"])

with tab1:
    st.subheader("Visão Geral das Cargas")
    
    # Seção de Edição de Carga (caso o usuário clique em Editar)
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
                    
                    # Ajudantes salvos anteriormente (pode ser string ou lista)
                    ajudantes_atuais = carga_edit.get("ajudantes", [])
                    if isinstance(ajudantes_atuais, str):
                        ajudantes_atuais = [a.strip() for a in ajudantes_atuais.split(",") if a.strip()]
                    novos_ajudantes = st.multiselect("Ajudantes", ajudantes_lista, default=[a for a in ajudantes_atuais if a in ajudantes_lista])
                
                with col_e2:
                    dt_car = datetime.datetime.strptime(carga_edit.get("data_carregamento", str(datetime.date.today())), "%Y-%m-%d").date() if carga_edit.get("data_carregamento") else datetime.date.today()
                    dt_said = datetime.datetime.strptime(carga_edit.get("data_saida", str(datetime.date.today())), "%Y-%m-%d").date() if carga_edit.get("data_saida") else datetime.date.today()
                    
                    nova_data_carregamento = st.date_input("Data do Carregamento Real", value=dt_car)
                    nova_data_saida = st.date_input("Data de Saída", value=dt_said)
                    
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
                    cargas_lista[idx_edit].update({
                        "destino": novo_destino,
                        "motorista": novo_motorista,
                        "ajudantes": novos_ajudantes,
                        "data_carregamento": str(nova_data_carregamento),
                        "data_saida": str(nova_data_saida),
                        "status": novo_status
                    })
                    db["cargas"] = cargas_lista
                    salvar_dados_locais(db)
                    st.session_state["editando_idx"] = None
                    st.success("Carga atualizada com sucesso!")
                    st.rerun()

                if cancelar_edicao:
                    st.session_state["editando_idx"] = None
                    st.rerun()
            st.markdown("---")

    if not cargas_lista:
        st.info("Nenhuma carga cadastrada. Utilize a aba 'Nova Carga' para começar.")
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
                    st.markdown(f"**Saída:** {formatar_data_br(carga.get('data_saida'))}")
                    
                    novo_status = st.selectbox(
                        "Status", 
                        ["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"], 
                        index=["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"].index(status_atual) if status_atual in ["Previsão de Carregamento", "Carregando", "Em Rota", "Concluído"] else 0,
                        key=f"status_select_{idx}"
                    )
                    
                    if novo_status != status_atual:
                        cargas_lista[idx]["status"] = novo_status
                        # Se mudou para carregando e não tinha data de carregamento, preenche com hoje
                        if novo_status == "Carregando" and not cargas_lista[idx].get("data_carregamento"):
                            cargas_lista[idx]["data_carregamento"] = str(datetime.date.today())
                        db["cargas"] = cargas_lista
                        salvar_dados_locais(db)
                        st.rerun()

                    # Botões de Editar e Excluir no Cartão
                    bcol1, bcol2 = st.columns(2)
                    with bcol1:
                        if st.button("✏️ Editar", key=f"btn_edit_{idx}"):
                            st.session_state["editando_idx"] = idx
                            st.rerun()
                    with bcol2:
                        if st.button("🗑️ Excluir", key=f"btn_del_{idx}"):
                            cargas_lista.pop(idx)
                            db["cargas"] = cargas_lista
                            salvar_dados_locais(db)
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
            # Data de carregamento opcional na criação (inicializada vazia ou hoje)
            informar_carregamento = st.checkbox("Definir data de carregamento agora")
            data_carregamento = st.date_input("Data do Carregamento", datetime.date.today()) if informar_carregamento else None
            data_saida = st.date_input("Data de Saída", datetime.date.today())
            
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
                    "data_carregamento": str(data_carregamento) if informar_carregamento else "",
                    "data_saida": str(data_saida),
                    "status": "Previsão de Carregamento"
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
