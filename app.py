import datetime
import os

import pandas as pd
import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Gestão de Cargas - Logística",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
        #MainMenu {
            visibility: hidden;
        }

        footer {
            visibility: hidden;
        }

        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
            max-width: 98%;
        }

        .kanban-header {
            text-align: center;
            background: linear-gradient(
                135deg,
                #21262d 0%,
                #161b22 100%
            );
            color: #ffffff !important;
            padding: 10px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 14px;
            border: 1px solid #30363d;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
        }

        div[data-testid="stVerticalBlock"]
        div[data-testid="stContainer"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
            padding: 12px 14px !important;
            margin-bottom: 10px !important;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
        }

        div[data-testid="stVerticalBlock"]
        div[data-testid="stContainer"] p,

        div[data-testid="stVerticalBlock"]
        div[data-testid="stContainer"] span {
            color: #c9d1d9 !important;
            font-size: 13px !important;
            margin-bottom: 2px !important;
        }

        .stSelectbox label,
        .stDateInput label,
        .stTextInput label,
        .stMultiSelect label {
            font-size: 12px !important;
            color: #8b949e !important;
        }

        .status-card {
            padding: 8px;
            border-radius: 6px;
            margin-bottom: 8px;
        }

        .metric-card {
            background-color: #161b22;
            border: 1px solid #30363d;
            border-radius: 10px;
            padding: 15px;
            text-align: center;
        }

        .metric-number {
            font-size: 28px;
            font-weight: bold;
            color: #ffffff;
        }

        .metric-label {
            font-size: 12px;
            color: #8b949e;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CONSTANTES
# ============================================================

STATUS_CARGA = [
    "Pendente",
    "Em Rota",
    "Entregue",
    "Cancelado",
]

COLECAO_CARGAS = "cargas"
COLECAO_MOTORISTAS = "motoristas"
COLECAO_AJUDANTES = "ajudantes"


# ============================================================
# FUNÇÕES DO FIREBASE
# ============================================================

@st.cache_resource
def conectar_firebase():
    """
    Inicializa o Firebase usando os Secrets do Streamlit.

    Espera encontrar no Streamlit Secrets:

    [firebase]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "..."
    client_email = "..."
    client_id = "..."
    auth_uri = "..."
    token_uri = "..."
    auth_provider_x509_cert_url = "..."
    client_x509_cert_url = "..."
    """

    try:
        # Se o Firebase já estiver inicializado,
        # reutilizamos a aplicação existente.
        if firebase_admin._apps:
            return firestore.client()

        # ----------------------------------------------------
        # 1. Tenta usar Streamlit Secrets
        # ----------------------------------------------------

        if "firebase" in st.secrets:
            firebase_secrets = st.secrets["firebase"]

            firebase_config = {
                "type": firebase_secrets["type"],
                "project_id": firebase_secrets["project_id"],
                "private_key_id": firebase_secrets["private_key_id"],
                "private_key": firebase_secrets["private_key"],
                "client_email": firebase_secrets["client_email"],
                "client_id": firebase_secrets["client_id"],
                "auth_uri": firebase_secrets["auth_uri"],
                "token_uri": firebase_secrets["token_uri"],
                "auth_provider_x509_cert_url": (
                    firebase_secrets["auth_provider_x509_cert_url"]
                ),
                "client_x509_cert_url": (
                    firebase_secrets["client_x509_cert_url"]
                ),
            }

            cred = credentials.Certificate(firebase_config)

            firebase_admin.initialize_app(cred)

            return firestore.client()

        # ----------------------------------------------------
        # 2. Fallback para ambiente local
        # ----------------------------------------------------
        # Isso permite continuar testando localmente com
        # serviceAccountKey.json.
        #
        # IMPORTANTE:
        # NÃO envie esse arquivo para um repositório público.
        # ----------------------------------------------------

        arquivo_local = "serviceAccountKey.json"

        if os.path.exists(arquivo_local):

            cred = credentials.Certificate(arquivo_local)

            firebase_admin.initialize_app(cred)

            return firestore.client()

        raise RuntimeError(
            "Credenciais do Firebase não encontradas. "
            "Configure os Secrets do Streamlit ou, apenas para "
            "uso local, disponibilize o arquivo serviceAccountKey.json."
        )

    except Exception as e:
        raise RuntimeError(
            f"Não foi possível inicializar o Firebase: {e}"
        ) from e


# ============================================================
# CONEXÃO
# ============================================================

try:
    db = conectar_firebase()
    firebase_conectado = True

except Exception as erro_firebase:
    db = None
    firebase_conectado = False

    st.error(
        "⚠️ Não foi possível conectar ao Firebase."
    )

    with st.expander("Ver detalhes do erro"):
        st.code(str(erro_firebase))

    st.info(
        "Verifique os Secrets do Streamlit e as permissões "
        "do Firebase/Firestore."
    )


# ============================================================
# FUNÇÕES DE LEITURA
# ============================================================

def carregar_dados_seguro(colecao):
    """
    Carrega documentos de uma coleção do Firestore.

    O timeout evita que o aplicativo fique indefinidamente
    aguardando uma resposta do banco.
    """

    if db is None:
        return []

    try:
        documentos = db.collection(colecao).stream(
            timeout=15
        )

        dados = []

        for documento in documentos:
            registro = documento.to_dict() or {}

            # Guarda o ID real do documento caso o campo "id"
            # não exista.
            if "id" not in registro:
                registro["id"] = documento.id

            dados.append(registro)

        return dados

    except Exception as erro:
        st.warning(
            f"⚠️ Não foi possível carregar a coleção "
            f"'{colecao}'."
        )

        with st.expander(
            f"Detalhes do erro - {colecao}"
        ):
            st.code(str(erro))

        return []


# ============================================================
# FUNÇÕES DE ESCRITA
# ============================================================

def salvar_dado_seguro(colecao, dados, doc_id):
    """
    Salva ou atualiza um documento no Firestore.
    """

    if db is None:
        st.error(
            "Firebase não está conectado. "
            "O dado não foi salvo."
        )
        return False

    if doc_id is None or str(doc_id).strip() == "":
        st.error(
            "Não foi possível salvar: ID do documento vazio."
        )
        return False

    try:
        doc_id = str(doc_id).strip()

        # Evita erro de caminho do Firestore caso alguém
        # coloque "/" dentro do ID.
        if "/" in doc_id:
            st.error(
                "O ID não pode conter o caractere '/'."
            )
            return False

        db.collection(colecao).document(doc_id).set(
            dados
        )

        return True

    except Exception as erro:
        st.error(
            f"❌ Erro ao salvar na coleção "
            f"'{colecao}': {erro}"
        )
        return False


# ============================================================
# CARREGAMENTO DOS DADOS
# ============================================================

motoristas_raw = []
ajudantes_raw = []
cargas_lista = []

if firebase_conectado:

    with st.spinner("Carregando dados do Firebase..."):

        motoristas_raw = carregar_dados_seguro(
            COLECAO_MOTORISTAS
        )

        ajudantes_raw = carregar_dados_seguro(
            COLECAO_AJUDANTES
        )

        cargas_lista = carregar_dados_seguro(
            COLECAO_CARGAS
        )


# ============================================================
# PREPARAÇÃO DAS LISTAS
# ============================================================

motoristas_lista = sorted(
    {
        str(m.get("nome", "")).strip()
        for m in motoristas_raw
        if str(m.get("nome", "")).strip()
    }
)

ajudantes_lista = sorted(
    {
        str(a.get("nome", "")).strip()
        for a in ajudantes_raw
        if str(a.get("nome", "")).strip()
    }
)


# ============================================================
# FUNÇÃO PARA CONTAR STATUS
# ============================================================

def contar_status(cargas):
    contagem = {
        "Pendente": 0,
        "Em Rota": 0,
        "Entregue": 0,
        "Cancelado": 0,
    }

    for carga in cargas:
        status = carga.get(
            "status",
            "Pendente"
        )

        if status not in contagem:
            status = "Pendente"

        contagem[status] += 1

    return contagem


# ============================================================
# TÍTULO
# ============================================================

st.title(
    "🚚 Painel de Controle de Cargas e Agendamentos"
)

# Indicador da conexão
if firebase_conectado:
    st.success(
        "🟢 Firebase conectado"
    )
else:
    st.error(
        "🔴 Firebase desconectado"
    )


# ============================================================
# BOTÃO ATUALIZAR
# ============================================================

col_refresh_1, col_refresh_2 = st.columns(
    [8, 1]
)

with col_refresh_2:

    if st.button(
        "🔄 Atualizar",
        use_container_width=True
    ):
        st.cache_resource.clear()
        st.rerun()


# ============================================================
# INDICADORES
# ============================================================

contagem = contar_status(cargas_lista)

m1, m2, m3, m4, m5 = st.columns(5)

with m1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-number">
                {len(cargas_lista)}
            </div>
            <div class="metric-label">
                Total de Cargas
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with m2:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-number">
                {contagem["Pendente"]}
            </div>
            <div class="metric-label">
                📌 Pendentes
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with m3:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-number">
                {contagem["Em Rota"]}
            </div>
            <div class="metric-label">
                🚚 Em Rota
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with m4:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-number">
                {contagem["Entregue"]}
            </div>
            <div class="metric-label">
                ✅ Entregues
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with m5:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-number">
                {contagem["Cancelado"]}
            </div>
            <div class="metric-label">
                ❌ Canceladas
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.divider()


# ============================================================
# ABAS
# ============================================================

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📋 Painel (Kanban)",
        "➕ Nova Carga",
        "👥 Cadastros (Equipe)",
        "📊 Relatório Semanal",
    ]
)


# ============================================================
# ABA 1 - KANBAN
# ============================================================

with tab1:

    st.subheader(
        "Visão Geral das Cargas"
    )

    if not firebase_conectado:

        st.warning(
            "O painel está aguardando a conexão com o Firebase."
        )

    elif not cargas_lista:

        st.info(
            "Nenhuma carga encontrada no banco de dados. "
            "Utilize a aba 'Nova Carga' para cadastrar."
        )

    else:

        col_k1, col_k2, col_k3, col_k4 = st.columns(4)

        status_map = {
            "Pendente": col_k1,
            "Em Rota": col_k2,
            "Entregue": col_k3,
            "Cancelado": col_k4,
        }

        # Cabeçalhos
        with col_k1:
            st.markdown(
                '<div class="kanban-header">'
                '📌 Pendente'
                '</div>',
                unsafe_allow_html=True,
            )

        with col_k2:
            st.markdown(
                '<div class="kanban-header">'
                '🚚 Em Rota'
                '</div>',
                unsafe_allow_html=True,
            )

        with col_k3:
            st.markdown(
                '<div class="kanban-header">'
                '✅ Entregue'
                '</div>',
                unsafe_allow_html=True,
            )

        with col_k4:
            st.markdown(
                '<div class="kanban-header">'
                '❌ Cancelado'
                '</div>',
                unsafe_allow_html=True,
            )

        # Cargas
        for indice, carga in enumerate(cargas_lista):

            carga_id = str(
                carga.get(
                    "id",
                    f"carga_{indice}"
                )
            )

            status_atual = carga.get(
                "status",
                "Pendente"
            )

            if status_atual not in STATUS_CARGA:
                status_atual = "Pendente"

            col_alvo = status_map.get(
                status_atual,
                col_k1
            )

            with col_alvo:

                with st.container():

                    st.markdown(
                        f"**🆔 ID:** {carga_id}"
                    )

                    st.markdown(
                        "**👤 Motorista:** "
                        f"{carga.get('motorista', 'Não informado')}"
                    )

                    st.markdown(
                        "**📍 Destino:** "
                        f"{carga.get('destino', 'Não informado')}"
                    )

                    st.markdown(
                        "**📅 Saída:** "
                        f"{carga.get('data_saida', '')}"
                    )

                    st.markdown(
                        "**🏁 Entrega:** "
                        f"{carga.get('data_entrega', '')}"
                    )

                    ajudantes = carga.get(
                        "ajudantes",
                        []
                    )

                    if isinstance(
                        ajudantes,
                        list
                    ):
                        ajudantes_texto = ", ".join(
                            map(str, ajudantes)
                        )
                    else:
                        ajudantes_texto = str(
                            ajudantes
                        )

                    if ajudantes_texto:
                        st.markdown(
                            "**👥 Ajudantes:** "
                            f"{ajudantes_texto}"
                        )

                    novo_status = st.selectbox(
                        "Status",
                        STATUS_CARGA,
                        index=STATUS_CARGA.index(
                            status_atual
                        ),
                        key=(
                            f"status_"
                            f"{carga_id}_"
                            f"{indice}"
                        ),
                    )

                    if novo_status != status_atual:

                        carga_atualizada = dict(
                            carga
                        )

                        carga_atualizada[
                            "status"
                        ] = novo_status

                        sucesso = salvar_dado_seguro(
                            COLECAO_CARGAS,
                            carga_atualizada,
                            carga_id,
                        )

                        if sucesso:

                            st.success(
                                "Status atualizado!"
                            )

                            st.cache_resource.clear()
                            st.rerun()


# ============================================================
# ABA 2 - NOVA CARGA
# ============================================================

with tab2:

    st.subheader(
        "Cadastrar Nova Carga"
    )

    if not firebase_conectado:

        st.warning(
            "Conecte o Firebase antes de cadastrar uma carga."
        )

    else:

        with st.form(
            "form_nova_carga",
            clear_on_submit=True
        ):

            col_f1, col_f2 = st.columns(2)

            with col_f1:

                carga_id = st.text_input(
                    "ID / Número do Pedido ou Carga",
                    placeholder="Ex.: CARGA-001",
                )

                if motoristas_lista:

                    motorista = st.selectbox(
                        "Motorista",
                        motoristas_lista,
                    )

                else:

                    motorista = st.text_input(
                        "Motorista",
                        placeholder=(
                            "Cadastre um motorista "
                            "na aba Equipe"
                        ),
                    )

                destino = st.text_input(
                    "Destino (Cidade / Bairro)",
                    placeholder="Ex.: São Paulo - SP",
                )

            with col_f2:

                ajudantes_selecionados = (
                    st.multiselect(
                        "Ajudantes",
                        ajudantes_lista,
                    )
                )

                data_saida = st.date_input(
                    "Data de Saída",
                    datetime.date.today(),
                )

                data_entrega = st.date_input(
                    "Previsão de Entrega",
                    datetime.date.today(),
                )

            submitted = st.form_submit_button(
                "💾 Salvar Carga no Firebase",
                use_container_width=True,
            )

            if submitted:

                carga_id = carga_id.strip()
                destino = destino.strip()
                motorista = motorista.strip()

                if not carga_id:

                    st.error(
                        "❌ O ID da carga é obrigatório!"
                    )

                elif "/" in carga_id:

                    st.error(
                        "❌ O ID da carga não pode "
                        "conter '/'."
                    )

                elif not destino:

                    st.error(
                        "❌ Informe o destino da carga."
                    )

                elif data_entrega < data_saida:

                    st.error(
                        "❌ A previsão de entrega não "
                        "pode ser anterior à data de saída."
                    )

                else:

                    dados_carga = {
                        "id": carga_id,
                        "motorista": motorista,
                        "destino": destino,
                        "ajudantes": (
                            ajudantes_selecionados
                        ),
                        "data_saida": str(
                            data_saida
                        ),
                        "data_entrega": str(
                            data_entrega
                        ),
                        "status": "Pendente",
                        "criado_em": (
                            datetime.datetime.now(
                                datetime.timezone.utc
                            ).isoformat()
                        ),
                    }

                    sucesso = salvar_dado_seguro(
                        COLECAO_CARGAS,
                        dados_carga,
                        carga_id,
                    )

                    if sucesso:

                        st.success(
                            f"✅ Carga {carga_id} "
                            "salva com sucesso!"
                        )

                        st.cache_resource.clear()
                        st.rerun()


# ============================================================
# ABA 3 - CADASTROS
# ============================================================

with tab3:

    st.subheader(
        "Gerenciamento de Cadastros (Equipe)"
    )

    col_c1, col_c2 = st.columns(2)

    # --------------------------------------------------------
    # MOTORISTAS
    # --------------------------------------------------------

    with col_c1:

        st.markdown(
            "### 🚚 Cadastrar Motorista"
        )

        with st.form(
            "form_motorista",
            clear_on_submit=True
        ):

            nome_mot = st.text_input(
                "Nome do Motorista",
                placeholder="Nome completo",
            )

            adicionar_motorista = (
                st.form_submit_button(
                    "➕ Adicionar Motorista",
                    use_container_width=True,
                )
            )

            if adicionar_motorista:

                nome_mot = nome_mot.strip()

                if not nome_mot:

                    st.error(
                        "Informe o nome do motorista."
                    )

                elif "/" in nome_mot:

                    st.error(
                        "O nome não pode conter '/'."
                    )

                elif nome_mot in motoristas_lista:

                    st.warning(
                        "Esse motorista já está cadastrado."
                    )

                else:

                    sucesso = salvar_dado_seguro(
                        COLECAO_MOTORISTAS,
                        {
                            "nome": nome_mot
                        },
                        nome_mot,
                    )

                    if sucesso:

                        st.success(
                            f"✅ Motorista "
                            f"{nome_mot} adicionado!"
                        )

                        st.cache_resource.clear()
                        st.rerun()

        st.markdown(
            "#### Motoristas Cadastrados:"
        )

        if motoristas_lista:

            for motorista_nome in motoristas_lista:

                st.markdown(
                    f"• {motorista_nome}"
                )

        else:

            st.info(
                "Nenhum motorista cadastrado."
            )

    # --------------------------------------------------------
    # AJUDANTES
    # --------------------------------------------------------

    with col_c2:

        st.markdown(
            "### 👥 Cadastrar Ajudante"
        )

        with st.form(
            "form_ajudante",
            clear_on_submit=True
        ):

            nome_aju = st.text_input(
                "Nome do Ajudante",
                placeholder="Nome completo",
            )

            adicionar_ajudante = (
                st.form_submit_button(
                    "➕ Adicionar Ajudante",
                    use_container_width=True,
                )
            )

            if adicionar_ajudante:

                nome_aju = nome_aju.strip()

                if not nome_aju:

                    st.error(
                        "Informe o nome do ajudante."
                    )

                elif "/" in nome_aju:

                    st.error(
                        "O nome não pode conter '/'."
                    )

                elif nome_aju in ajudantes_lista:

                    st.warning(
                        "Esse ajudante já está cadastrado."
                    )

                else:

                    sucesso = salvar_dado_seguro(
                        COLECAO_AJUDANTES,
                        {
                            "nome": nome_aju
                        },
                        nome_aju,
                    )

                    if sucesso:

                        st.success(
                            f"✅ Ajudante "
                            f"{nome_aju} adicionado!"
                        )

                        st.cache_resource.clear()
                        st.rerun()

        st.markdown(
            "#### Ajudantes Cadastrados:"
        )

        if ajudantes_lista:

            for ajudante_nome in ajudantes_lista:

                st.markdown(
                    f"• {ajudante_nome}"
                )

        else:

            st.info(
                "Nenhum ajudante cadastrado."
            )


# ============================================================
# ABA 4 - RELATÓRIO
# ============================================================

with tab4:

    st.subheader(
        "Relatório Semanal e Exportação"
    )

    if not cargas_lista:

        st.info(
            "Sem dados para gerar relatório."
        )

    else:

        # ----------------------------------------------------
        # Conversão para DataFrame
        # ----------------------------------------------------

        df_relatorio = pd.DataFrame(
            cargas_lista
        )

        # Garante colunas
        colunas_esperadas = [
            "id",
            "motorista",
            "destino",
            "ajudantes",
            "data_saida",
            "data_entrega",
            "status",
        ]

        for coluna in colunas_esperadas:

            if coluna not in df_relatorio.columns:

                df_relatorio[coluna] = ""

        # ----------------------------------------------------
        # Filtro de período
        # ----------------------------------------------------

        col_data_1, col_data_2 = st.columns(2)

        hoje = datetime.date.today()

        inicio_semana = hoje - datetime.timedelta(
            days=hoje.weekday()
        )

        with col_data_1:

            data_inicio_relatorio = st.date_input(
                "Data inicial",
                inicio_semana,
            )

        with col_data_2:

            data_fim_relatorio = st.date_input(
                "Data final",
                hoje,
            )

        if data_fim_relatorio < data_inicio_relatorio:

            st.error(
                "A data final não pode ser anterior "
                "à data inicial."
            )

        else:

            # ------------------------------------------------
            # Conversão da data
            # ------------------------------------------------

            df_relatorio[
                "data_saida_dt"
            ] = pd.to_datetime(
                df_relatorio["data_saida"],
                errors="coerce",
            )

            # ------------------------------------------------
            # Filtro
            # ------------------------------------------------

            filtro = (
                df_relatorio["data_saida_dt"].dt.date
                >= data_inicio_relatorio
            ) & (
                df_relatorio["data_saida_dt"].dt.date
                <= data_fim_relatorio
            )

            df_filtrado = df_relatorio[
                filtro
            ].copy()

            # ------------------------------------------------
            # Resumo
            # ------------------------------------------------

            st.markdown(
                "### 📊 Resumo do Período"
            )

            r1, r2, r3, r4 = st.columns(4)

            with r1:

                st.metric(
                    "Cargas",
                    len(df_filtrado),
                )

            with r2:

                st.metric(
                    "Pendentes",
                    len(
                        df_filtrado[
                            df_filtrado["status"]
                            == "Pendente"
                        ]
                    ),
                )

            with r3:

                st.metric(
                    "Em Rota",
                    len(
                        df_filtrado[
                            df_filtrado["status"]
                            == "Em Rota"
                        ]
                    ),
                )

            with r4:

                st.metric(
                    "Entregues",
                    len(
                        df_filtrado[
                            df_filtrado["status"]
                            == "Entregue"
                        ]
                    ),
                )

            st.divider()

            # ------------------------------------------------
            # Tabela
            # ------------------------------------------------

            st.markdown(
                "### 📋 Cargas do Período"
            )

            colunas_exibicao = [
                "id",
                "motorista",
                "destino",
                "ajudantes",
                "data_saida",
                "data_entrega",
                "status",
            ]

            df_exibicao = df_filtrado[
                colunas_exibicao
            ].copy()

            df_exibicao.columns = [
                "ID",
                "Motorista",
                "Destino",
                "Ajudantes",
                "Data de Saída",
                "Previsão de Entrega",
                "Status",
            ]

            st.dataframe(
                df_exibicao,
                use_container_width=True,
                hide_index=True,
            )

            # ------------------------------------------------
            # CSV
            # ------------------------------------------------

            csv = df_exibicao.to_csv(
                index=False
            ).encode("utf-8-sig")

            st.download_button(
                label="⬇️ Baixar Relatório em CSV",
                data=csv,
                file_name=(
                    "relatorio_cargas.csv"
                ),
                mime="text/csv",
                use_container_width=True,
            )


# ============================================================
# RODAPÉ
# ============================================================

st.divider()

st.caption(
    "🚚 Sistema de Gestão de Cargas e Agendamentos"
)
