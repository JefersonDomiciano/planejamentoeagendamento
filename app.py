import datetime
import io
import json
import os
import requests
import pandas as pd
import streamlit as st
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


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
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
            max-width: 100%;
        }

        /* =========================
           KANBAN
        ========================== */

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

        .stSelectbox label,
        .stDateInput label,
        .stTextInput label,
        .stMultiSelect label {
            font-size: 12px !important;
            color: #8b949e !important;
        }

        /* =========================
           CARDS DE INDICADORES
        ========================== */

        .metric-card {
            background: linear-gradient(135deg, #161b22 0%, #21262d 100%);
            border: 1px solid #30363d;
            border-radius: 10px;
            padding: 14px;
            min-height: 105px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.18);
        }

        .metric-title {
            color: #8b949e;
            font-size: 12px;
            font-weight: 600;
            margin-bottom: 5px;
        }

        .metric-value {
            color: #ffffff;
            font-size: 28px;
            font-weight: 700;
        }

        .metric-subtitle {
            color: #8b949e;
            font-size: 11px;
            margin-top: 3px;
        }

        .metric-blue {
            border-left: 4px solid #58a6ff;
        }

        .metric-green {
            border-left: 4px solid #3fb950;
        }

        .metric-yellow {
            border-left: 4px solid #d29922;
        }

        .metric-purple {
            border-left: 4px solid #bc8cff;
        }

        .metric-red {
            border-left: 4px solid #f85149;
        }

        /* =========================
           ALERTAS
        ========================== */

        .alert-box {
            padding: 12px 15px;
            border-radius: 8px;
            margin-bottom: 10px;
            font-size: 13px;
            font-weight: 500;
        }

        .alert-red {
            background: rgba(248, 81, 73, 0.12);
            border: 1px solid #f85149;
            color: #ff7b72;
        }

        .alert-yellow {
            background: rgba(210, 153, 34, 0.12);
            border: 1px solid #d29922;
            color: #e3b341;
        }

        .alert-green {
            background: rgba(63, 185, 80, 0.12);
            border: 1px solid #3fb950;
            color: #56d364;
        }

        /* =========================
           BADGES
        ========================== */

        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 20px;
            font-size: 10px;
            font-weight: 700;
            margin-right: 4px;
            margin-bottom: 4px;
        }

        .badge-red {
            background: rgba(248, 81, 73, 0.18);
            color: #ff7b72;
            border: 1px solid rgba(248, 81, 73, 0.4);
        }

        .badge-yellow {
            background: rgba(210, 153, 34, 0.18);
            color: #e3b341;
            border: 1px solid rgba(210, 153, 34, 0.4);
        }

        .badge-green {
            background: rgba(63, 185, 80, 0.18);
            color: #56d364;
            border: 1px solid rgba(63, 185, 80, 0.4);
        }

        .badge-blue {
            background: rgba(88, 166, 255, 0.18);
            color: #58a6ff;
            border: 1px solid rgba(88, 166, 255, 0.4);
        }

        /* =========================
           MOBILE
        ========================== */

        @media (max-width: 768px) {

            .block-container {
                padding-left: 0.5rem;
                padding-right: 0.5rem;
                padding-top: 0.5rem;
            }

            h1 {
                font-size: 22px !important;
            }

            h2 {
                font-size: 20px !important;
            }

            h3 {
                font-size: 17px !important;
            }

            .metric-card {
                min-height: 90px;
                padding: 10px;
            }

            .metric-value {
                font-size: 23px;
            }

            .kanban-header {
                font-size: 12px;
                padding: 8px 5px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# TÍTULO
# ============================================================

st.title("🚚 Painel de Controle de Cargas e Agendamentos")


# ============================================================
# FIREBASE
# ============================================================

FIREBASE_PROJECT_ID = "logistica-d6c14"


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def formatar_data_br(data_str):
    if not data_str or str(data_str).lower() in ["nan", "none", ""]:
        return ""

    try:
        dt = datetime.datetime.strptime(
            str(data_str).split("T")[0],
            "%Y-%m-%d"
        )
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(data_str)


def converter_para_data(data_str):
    """
    Converte datas armazenadas no Firebase para datetime.date.
    Retorna None quando não for possível converter.
    """
    if not data_str:
        return None

    try:
        return datetime.date.fromisoformat(
            str(data_str).split("T")[0]
        )
    except Exception:
        return None


def carregar_colecao(colecao):
    url = (
        f"https://firestore.googleapis.com/v1/projects/"
        f"{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}"
    )

    try:
        response = requests.get(url, timeout=10)

        if response.status_code == 200:
            data = response.json()
            documentos = []

            for doc in data.get("documents", []):
                doc_id = doc["name"].split("/")[-1]
                fields = doc.get("fields", {})

                item = {"id": doc_id}

                for k, v in fields.items():

                    if "arrayValue" in v:
                        arr_vals = v["arrayValue"].get("values", [])

                        valores = []

                        for x in arr_vals:
                            if "stringValue" in x:
                                valores.append(x["stringValue"])
                            elif "integerValue" in x:
                                valores.append(int(x["integerValue"]))
                            elif "doubleValue" in x:
                                valores.append(float(x["doubleValue"]))
                            elif "booleanValue" in x:
                                valores.append(bool(x["booleanValue"]))

                        item[k] = valores

                    elif "stringValue" in v:
                        item[k] = v["stringValue"]

                    elif "integerValue" in v:
                        item[k] = int(v["integerValue"])

                    elif "doubleValue" in v:
                        item[k] = float(v["doubleValue"])

                    elif "booleanValue" in v:
                        item[k] = bool(v["booleanValue"])

                    elif "nullValue" in v:
                        item[k] = None

                    else:
                        item[k] = list(v.values())[0]

                documentos.append(item)

            return documentos

        else:
            st.warning(
                f"Não foi possível carregar '{colecao}'. "
                f"Firebase retornou HTTP {response.status_code}."
            )

    except requests.exceptions.Timeout:
        st.error(f"Tempo esgotado ao consultar a coleção '{colecao}'.")

    except requests.exceptions.RequestException as e:
        st.error(f"Erro de conexão com o Firebase: {e}")

    except Exception as e:
        st.error(f"Erro ao carregar '{colecao}': {e}")

    return []


def salvar_documento(colecao, doc_id, dados):
    url = (
        f"https://firestore.googleapis.com/v1/projects/"
        f"{FIREBASE_PROJECT_ID}/databases/(default)/documents/"
        f"{colecao}/{doc_id}"
    )

    fields = {}

    for k, v in dados.items():

        if isinstance(v, list):

            values = []

            for x in v:
                values.append({
                    "stringValue": str(x)
                })

            fields[k] = {
                "arrayValue": {
                    "values": values
                }
            }

        elif isinstance(v, bool):

            fields[k] = {
                "booleanValue": v
            }

        elif isinstance(v, int):

            fields[k] = {
                "integerValue": str(v)
            }

        elif isinstance(v, float):

            fields[k] = {
                "doubleValue": v
            }

        elif v is None:

            fields[k] = {
                "nullValue": None
            }

        else:

            fields[k] = {
                "stringValue": str(v)
            }

    try:

        response = requests.patch(
            url,
            json={"fields": fields},
            timeout=10
        )

        if response.status_code not in [200, 201]:

            st.error(
                f"Erro ao salvar no Firebase. "
                f"HTTP {response.status_code}"
            )

            return False

        return True

    except requests.exceptions.Timeout:

        st.error("Tempo esgotado ao salvar no Firebase.")
        return False

    except requests.exceptions.RequestException as e:

        st.error(f"Erro de conexão com o Firebase: {e}")
        return False

    except Exception as e:

        st.error(f"Erro ao salvar no Firebase: {e}")
        return False


def deletar_documento(colecao, doc_id):
    url = (
        f"https://firestore.googleapis.com/v1/projects/"
        f"{FIREBASE_PROJECT_ID}/databases/(default)/documents/"
        f"{colecao}/{doc_id}"
    )

    try:

        response = requests.delete(
            url,
            timeout=10
        )

        if response.status_code not in [200, 204]:

            st.error(
                f"Erro ao excluir documento. "
                f"HTTP {response.status_code}"
            )

            return False

        return True

    except requests.exceptions.Timeout:

        st.error("Tempo esgotado ao excluir do Firebase.")
        return False

    except requests.exceptions.RequestException as e:

        st.error(f"Erro de conexão com o Firebase: {e}")
        return False

    except Exception as e:

        st.error(f"Erro ao excluir no Firebase: {e}")
        return False


def atualizar_dados():

    st.session_state["cargas"] = carregar_colecao("cargas")
    st.session_state["motoristas"] = carregar_colecao("motoristas")
    st.session_state["ajudantes"] = carregar_colecao("ajudantes")

    st.session_state["ultima_atualizacao"] = datetime.datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )


def carga_atrasada(carga):

    status = carga.get("status", "")

    if status == "Entregue / Concluído":
        return False

    data_entrega = converter_para_data(
        carga.get("data_entrega")
    )

    if not data_entrega:
        return False

    return data_entrega < datetime.date.today()


def carga_saida_hoje(carga):

    data_saida = converter_para_data(
        carga.get("data_saida")
    )

    if not data_saida:
        return False

    return data_saida == datetime.date.today()


def carga_entrega_hoje(carga):

    data_entrega = converter_para_data(
        carga.get("data_entrega")
    )

    if not data_entrega:
        return False

    return data_entrega == datetime.date.today()


def dias_para_entrega(carga):

    data_entrega = converter_para_data(
        carga.get("data_entrega")
    )

    if not data_entrega:
        return None

    return (data_entrega - datetime.date.today()).days


def texto_prazo(carga):

    dias = dias_para_entrega(carga)

    if dias is None:
        return ""

    if carga_atrasada(carga):
        dias_atrasados = abs(dias)

        if dias_atrasados == 1:
            return "🔴 1 dia atrasada"

        return f"🔴 {dias_atrasados} dias atrasada"

    if dias == 0:
        return "🟠 Entrega hoje"

    if dias == 1:
        return "🟡 Entrega amanhã"

    if dias < 0:
        return "🔴 Atrasada"

    return f"🟢 {dias} dias para entrega"


def preparar_dataframe(cargas_lista):

    df = pd.DataFrame(cargas_lista)

    if df.empty:
        return df

    if "ajudantes" in df.columns:

        df["ajudantes"] = df["ajudantes"].apply(
            lambda x: ", ".join(map(str, x))
            if isinstance(x, list)
            else str(x)
        )

    for col in [
        "data_carga",
        "data_saida",
        "data_entrega"
    ]:

        if col in df.columns:

            df[col] = df[col].apply(
                formatar_data_br
            )

    colunas_desejadas = [
        "id",
        "motorista",
        "destino",
        "observacoes",
        "ajudantes",
        "data_carga",
        "data_saida",
        "data_entrega",
        "status"
    ]

    colunas_existentes = [
        col
        for col in colunas_desejadas
        if col in df.columns
    ]

    df = df[colunas_existentes]

    return df


# ============================================================
# EXCEL
# ============================================================

def gerar_excel_profissional(df):

    output = io.BytesIO()

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Relatorio de Cargas"

    header_font = Font(
        name="Arial",
        size=10,
        bold=True,
        color="FFFFFF"
    )

    header_fill = PatternFill(
        start_color="2F75B5",
        end_color="2F75B5",
        fill_type="solid"
    )

    data_font = Font(
        name="Arial",
        size=9
    )

    center_alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    left_alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    headers = [
        "ID Planejamento",
        "Motorista",
        "Destino",
        "Observações",
        "Ajudantes",
        "Data Carga",
        "Data Saída",
        "Data Entrega",
        "Status"
    ]

    for col_num, header in enumerate(
        headers[:len(df.columns)],
        1
    ):

        cell = ws.cell(
            row=1,
            column=col_num
        )

        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border

    ws.row_dimensions[1].height = 24

    for row_num, row_data in enumerate(
        df.values,
        2
    ):

        ws.row_dimensions[row_num].height = 20

        for col_num, value in enumerate(
            row_data,
            1
        ):

            cell = ws.cell(
                row=row_num,
                column=col_num
            )

            cell.value = (
                ""
                if str(value).lower() in ["nan", "none"]
                else value
            )

            cell.font = data_font
            cell.border = thin_border

            if col_num in [1, 6, 7, 8]:

                cell.alignment = center_alignment

            else:

                cell.alignment = left_alignment

    for col in ws.columns:

        max_len = max(
            len(str(cell.value or ""))
            for cell in col
        )

        col_letter = get_column_letter(
            col[0].column
        )

        ws.column_dimensions[
            col_letter
        ].width = max(
            max_len + 5,
            15
        )

    wb.save(output)

    return output.getvalue()


# ============================================================
# PDF
# ============================================================

def gerar_pdf(df):

    pdf = FPDF(
        orientation="L",
        unit="mm",
        format="A4"
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        "B",
        14
    )

    pdf.cell(
        277,
        8,
        txt="Relatório de Cargas",
        ln=True,
        align="C"
    )

    pdf.set_font(
        "Arial",
        "",
        9
    )

    pdf.cell(
        277,
        5,
        txt=(
            f"Data de geração: "
            f"{datetime.date.today().strftime('%d/%m/%Y')}"
        ),
        ln=True,
        align="C"
    )

    pdf.ln(4)

    pdf.set_font(
        "Arial",
        "B",
        9
    )

    pdf.set_fill_color(
        47,
        117,
        181
    )

    pdf.set_text_color(
        255,
        255,
        255
    )

    larguras = [
        25,
        40,
        50,
        45,
        28,
        28,
        28,
        33
    ]

    nomes_colunas = [
        "ID",
        "Motorista",
        "Destino",
        "Ajudantes",
        "Carga",
        "Saída",
        "Entrega",
        "Status"
    ]

    for i, nome in enumerate(
        nomes_colunas
    ):

        pdf.cell(
            larguras[i],
            7,
            nome,
            1,
            0,
            "C",
            True
        )

    pdf.ln()

    pdf.set_font(
        "Arial",
        "",
        8
    )

    pdf.set_text_color(
        0,
        0,
        0
    )

    for _, row in df.iterrows():

        pdf.cell(
            larguras[0],
            6,
            str(row.get("id", "")),
            1,
            0,
            "C"
        )

        pdf.cell(
            larguras[1],
            6,
            str(row.get("motorista", ""))[:22],
            1,
            0,
            "L"
        )

        pdf.cell(
            larguras[2],
            6,
            str(row.get("destino", ""))[:30],
            1,
            0,
            "L"
        )

        pdf.cell(
            larguras[3],
            6,
            str(row.get("ajudantes", ""))[:25],
            1,
            0,
            "L"
        )

        pdf.cell(
            larguras[4],
            6,
            str(row.get("data_carga", "")),
            1,
            0,
            "C"
        )

        pdf.cell(
            larguras[5],
            6,
            str(row.get("data_saida", "")),
            1,
            0,
            "C"
        )

        pdf.cell(
            larguras[6],
            6,
            str(row.get("data_entrega", "")),
            1,
            0,
            "C"
        )

        pdf.cell(
            larguras[7],
            6,
            str(row.get("status", ""))[:18],
            1,
            0,
            "C"
        )

        pdf.ln()

    return bytes(
        pdf.output(dest="S")
    )


# ============================================================
# CARREGAMENTO INICIAL
# ============================================================

if "cargas" not in st.session_state:
    st.session_state["cargas"] = carregar_colecao("cargas")

if "motoristas" not in st.session_state:
    st.session_state["motoristas"] = carregar_colecao("motoristas")

if "ajudantes" not in st.session_state:
    st.session_state["ajudantes"] = carregar_colecao("ajudantes")


# ============================================================
# LISTAS DA EQUIPE
# ============================================================

cargas_lista = st.session_state["cargas"]

motoristas_lista = [
    m.get("nome", "")
    for m in st.session_state["motoristas"]
    if m.get("nome")
]

ajudantes_lista = [
    a.get("nome", "")
    for a in st.session_state["ajudantes"]
    if a.get("nome")
]


if not motoristas_lista:

    motoristas_lista = [
        "Carlos Silva",
        "João Pereira",
        "Maurício",
        "Cícero Taveira"
    ]


if not ajudantes_lista:

    ajudantes_lista = [
        "Pedrinho",
        "Lucas Souza"
    ]


# ============================================================
# MENU
# ============================================================

menu = st.radio(
    "Menu Principal",
    [
        "📋 Painel (Kanban)",
        "➕ Nova Carga",
        "👥 Cadastros (Equipe)",
        "📈 Relatórios",
    ],
    horizontal=True,
)


# ============================================================
# BARRA DE ATUALIZAÇÃO
# ============================================================

col_atualizacao1, col_atualizacao2 = st.columns(
    [6, 1]
)

with col_atualizacao1:

    if "ultima_atualizacao" in st.session_state:

        st.caption(
            f"🕐 Última atualização: "
            f"{st.session_state['ultima_atualizacao']}"
        )

with col_atualizacao2:

    if st.button(
        "🔄 Atualizar",
        use_container_width=True
    ):

        atualizar_dados()
        st.rerun()


st.markdown("---")


# ============================================================
# 1. PAINEL KANBAN
# ============================================================

if menu == "📋 Painel (Kanban)":

    st.subheader(
        "📊 Torre de Controle da Operação"
    )

    hoje = datetime.date.today()

    total_cargas = len(cargas_lista)

    total_aguardando = sum(
        1
        for c in cargas_lista
        if c.get("status") == "Aguardando Carregamento"
    )

    total_patio = sum(
        1
        for c in cargas_lista
        if c.get("status") == "Carregado / No Pátio"
    )

    total_transito = sum(
        1
        for c in cargas_lista
        if c.get("status") == "Em Trânsito / Viagem Iniciada"
    )

    total_entregues = sum(
        1
        for c in cargas_lista
        if c.get("status") == "Entregue / Concluído"
    )

    total_atrasadas = sum(
        1
        for c in cargas_lista
        if carga_atrasada(c)
    )

    total_saida_hoje = sum(
        1
        for c in cargas_lista
        if carga_saida_hoje(c)
    )

    total_entrega_hoje = sum(
        1
        for c in cargas_lista
        if carga_entrega_hoje(c)
    )


    # ========================================================
    # INDICADORES
    # ========================================================

    m1, m2, m3, m4, m5 = st.columns(5)

    with m1:

        st.markdown(
            f"""
            <div class="metric-card metric-blue">
                <div class="metric-title">TOTAL DE CARGAS</div>
                <div class="metric-value">{total_cargas}</div>
                <div class="metric-subtitle">Planejamentos cadastrados</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with m2:

        st.markdown(
            f"""
            <div class="metric-card metric-yellow">
                <div class="metric-title">AGUARDANDO</div>
                <div class="metric-value">{total_aguardando}</div>
                <div class="metric-subtitle">Aguardando carregamento</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with m3:

        st.markdown(
            f"""
            <div class="metric-card metric-green">
                <div class="metric-title">EM TRÂNSITO</div>
                <div class="metric-value">{total_transito}</div>
                <div class="metric-subtitle">Viagens iniciadas</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with m4:

        st.markdown(
            f"""
            <div class="metric-card metric-purple">
                <div class="metric-title">ENTREGUES</div>
                <div class="metric-value">{total_entregues}</div>
                <div class="metric-subtitle">Operações concluídas</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with m5:

        st.markdown(
            f"""
            <div class="metric-card metric-red">
                <div class="metric-title">ATRASADAS</div>
                <div class="metric-value">{total_atrasadas}</div>
                <div class="metric-subtitle">Precisam de atenção</div>
            </div>
            """,
            unsafe_allow_html=True
        )


    st.markdown("")


    # ========================================================
    # ALERTAS
    # ========================================================

    if total_atrasadas > 0:

        st.markdown(
            f"""
            <div class="alert-box alert-red">
                🔴 <b>Atenção:</b> existem
                <b>{total_atrasadas}</b> carga(s)
                com prazo de entrega vencido.
            </div>
            """,
            unsafe_allow_html=True
        )


    if total_saida_hoje > 0:

        st.markdown(
            f"""
            <div class="alert-box alert-yellow">
                🟠 <b>Operação de hoje:</b>
                <b>{total_saida_hoje}</b> carga(s)
                possuem saída prevista para hoje.
            </div>
            """,
            unsafe_allow_html=True
        )


    if total_entrega_hoje > 0:

        st.markdown(
            f"""
            <div class="alert-box alert-yellow">
                📦 <b>Entregas de hoje:</b>
                <b>{total_entrega_hoje}</b> carga(s)
                possuem entrega prevista para hoje.
            </div>
            """,
            unsafe_allow_html=True
        )


    if (
        total_atrasadas == 0
        and total_saida_hoje == 0
        and total_entrega_hoje == 0
        and total_cargas > 0
    ):

        st.markdown(
            """
            <div class="alert-box alert-green">
                🟢 <b>Operação normal:</b>
                nenhuma carga atrasada ou com alerta para hoje.
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # FILTROS
    # ========================================================

    st.markdown("### 🔎 Filtros da Operação")

    col_f1, col_f2, col_f3, col_f4 = st.columns(
        [2, 1.5, 1.5, 2]
    )


    with col_f1:

        motoristas_filtro_opcoes = [
            "Todos os Motoristas"
        ] + motoristas_lista

        motorista_selecionado = st.selectbox(
            "Filtrar por Motorista",
            motoristas_filtro_opcoes
        )


    with col_f2:

        data_inicial_filtro = st.date_input(
            "Data Inicial (Saída)",
            value=(
                datetime.date.today()
                - datetime.timedelta(days=7)
            )
        )


    with col_f3:

        data_final_filtro = st.date_input(
            "Data Final (Saída)",
            value=(
                datetime.date.today()
                + datetime.timedelta(days=30)
            )
        )


    with col_f4:

        pesquisa = st.text_input(
            "🔎 Pesquisar",
            placeholder="ID, motorista ou destino..."
        )


    # ========================================================
    # FILTRAGEM
    # ========================================================

    cargas_filtradas_periodo = []

    for c in cargas_lista:

        data_str = (
            c.get("data_saida")
            or c.get("data_carga")
        )

        incluir = True

        if data_str:

            dt_obj = converter_para_data(
                data_str
            )

            if dt_obj:

                if not (
                    data_inicial_filtro
                    <= dt_obj
                    <= data_final_filtro
                ):

                    incluir = False


        if incluir:

            cargas_filtradas_periodo.append(c)


    if motorista_selecionado != "Todos os Motoristas":

        cargas_filtradas_periodo = [
            c
            for c in cargas_filtradas_periodo
            if c.get("motorista")
            == motorista_selecionado
        ]


    if pesquisa:

        termo = pesquisa.lower().strip()

        cargas_filtradas_periodo = [
            c
            for c in cargas_filtradas_periodo

            if (
                termo
                in str(c.get("id", "")).lower()
                or termo
                in str(c.get("motorista", "")).lower()
                or termo
                in str(c.get("destino", "")).lower()
                or termo
                in str(c.get("observacoes", "")).lower()
            )
        ]


    st.caption(
        f"🔎 {len(cargas_filtradas_periodo)} "
        f"carga(s) encontrada(s) com os filtros atuais."
    )


    # ========================================================
    # STATUS
    # ========================================================

    colunas_status = [
        "Aguardando Carregamento",
        "Carregado / No Pátio",
        "Em Trânsito / Viagem Iniciada",
        "Entregue / Concluído",
    ]


    paleta_cores = [
        "#58a6ff",
        "#3fb950",
        "#d29922",
        "#bc8cff",
        "#f85149",
        "#39c5bb",
        "#f0883e",
        "#db61a2"
    ]


    mapa_cores = {
        mot: paleta_cores[
            i % len(paleta_cores)
        ]
        for i, mot
        in enumerate(motoristas_lista)
    }


    # ========================================================
    # KANBAN
    # ========================================================

    cols = st.columns(
        len(colunas_status)
    )


    for idx, status in enumerate(
        colunas_status
    ):

        with cols[idx]:

            quantidade_status = sum(
                1
                for c in cargas_filtradas_periodo
                if c.get("status") == status
            )

            st.markdown(
                f"""
                <div class='kanban-header'>
                    {status}
                    <br>
                    <span style="
                        font-size:11px;
                        color:#8b949e;
                    ">
                        {quantidade_status} carga(s)
                    </span>
                </div>
                """,
                unsafe_allow_html=True
            )


            cargas_status_filtradas = [
                c
                for c in cargas_filtradas_periodo
                if c.get("status") == status
            ]


            for carga in cargas_status_filtradas:

                carga_id = carga.get("id")

                motorista_atual = carga.get(
                    "motorista",
                    ""
                )

                cor_motorista = mapa_cores.get(
                    motorista_atual,
                    "#8b949e"
                )

                saida_br = formatar_data_br(
                    carga.get("data_saida")
                )

                entrega_br = formatar_data_br(
                    carga.get("data_entrega")
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


                atrasada = carga_atrasada(
                    carga
                )

                saida_hoje = carga_saida_hoje(
                    carga
                )

                entrega_hoje = carga_entrega_hoje(
                    carga
                )

                prazo_texto = texto_prazo(
                    carga
                )


                # ==================================================
                # CARD
                # ==================================================

                with st.container():

                    # Cor lateral do card
                    cor_borda = (
                        "#f85149"
                        if atrasada
                        else cor_motorista
                    )


                    c_info, c_btn = st.columns(
                        [3.5, 2.5]
                    )


                    with c_info:

                        badges = ""

                        if atrasada:

                            badges += (
                                '<span class="badge badge-red">'
                                '🔴 ATRASADA'
                                '</span>'
                            )

                        elif entrega_hoje:

                            badges += (
                                '<span class="badge badge-yellow">'
                                '📦 ENTREGA HOJE'
                                '</span>'
                            )

                        elif saida_hoje:

                            badges += (
                                '<span class="badge badge-yellow">'
                                '🚚 SAI HOJE'
                                '</span>'
                            )

                        else:

                            badges += (
                                '<span class="badge badge-green">'
                                '✓ NO PRAZO'
                                '</span>'
                            )


                        if prazo_texto:

                            badges += (
                                f'<span class="badge badge-blue">'
                                f'{prazo_texto}'
                                f'</span>'
                            )


                        observacoes = str(
                            carga.get(
                                "observacoes",
                                ""
                            )
                        ).strip()


                        ajudantes_html = ""

                        if ajudantes_texto.strip():

                            ajudantes_html = (
                                f"""
                                <br>
                                <span style="
                                    font-size:11px;
                                    color:#8b949e;
                                ">
                                    👥 Ajudantes:
                                </span>
                                <span style="
                                    font-size:11px;
                                    color:#c9d1d9;
                                ">
                                    {ajudantes_texto}
                                </span>
                                """
                            )


                        observacoes_html = ""

                        if observacoes:

                            observacoes_html = (
                                f"""
                                <br>
                                <span style="
                                    font-size:11px;
                                    color:#8b949e;
                                ">
                                    📝 Obs.:
                                </span>
                                <span style="
                                    font-size:11px;
                                    color:#c9d1d9;
                                ">
                                    {observacoes}
                                </span>
                                """
                            )


                        st.markdown(
                            f"""
                            <div style="
                                border-left:4px solid {cor_borda};
                                padding-left:8px;
                                margin-bottom:2px;
                            ">

                                <div style="
                                    margin-bottom:5px;
                                ">
                                    {badges}
                                </div>

                                <b style="
                                    font-size:13px;
                                    color:#58a6ff;
                                ">
                                    📌 ID / Planejamento:
                                    {carga_id}
                                </b>

                                <br>

                                <b style="
                                    font-size:14px;
                                    color:#ffffff;
                                ">
                                    🚚 {motorista_atual}
                                </b>

                                <br>

                                <span style="
                                    font-size:11px;
                                    color:#8b949e;
                                ">
                                    Destino:
                                </span>

                                <span style="
                                    color:#c9d1d9;
                                    font-weight:500;
                                    font-size:13px;
                                ">
                                    {carga.get('destino', '')}
                                </span>

                                <br>

                                <span style="
                                    font-size:11px;
                                    color:#8b949e;
                                ">
                                    📅 Saída:
                                </span>

                                <span style="
                                    font-size:11px;
                                    color:#c9d1d9;
                                ">
                                    {saida_br}
                                </span>

                                <span style="
                                    font-size:11px;
                                    color:#8b949e;
                                ">
                                    &nbsp;|&nbsp;
                                    Entrega:
                                </span>

                                <span style="
                                    font-size:11px;
                                    color:#c9d1d9;
                                ">
                                    {entrega_br}
                                </span>

                                {ajudantes_html}

                                {observacoes_html}

                            </div>
                            """,
                            unsafe_allow_html=True
                        )


                    # ==================================================
                    # BOTÕES
                    # ==================================================

                    with c_btn:

                        col_e, col_d = st.columns(
                            2
                        )


                        with col_e:

                            if st.button(
                                "✏️",
                                key=f"btn_edit_{carga_id}",
                                help="Editar Carga",
                                use_container_width=True
                            ):

                                st.session_state[
                                    f"editando_{carga_id}"
                                ] = not st.session_state.get(
                                    f"editando_{carga_id}",
                                    False
                                )


                        with col_d:

                            if st.button(
                                "🗑️",
                                key=f"btn_del_{carga_id}",
                                help="Excluir Carga",
                                use_container_width=True
                            ):

                                st.session_state[
                                    f"confirmar_exclusao_{carga_id}"
                                ] = True


                    # ==================================================
                    # CONFIRMAÇÃO EXCLUSÃO
                    # ==================================================

                    if st.session_state.get(
                        f"confirmar_exclusao_{carga_id}",
                        False
                    ):

                        st.warning(
                            f"Tem certeza que deseja excluir "
                            f"a carga/planejamento {carga_id}?"
                        )

                        cx1, cx2 = st.columns(2)

                        with cx1:

                            if st.button(
                                "Sim, excluir",
                                key=f"confirm_del_{carga_id}",
                                type="primary"
                            ):

                                sucesso = deletar_documento(
                                    "cargas",
                                    carga_id
                                )

                                if sucesso:

                                    st.session_state[
                                        "cargas"
                                    ] = [
                                        c
                                        for c in cargas_lista
                                        if c.get("id")
                                        != carga_id
                                    ]

                                    st.session_state[
                                        f"confirmar_exclusao_{carga_id}"
                                    ] = False

                                    st.success(
                                        "Carga excluída."
                                    )

                                    st.rerun()


                        with cx2:

                            if st.button(
                                "Cancelar",
                                key=f"cancel_del_{carga_id}"
                            ):

                                st.session_state[
                                    f"confirmar_exclusao_{carga_id}"
                                ] = False

                                st.rerun()


                    # ==================================================
                    # STATUS
                    # ==================================================

                    novo_status = st.selectbox(
                        "Mover Status",
                        colunas_status,
                        index=(
                            colunas_status.index(status)
                            if status in colunas_status
                            else 0
                        ),
                        key=f"status_{carga_id}",
                    )


                    if novo_status != carga.get(
                        "status"
                    ):

                        carga["status"] = novo_status

                        if (
                            novo_status
                            == "Entregue / Concluído"
                            and not carga.get("data_entrega")
                        ):

                            carga["data_entrega"] = str(
                                datetime.date.today()
                            )


                        sucesso = salvar_documento(
                            "cargas",
                            carga_id,
                            carga
                        )

                        if sucesso:

                            st.session_state[
                                f"editando_{carga_id}"
                            ] = False

                            st.rerun()


                    # ==================================================
                    # FORMULÁRIO DE EDIÇÃO
                    # ==================================================

                    if st.session_state.get(
                        f"editando_{carga_id}",
                        False
                    ):

                        with st.form(
                            key=f"form_edit_{carga_id}"
                        ):

                            st.markdown(
                                f"**✏️ Editando Planejamento #{carga_id}**"
                            )


                            mot_idx = (
                                motoristas_lista.index(
                                    carga.get("motorista")
                                )
                                if carga.get("motorista")
                                in motoristas_lista
                                else 0
                            )


                            novo_mot = st.selectbox(
                                "Motorista",
                                motoristas_lista
                                if motoristas_lista
                                else [""],
                                index=mot_idx
                            )


                            novo_dest = st.text_input(
                                "Destino",
                                value=carga.get(
                                    "destino",
                                    ""
                                )
                            )


                            novo_obs = st.text_area(
                                "Observações / Rota",
                                value=carga.get(
                                    "observacoes",
                                    ""
                                )
                            )


                            ajudantes_existentes = carga.get(
                                "ajudantes",
                                []
                            )

                            if not isinstance(
                                ajudantes_existentes,
                                list
                            ):

                                ajudantes_existentes = []


                            ajudantes_editados = st.multiselect(
                                "Ajudantes",
                                ajudantes_lista,
                                default=[
                                    a
                                    for a in ajudantes_existentes
                                    if a in ajudantes_lista
                                ]
                            )


                            dt_saida_val = (
                                converter_para_data(
                                    carga.get(
                                        "data_saida"
                                    )
                                )
                                or datetime.date.today()
                            )


                            dt_ent_val = (
                                converter_para_data(
                                    carga.get(
                                        "data_entrega"
                                    )
                                )
                                or datetime.date.today()
                            )


                            nova_saida = st.date_input(
                                "Data Saída",
                                value=dt_saida_val,
                                key=f"saida_{carga_id}"
                            )


                            nova_entrega = st.date_input(
                                "Data Entrega",
                                value=dt_ent_val,
                                key=f"entrega_{carga_id}"
                            )


                            salvar_edicao = st.form_submit_button(
                                "💾 Salvar Alterações"
                            )


                            if salvar_edicao:

                                carga["motorista"] = novo_mot
                                carga["destino"] = novo_dest
                                carga["observacoes"] = novo_obs
                                carga["ajudantes"] = (
                                    ajudantes_editados
                                )
                                carga["data_saida"] = str(
                                    nova_saida
                                )
                                carga["data_entrega"] = str(
                                    nova_entrega
                                )


                                sucesso = salvar_documento(
                                    "cargas",
                                    carga_id,
                                    carga
                                )


                                if sucesso:

                                    st.session_state[
                                        f"editando_{carga_id}"
                                    ] = False

                                    st.success(
                                        "Carga atualizada com sucesso!"
                                    )

                                    st.rerun()


# ============================================================
# 2. NOVA CARGA
# ============================================================

elif menu == "➕ Nova Carga":

    st.subheader(
        "➕ Cadastrar Novo Agendamento de Carga"
    )

    st.info(
        "📌 O número do planejamento é fornecido pelo "
        "seu sistema de montagem de cargas. "
        "Digite exatamente o número recebido."
    )


    with st.form(
        "form_nova_carga"
    ):

        col_id_manual, _ = st.columns(
            [2, 2]
        )


        with col_id_manual:

            id_planejamento = st.text_input(
                "Número do Planejamento / ID da Carga",
                placeholder="Ex: 1042"
            )


        col1, col2 = st.columns(2)


        with col1:

            motorista = st.selectbox(
                "Motorista Responsável",
                motoristas_lista
                if motoristas_lista
                else ["Nenhum cadastrado"]
            )


            destino = st.text_input(
                "Região / Cidades de Destino",
                placeholder=(
                    "Ex: Uberaba, Araxá "
                    "(Múltiplas entregas)"
                )
            )


            observacoes = st.text_area(
                "Observações / Rota",
                placeholder=(
                    "Ex: Carga com entregas "
                    "em lojas diferentes"
                )
            )


        with col2:

            ajudantes = st.multiselect(
                "Ajudantes da Viagem",
                ajudantes_lista
            )


            data_carga = st.date_input(
                "Data do Carregamento"
            )


            data_saida = st.date_input(
                "Data de Saída"
            )


            data_entrega = st.date_input(
                "Data Prevista de Entrega"
            )


        status_inicial = st.selectbox(
            "Status Inicial",
            [
                "Aguardando Carregamento",
                "Carregado / No Pátio",
                "Em Trânsito / Viagem Iniciada",
                "Entregue / Concluído",
            ],
        )


        submit = st.form_submit_button(
            "💾 Salvar e Agendar Carga",
            type="primary"
        )


        if submit:

            id_planejamento = str(
                id_planejamento
            ).strip()


            if (
                id_planejamento
                and destino
                and motorista
            ):

                ids_existentes = [
                    str(c.get("id"))
                    for c in cargas_lista
                ]


                if (
                    id_planejamento
                    in ids_existentes
                ):

                    st.error(
                        f"Já existe uma carga cadastrada "
                        f"com o ID/Planejamento "
                        f"'{id_planejamento}'. "
                        f"Use o número correto ou verifique "
                        f"se esta carga já foi cadastrada."
                    )

                else:

                    nova_carga = {
                        "id": id_planejamento,
                        "motorista": motorista,
                        "destino": destino,
                        "observacoes": observacoes,
                        "ajudantes": ajudantes,
                        "data_carga": str(data_carga),
                        "data_saida": str(data_saida),
                        "data_entrega": str(data_entrega),
                        "status": status_inicial,
                    }


                    sucesso = salvar_documento(
                        "cargas",
                        id_planejamento,
                        nova_carga
                    )


                    if sucesso:

                        st.session_state[
                            "cargas"
                        ].append(
                            nova_carga
                        )

                        st.success(
                            f"✅ Carga/Planejamento "
                            f"{id_planejamento} cadastrada "
                            f"com sucesso!"
                        )

                        st.rerun()

            else:

                st.error(
                    "Preencha o Número do Planejamento, "
                    "o Motorista e a Região de Destino."
                )


# ============================================================
# 3. CADASTROS
# ============================================================

elif menu == "👥 Cadastros (Equipe)":

    st.subheader(
        "👥 Gerenciamento de Motoristas e Ajudantes"
    )


    col1, col2 = st.columns(2)


    # ========================================================
    # MOTORISTAS
    # ========================================================

    with col1:

        st.markdown("### 🚚 Motoristas")


        with st.form(
            "form_cad_mot",
            clear_on_submit=True
        ):

            novo_mot = st.text_input(
                "Adicionar novo motorista"
            )


            cad_mot_btn = st.form_submit_button(
                "Cadastrar Motorista"
            )


            if cad_mot_btn and novo_mot:

                novo_mot = novo_mot.strip()


                nomes_existentes = [
                    str(m.get("nome", "")).lower()
                    for m
                    in st.session_state["motoristas"]
                ]


                if novo_mot.lower() in nomes_existentes:

                    st.error(
                        "Este motorista já está cadastrado."
                    )

                else:

                    doc_id = (
                        f"mot_"
                        f"{int(datetime.datetime.now().timestamp())}"
                    )


                    dados_mot = {
                        "id": doc_id,
                        "nome": novo_mot
                    }


                    sucesso = salvar_documento(
                        "motoristas",
                        doc_id,
                        dados_mot
                    )


                    if sucesso:

                        st.session_state[
                            "motoristas"
                        ].append(
                            dados_mot
                        )

                        st.success(
                            f"Motorista {novo_mot} adicionado!"
                        )

                        st.rerun()


        st.markdown("---")

        st.write(
            "**Motoristas Atuais:**"
        )


        if not st.session_state["motoristas"]:

            st.info(
                "Nenhum motorista cadastrado."
            )


        for m_obj in st.session_state[
            "motoristas"
        ]:

            m_nome = m_obj.get(
                "nome",
                ""
            )

            m_id = m_obj.get(
                "id",
                m_nome
            )


            c_mot1, c_mot2 = st.columns(
                [4, 2]
            )


            c_mot1.write(
                f"🚚 {m_nome}"
            )


            if c_mot2.button(
                "Excluir",
                key=f"del_mot_{m_id}"
            ):

                st.session_state[
                    f"confirmar_mot_{m_id}"
                ] = True


            if st.session_state.get(
                f"confirmar_mot_{m_id}",
                False
            ):

                st.warning(
                    f"Excluir o motorista "
                    f"'{m_nome}'?"
                )


                cm1, cm2 = st.columns(2)


                with cm1:

                    if st.button(
                        "Confirmar",
                        key=f"conf_mot_{m_id}"
                    ):

                        sucesso = deletar_documento(
                            "motoristas",
                            m_id
                        )


                        if sucesso:

                            st.session_state[
                                "motoristas"
                            ] = [
                                m
                                for m
                                in st.session_state[
                                    "motoristas"
                                ]
                                if m.get("id")
                                != m_id
                            ]


                            st.rerun()


                with cm2:

                    if st.button(
                        "Cancelar",
                        key=f"canc_mot_{m_id}"
                    ):

                        st.session_state[
                            f"confirmar_mot_{m_id}"
                        ] = False

                        st.rerun()


    # ========================================================
    # AJUDANTES
    # ========================================================

    with col2:

        st.markdown("### 👥 Ajudantes")


        with st.form(
            "form_cad_aju",
            clear_on_submit=True
        ):

            novo_aju = st.text_input(
                "Adicionar novo ajudante"
            )


            cad_aju_btn = st.form_submit_button(
                "Cadastrar Ajudante"
            )


            if cad_aju_btn and novo_aju:

                novo_aju = novo_aju.strip()


                nomes_existentes = [
                    str(a.get("nome", "")).lower()
                    for a
                    in st.session_state["ajudantes"]
                ]


                if novo_aju.lower() in nomes_existentes:

                    st.error(
                        "Este ajudante já está cadastrado."
                    )

                else:

                    doc_id = (
                        f"aju_"
                        f"{int(datetime.datetime.now().timestamp())}"
                    )


                    dados_aju = {
                        "id": doc_id,
                        "nome": novo_aju
                    }


                    sucesso = salvar_documento(
                        "ajudantes",
                        doc_id,
                        dados_aju
                    )


                    if sucesso:

                        st.session_state[
                            "ajudantes"
                        ].append(
                            dados_aju
                        )

                        st.success(
                            f"Ajudante {novo_aju} adicionado!"
                        )

                        st.rerun()


        st.markdown("---")

        st.write(
            "**Ajudantes Atuais:**"
        )


        if not st.session_state["ajudantes"]:

            st.info(
                "Nenhum ajudante cadastrado."
            )


        for a_obj in st.session_state[
            "ajudantes"
        ]:

            a_nome = a_obj.get(
                "nome",
                ""
            )

            a_id = a_obj.get(
                "id",
                a_nome
            )


            c_aju1, c_aju2 = st.columns(
                [4, 2]
            )


            c_aju1.write(
                f"👤 {a_nome}"
            )


            if c_aju2.button(
                "Excluir",
                key=f"del_aju_{a_id}"
            ):

                st.session_state[
                    f"confirmar_aju_{a_id}"
                ] = True


            if st.session_state.get(
                f"confirmar_aju_{a_id}",
                False
            ):

                st.warning(
                    f"Excluir o ajudante "
                    f"'{a_nome}'?"
                )


                ca1, ca2 = st.columns(2)


                with ca1:

                    if st.button(
                        "Confirmar",
                        key=f"conf_aju_{a_id}"
                    ):

                        sucesso = deletar_documento(
                            "ajudantes",
                            a_id
                        )


                        if sucesso:

                            st.session_state[
                                "ajudantes"
                            ] = [
                                a
                                for a
                                in st.session_state[
                                    "ajudantes"
                                ]
                                if a.get("id")
                                != a_id
                            ]


                            st.rerun()


                with ca2:

                    if st.button(
                        "Cancelar",
                        key=f"canc_aju_{a_id}"
                    ):

                        st.session_state[
                            f"confirmar_aju_{a_id}"
                        ] = False

                        st.rerun()


# ============================================================
# 4. RELATÓRIOS
# ============================================================

elif menu == "📈 Relatórios":

    st.subheader(
        "📈 Relatórios e Exportação de Dados"
    )


    if not cargas_lista:

        st.info(
            "Nenhuma carga cadastrada "
            "para gerar relatórios."
        )

    else:

        # ====================================================
        # FILTROS
        # ====================================================

        st.markdown(
            "### 🔎 Filtros do Relatório"
        )


        rf1, rf2, rf3 = st.columns(3)


        with rf1:

            filtro_motorista = st.selectbox(
                "Motorista",
                ["Todos"] + motoristas_lista,
                key="rel_motorista"
            )


        with rf2:

            filtro_status = st.selectbox(
                "Status",
                [
                    "Todos",
                    "Aguardando Carregamento",
                    "Carregado / No Pátio",
                    "Em Trânsito / Viagem Iniciada",
                    "Entregue / Concluído",
                ],
                key="rel_status"
            )


        with rf3:

            filtro_busca = st.text_input(
                "Pesquisar ID ou destino",
                placeholder="Ex: 1042 ou Uberaba",
                key="rel_busca"
            )


        cargas_relatorio = list(
            cargas_lista
        )


        if filtro_motorista != "Todos":

            cargas_relatorio = [
                c
                for c in cargas_relatorio
                if c.get("motorista")
                == filtro_motorista
            ]


        if filtro_status != "Todos":

            cargas_relatorio = [
                c
                for c in cargas_relatorio
                if c.get("status")
                == filtro_status
            ]


        if filtro_busca:

            termo = filtro_busca.lower().strip()


            cargas_relatorio = [
                c
                for c in cargas_relatorio

                if (
                    termo
                    in str(
                        c.get("id", "")
                    ).lower()

                    or

                    termo
                    in str(
                        c.get("destino", "")
                    ).lower()
                )
            ]


        # ====================================================
        # INDICADORES
        # ====================================================

        total_relatorio = len(
            cargas_relatorio
        )


        entregues_relatorio = sum(
            1
            for c in cargas_relatorio
            if c.get("status")
            == "Entregue / Concluído"
        )


        transito_relatorio = sum(
            1
            for c in cargas_relatorio
            if c.get("status")
            == "Em Trânsito / Viagem Iniciada"
        )


        atrasadas_relatorio = sum(
            1
            for c in cargas_relatorio
            if carga_atrasada(c)
        )


        percentual_entregues = (
            (
                entregues_relatorio
                / total_relatorio
            )
            * 100
            if total_relatorio
            else 0
        )


        r1, r2, r3, r4, r5 = st.columns(5)


        with r1:

            st.markdown(
                f"""
                <div class="metric-card metric-blue">
                    <div class="metric-title">
                        CARGAS
                    </div>
                    <div class="metric-value">
                        {total_relatorio}
                    </div>
                    <div class="metric-subtitle">
                        Resultado do filtro
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


        with r2:

            st.markdown(
                f"""
                <div class="metric-card metric-green">
                    <div class="metric-title">
                        ENTREGUES
                    </div>
                    <div class="metric-value">
                        {entregues_relatorio}
                    </div>
                    <div class="metric-subtitle">
                        Concluídas
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


        with r3:

            st.markdown(
                f"""
                <div class="metric-card metric-yellow">
                    <div class="metric-title">
                        EM TRÂNSITO
                    </div>
                    <div class="metric-value">
                        {transito_relatorio}
                    </div>
                    <div class="metric-subtitle">
                        Viagens em andamento
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


        with r4:

            st.markdown(
                f"""
                <div class="metric-card metric-red">
                    <div class="metric-title">
                        ATRASADAS
                    </div>
                    <div class="metric-value">
                        {atrasadas_relatorio}
                    </div>
                    <div class="metric-subtitle">
                        Fora do prazo
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


        with r5:

            st.markdown(
                f"""
                <div class="metric-card metric-purple">
                    <div class="metric-title">
                        CONCLUSÃO
                    </div>
                    <div class="metric-value">
                        {percentual_entregues:.1f}%
                    </div>
                    <div class="metric-subtitle">
                        Taxa de entregas
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


        st.markdown("---")


        # ====================================================
        # GRÁFICOS
        # ====================================================

        st.markdown(
            "### 📊 Visão Gerencial"
        )


        gc1, gc2 = st.columns(2)


        with gc1:

            st.markdown(
                "**Cargas por Status**"
            )


            status_contagem = {}

            for c in cargas_relatorio:

                status = c.get(
                    "status",
                    "Sem Status"
                )

                status_contagem[status] = (
                    status_contagem.get(
                        status,
                        0
                    ) + 1
                )


            if status_contagem:

                df_status = pd.DataFrame(
                    {
                        "Status": list(
                            status_contagem.keys()
                        ),
                        "Quantidade": list(
                            status_contagem.values()
                        )
                    }
                )

                st.bar_chart(
                    df_status.set_index(
                        "Status"
                    )
                )


        with gc2:

            st.markdown(
                "**Cargas por Motorista**"
            )


            motorista_contagem = {}


            for c in cargas_relatorio:

                motorista = c.get(
                    "motorista",
                    "Sem motorista"
                )

                motorista_contagem[motorista] = (
                    motorista_contagem.get(
                        motorista,
                        0
                    ) + 1
                )


            if motorista_contagem:

                df_motoristas = pd.DataFrame(
                    {
                        "Motorista": list(
                            motorista_contagem.keys()
                        ),
                        "Quantidade": list(
                            motorista_contagem.values()
                        )
                    }
                )

                st.bar_chart(
                    df_motoristas.set_index(
                        "Motorista"
                    )
                )


        # ====================================================
        # TABELA
        # ====================================================

        st.markdown(
            "### 📋 Detalhamento das Cargas"
        )


        df_tabela = preparar_dataframe(
            cargas_relatorio
        )


        if df_tabela.empty:

            st.info(
                "Nenhuma carga encontrada "
                "com os filtros selecionados."
            )

        else:

            st.dataframe(
                df_tabela,
                use_container_width=True,
                hide_index=True
            )


        # ====================================================
        # EXPORTAÇÃO
        # ====================================================

        st.markdown(
            "### 📥 Exportar Arquivos"
        )


        col_exp1, col_exp2 = st.columns(2)


        with col_exp1:

            excel_data = gerar_excel_profissional(
                df_tabela
            )


            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=excel_data,
                file_name="relatorio_de_cargas.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )


        with col_exp2:

            pdf_bytes = gerar_pdf(
                df_tabela
            )


            st.download_button(
                label="📄 Baixar Relatório em PDF",
                data=pdf_bytes,
                file_name="relatorio_de_cargas.pdf",
                mime="application/pdf",
                use_container_width=True
            )
