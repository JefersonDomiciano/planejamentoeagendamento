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
        .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}

        .app-kicker {color:#8ea0b8; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1.4px; margin-bottom:2px;}
        .app-title {color:#f8fafc; font-size:28px; font-weight:800; letter-spacing:-.5px; margin-bottom:0;}
        .app-subtitle {color:#8796aa; font-size:13px; margin-top:3px;}

        .kanban-header {text-align:left; background:linear-gradient(135deg,#182235 0%,#111827 100%); color:#f8fafc!important; padding:13px 14px; border-radius:12px; font-weight:750; font-size:13px; border:1px solid #273449; margin-bottom:10px; box-shadow:0 8px 24px rgba(0,0,0,.14);}
        .kanban-count {color:#71819a!important; font-size:11px; font-weight:600;}
        .kanban-card {background:linear-gradient(145deg,#172131 0%,#111827 100%); border:1px solid #263449; border-radius:12px; padding:11px 12px; margin:0 0 10px 0; box-shadow:0 10px 26px rgba(0,0,0,.16);}
        .card-label {color:#73849a; font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.7px;}
        .card-id {color:#60a5fa; font-size:12px; font-weight:800;}
        .card-driver {color:#f8fafc; font-size:13px; font-weight:800; margin:7px 0 8px;}
        .card-destination {color:#d6deea; font-size:11px; font-weight:600;}
        .card-meta {color:#a5b3c5; font-size:10px; line-height:1.7;}
        .card-divider {height:1px; background:#243145; margin:10px 0;}

        .metric-card {position:relative; overflow:hidden; background:linear-gradient(145deg,#172131 0%,#111827 100%); border:1px solid #263449; border-radius:12px; padding:13px 15px; min-height:92px; box-shadow:0 10px 28px rgba(0,0,0,.14);}
        .metric-card::after {content:""; position:absolute; right:-28px; top:-35px; width:90px; height:90px; border-radius:50%; background:rgba(255,255,255,.025);}
        .metric-title {color:#8998ac; font-size:10px; font-weight:800; letter-spacing:.8px;}
        .metric-value {color:#f8fafc; font-size:25px; font-weight:850; line-height:1.1; margin-top:7px;}
        .metric-subtitle {color:#718198; font-size:10px; margin-top:6px;}
        .metric-blue{border-left:3px solid #3b82f6;} .metric-green{border-left:3px solid #22c55e;} .metric-yellow{border-left:3px solid #f59e0b;} .metric-purple{border-left:3px solid #a855f7;} .metric-red{border-left:3px solid #ef4444;}

        .alert-box {padding:12px 15px; border-radius:11px; margin-bottom:10px; font-size:12px; font-weight:600;}
        .alert-red {background:rgba(239,68,68,.08); border:1px solid rgba(239,68,68,.28); color:#fca5a5;}
        .alert-yellow {background:rgba(245,158,11,.08); border:1px solid rgba(245,158,11,.28); color:#fcd34d;}
        .alert-green {background:rgba(34,197,94,.08); border:1px solid rgba(34,197,94,.25); color:#86efac;}

        .badge {display:inline-block; padding:4px 8px; border-radius:999px; font-size:9px; font-weight:800; letter-spacing:.25px; margin-right:4px; margin-bottom:5px;}
        .badge-red {background:rgba(239,68,68,.12); color:#fca5a5; border:1px solid rgba(239,68,68,.24);}
        .badge-yellow {background:rgba(245,158,11,.12); color:#fcd34d; border:1px solid rgba(245,158,11,.24);}
        .badge-green {background:rgba(34,197,94,.12); color:#86efac; border:1px solid rgba(34,197,94,.22);}
        .badge-blue {background:rgba(59,130,246,.12); color:#93c5fd; border:1px solid rgba(59,130,246,.22);}

        .section-title {color:#f1f5f9; font-size:18px; font-weight:800; margin:4px 0 14px;}
        .stSelectbox label,.stDateInput label,.stTextInput label,.stMultiSelect label,.stTextArea label {font-size:11px!important; color:#8ea0b8!important; font-weight:650!important;}
        div[data-testid="stDataFrame"] {border:1px solid #263449; border-radius:12px; overflow:hidden;}

        .chart-card {background:linear-gradient(145deg,#172131 0%,#111827 100%); border:1px solid #263449; border-radius:14px; padding:18px; margin:0 0 14px 0; box-shadow:0 10px 28px rgba(0,0,0,.14);}
        .chart-heading {display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:16px; color:#f1f5f9; font-size:14px; font-weight:800;}
        .chart-heading small {color:#718198; font-size:10px; font-weight:600;}
        .donut-layout {display:flex; align-items:center; gap:30px; min-height:230px;}
        .donut {width:190px; height:190px; border-radius:50%; display:flex; align-items:center; justify-content:center; flex-shrink:0;}
        .donut-hole {width:116px; height:116px; border-radius:50%; background:#111827; display:flex; flex-direction:column; align-items:center; justify-content:center; box-shadow:0 0 0 1px #263449 inset;}
        .donut-hole strong {font-size:28px; line-height:1; color:#f8fafc;}
        .donut-hole span {font-size:10px; color:#718198; margin-top:5px;}
        .legend-list {flex:1;} .legend-row {display:grid; grid-template-columns:12px 1fr auto; gap:8px; align-items:center; padding:8px 0; color:#cbd5e1; font-size:11px; border-bottom:1px solid rgba(148,163,184,.08);}
        .legend-row b {color:#f8fafc; font-size:12px;} .legend-dot {width:8px; height:8px; border-radius:50%; display:block;}
        .bar-list {display:flex; flex-direction:column; gap:13px;}
        .bar-row {display:grid; grid-template-columns:180px 1fr 34px; gap:10px; align-items:center;}
        .bar-label {overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#cbd5e1; font-size:11px;}
        .bar-track {height:9px; background:#202c3d; border-radius:999px; overflow:hidden;}
        .bar-fill {height:100%; border-radius:999px; background:linear-gradient(90deg,#3b82f6,#60a5fa); box-shadow:0 0 12px rgba(59,130,246,.18);}
        .bar-value {text-align:right; color:#f8fafc; font-size:12px; font-weight:800;}
        .chart-empty {padding:35px 10px; text-align:center; color:#718198; font-size:12px;}
        .evo-chart {height:220px; display:flex; align-items:flex-end; gap:10px; padding:12px 5px 0; border-bottom:1px solid #263449; overflow-x:auto;}
        .evo-item {height:100%; min-width:44px; display:flex; flex-direction:column; justify-content:flex-end; align-items:center; gap:6px;}
        .evo-value {color:#dbeafe; font-size:10px; font-weight:800;}
        .evo-bar {width:26px; min-height:8px; border-radius:7px 7px 2px 2px; background:linear-gradient(180deg,#60a5fa,#2563eb); box-shadow:0 4px 12px rgba(37,99,235,.18);}
        .evo-label {color:#718198; font-size:9px; white-space:nowrap;}

        @media (max-width:768px) {
            .block-container{padding-left:.65rem;padding-right:.65rem;padding-top:.6rem;}
            .app-title{font-size:22px;} .metric-card{min-height:96px;padding:13px;} .metric-value{font-size:24px;} .kanban-header{font-size:11px;padding:10px;}
        }
    </style>
    """,
    unsafe_allow_html=True
)


def render_html(html):
    """Renderiza HTML diretamente quando a versão do Streamlit oferece st.html."""
    if hasattr(st, "html"):
        st.html(html)
    else:
        st.markdown(html, unsafe_allow_html=True)


# ============================================================
# TÍTULO
# ============================================================

st.markdown(
    """
    <div class="app-kicker">LOGÍSTICA INTELIGENTE</div>
    <div class="app-title">🚚 Gestão de Cargas</div>
    <div class="app-subtitle">Torre de controle operacional e acompanhamento de planejamentos</div>
    """,
    unsafe_allow_html=True
)


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



def atualizar_campos_documento(colecao, doc_id, campos):
    """
    Atualiza somente os campos informados no documento.
    Usado para movimentar o status automaticamente sem botão Salvar Status.
    """
    url = (
        f"https://firestore.googleapis.com/v1/projects/"
        f"{FIREBASE_PROJECT_ID}/databases/(default)/documents/"
        f"{colecao}/{doc_id}"
    )

    fields = {}

    for k, v in campos.items():

        if isinstance(v, list):
            fields[k] = {
                "arrayValue": {
                    "values": [
                        {"stringValue": str(x)}
                        for x in v
                    ]
                }
            }

        elif isinstance(v, bool):
            fields[k] = {"booleanValue": v}

        elif isinstance(v, int):
            fields[k] = {"integerValue": str(v)}

        elif isinstance(v, float):
            fields[k] = {"doubleValue": v}

        elif v is None:
            fields[k] = {"nullValue": None}

        else:
            fields[k] = {"stringValue": str(v)}

    params = [
        ("updateMask.fieldPaths", campo)
        for campo in campos.keys()
    ]

    try:
        response = requests.patch(
            url,
            params=params,
            json={"fields": fields},
            timeout=10
        )

        if response.status_code not in [200, 201]:
            st.error(
                f"Erro ao atualizar no Firebase. "
                f"HTTP {response.status_code}"
            )
            return False

        return True

    except requests.exceptions.Timeout:
        st.error("Tempo esgotado ao atualizar o Firebase.")
        return False

    except requests.exceptions.RequestException as e:
        st.error(f"Erro de conexão com o Firebase: {e}")
        return False

    except Exception as e:
        st.error(f"Erro ao atualizar o Firebase: {e}")
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

    col_f1, col_f2 = st.columns([2, 3])

    with col_f1:

        motoristas_filtro_opcoes = [
            "Todos os Motoristas"
        ] + motoristas_lista

        motorista_selecionado = st.selectbox(
            "Filtrar por Motorista",
            motoristas_filtro_opcoes
        )

    with col_f2:

        pesquisa = st.text_input(
            "🔎 Pesquisar",
            placeholder="ID, motorista, destino ou observação..."
        )

    # ========================================================
    # FILTRAGEM
    # ========================================================

    # Sem filtro de período no painel.
    cargas_filtradas_periodo = list(cargas_lista)

    if motorista_selecionado != "Todos os Motoristas":

        cargas_filtradas_periodo = [
            c
            for c in cargas_filtradas_periodo
            if c.get("motorista") == motorista_selecionado
        ]

    if pesquisa:

        termo = pesquisa.lower().strip()

        cargas_filtradas_periodo = [
            c
            for c in cargas_filtradas_periodo
            if (
                termo in str(c.get("id", "")).lower()
                or termo in str(c.get("motorista", "")).lower()
                or termo in str(c.get("destino", "")).lower()
                or termo in str(c.get("observacoes", "")).lower()
                or termo in str(c.get("status", "")).lower()
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
                    <div>{status}</div>
                    <div class="kanban-count">{quantidade_status} carga(s)</div>
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
                            ajudantes_html = f'<div class="card-meta">👥 Ajudantes: <strong>{ajudantes_texto}</strong></div>'


                        observacoes_html = ""

                        if observacoes:
                            observacoes_html = f'<div class="card-meta">📝 Obs.: <strong>{observacoes}</strong></div>'
                        render_html(
                            f"""
                            <div class="kanban-card" style="border-left:3px solid {cor_borda};">
                                <div>{badges}</div>
                                <div class="card-id">📌 PLANEJAMENTO #{carga_id}</div>
                                <div class="card-driver">🚚 {motorista_atual}</div>
                                <div class="card-label">Destino</div>
                                <div class="card-destination">{carga.get('destino', '')}</div>
                                <div class="card-divider"></div>
                                <div class="card-meta">📅 Saída: <strong>{saida_br or '—'}</strong> &nbsp; • &nbsp; Entrega: <strong>{entrega_br or '—'}</strong></div>
                                {ajudantes_html}
                                {observacoes_html}
                            </div>
                            """
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
                    # MOVIMENTAÇÃO DO STATUS
                    # ==================================================
                    #
                    # Não existe botão "Salvar Status".
                    # Clicar em Voltar/Avançar salva o novo status
                    # automaticamente no Firebase.
                    # ==================================================

                    indice_status = (
                        colunas_status.index(status)
                        if status in colunas_status
                        else 0
                    )

                    mov1, mov2 = st.columns(2)

                    with mov1:

                        if indice_status > 0:

                            if st.button(
                                "⬅️ Voltar",
                                key=f"status_voltar_{carga_id}",
                                help="Voltar uma etapa",
                                use_container_width=True
                            ):

                                status_anterior = colunas_status[
                                    indice_status - 1
                                ]

                                campos_status = {
                                    "status": status_anterior
                                }

                                sucesso = atualizar_campos_documento(
                                    "cargas",
                                    carga_id,
                                    campos_status
                                )

                                if sucesso:

                                    carga["status"] = status_anterior

                                    st.session_state[
                                        f"editando_{carga_id}"
                                    ] = False

                                    st.rerun()

                        else:

                            st.button(
                                "⬅️ Voltar",
                                key=f"status_voltar_disabled_{carga_id}",
                                disabled=True,
                                use_container_width=True
                            )

                    with mov2:

                        if indice_status < len(colunas_status) - 1:

                            if st.button(
                                "Avançar ➡️",
                                key=f"status_avancar_{carga_id}",
                                help="Avançar uma etapa",
                                type="primary",
                                use_container_width=True
                            ):

                                proximo_status = colunas_status[
                                    indice_status + 1
                                ]

                                campos_status = {
                                    "status": proximo_status
                                }

                                # Registra a conclusão real sem alterar
                                # a data prevista de entrega.
                                if (
                                    proximo_status
                                    == "Entregue / Concluído"
                                ):
                                    campos_status[
                                        "data_conclusao"
                                    ] = str(
                                        datetime.date.today()
                                    )

                                sucesso = atualizar_campos_documento(
                                    "cargas",
                                    carga_id,
                                    campos_status
                                )

                                if sucesso:

                                    carga["status"] = proximo_status

                                    if (
                                        "data_conclusao"
                                        in campos_status
                                    ):
                                        carga[
                                            "data_conclusao"
                                        ] = campos_status[
                                            "data_conclusao"
                                        ]

                                    st.session_state[
                                        f"editando_{carga_id}"
                                    ] = False

                                    st.rerun()

                        else:

                            st.button(
                                "✅ Concluído",
                                key=f"status_concluido_{carga_id}",
                                disabled=True,
                                use_container_width=True
                            )


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

                    doc_id = f"mot_{int(datetime.datetime.now().timestamp() * 1000)}"


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

                    doc_id = f"aju_{int(datetime.datetime.now().timestamp() * 1000)}"


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
        # GRÁFICOS PROFISSIONAIS — SEM DEPENDÊNCIAS EXTERNAS
        # ====================================================

        st.markdown('<div class="section-title">📊 Visão Gerencial</div>', unsafe_allow_html=True)

        cores_status = {
            "Aguardando Carregamento": "#3b82f6",
            "Carregado / No Pátio": "#22c55e",
            "Em Trânsito / Viagem Iniciada": "#f59e0b",
            "Entregue / Concluído": "#a855f7",
            "Sem Status": "#64748b",
        }
        status_ordem = [
            "Aguardando Carregamento",
            "Carregado / No Pátio",
            "Em Trânsito / Viagem Iniciada",
            "Entregue / Concluído",
        ]

        status_contagem = {status: 0 for status in status_ordem}
        for c in cargas_relatorio:
            status = c.get("status", "Sem Status")
            status_contagem[status] = status_contagem.get(status, 0) + 1

        motorista_contagem = {}
        for c in cargas_relatorio:
            motorista = c.get("motorista", "Sem motorista") or "Sem motorista"
            motorista_contagem[motorista] = motorista_contagem.get(motorista, 0) + 1

        evolucao_contagem = {}
        for c in cargas_relatorio:
            data_obj = converter_para_data(c.get("data_saida") or c.get("data_carga"))
            if data_obj:
                evolucao_contagem[data_obj] = evolucao_contagem.get(data_obj, 0) + 1

        def grafico_status_html(contagem, total):
            partes = []
            inicio = 0.0
            for status in status_ordem:
                valor = contagem.get(status, 0)
                if not valor or not total:
                    continue
                fim = inicio + (valor / total) * 100
                partes.append(f"{cores_status[status]} {inicio:.2f}% {fim:.2f}%")
                inicio = fim
            if not partes:
                return '<div class="chart-empty">Nenhuma carga encontrada.</div>'
            gradient = ", ".join(partes)
            legenda = "".join(
                f'<div class="legend-row"><span class="legend-dot" style="background:{cores_status[s]}"></span><span>{s}</span><b>{contagem.get(s,0)}</b></div>'
                for s in status_ordem if contagem.get(s, 0)
            )
            return f"""
            <div class="chart-card">
              <div class="chart-heading"><span>Cargas por Status</span><small>Distribuição atual</small></div>
              <div class="donut-layout">
                <div class="donut" style="background:conic-gradient({gradient})">
                  <div class="donut-hole"><strong>{total}</strong><span>cargas</span></div>
                </div>
                <div class="legend-list">{legenda}</div>
              </div>
            </div>"""

        def grafico_barras_html(titulo, dados, limite=8):
            pares = sorted(dados.items(), key=lambda x: x[1], reverse=True)[:limite]
            if not pares:
                return f'<div class="chart-card"><div class="chart-heading"><span>{titulo}</span></div><div class="chart-empty">Nenhum dado disponível.</div></div>'
            maximo = max(v for _, v in pares) or 1
            linhas = ""
            for nome, valor in pares:
                largura = max(7, (valor / maximo) * 100)
                linhas += f"""
                <div class="bar-row">
                  <div class="bar-label" title="{nome}">{nome}</div>
                  <div class="bar-track"><div class="bar-fill" style="width:{largura:.1f}%"></div></div>
                  <div class="bar-value">{valor}</div>
                </div>"""
            return f"""
            <div class="chart-card">
              <div class="chart-heading"><span>{titulo}</span><small>Top {len(pares)}</small></div>
              <div class="bar-list">{linhas}</div>
            </div>"""

        st.markdown(grafico_status_html(status_contagem, total_relatorio), unsafe_allow_html=True)
        st.markdown(grafico_barras_html("Cargas por Motorista", motorista_contagem), unsafe_allow_html=True)

        if evolucao_contagem:
            datas = sorted(evolucao_contagem)
            max_val = max(evolucao_contagem.values()) or 1
            pontos = []
            for data_obj in datas:
                valor = evolucao_contagem[data_obj]
                altura = max(8, (valor / max_val) * 100)
                pontos.append(
                    f'<div class="evo-item"><div class="evo-value">{valor}</div><div class="evo-bar" style="height:{altura:.1f}%"></div><div class="evo-label">{data_obj.strftime("%d/%m")}</div></div>'
                )
            html_evo = f"""
            <div class="chart-card">
              <div class="chart-heading"><span>Movimentação por Data</span><small>Data de saída/carregamento</small></div>
              <div class="evo-chart">{"".join(pontos)}</div>
            </div>"""
            st.markdown(html_evo, unsafe_allow_html=True)

        st.markdown(grafico_barras_html("Top Motoristas", motorista_contagem, limite=5), unsafe_allow_html=True)

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
