import datetime
import io
import pandas as pd
import streamlit as st
import requests
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

        button.btn-kanban {
            width: 34px !important;
            height: 34px !important;
            min-height: 34px !important;
            max-height: 34px !important;
            padding: 0px !important;
            margin: 0 auto !important;
            background-color: #21262d !important;
            border: 1px solid #30363d !important;
            border-radius: 6px !important;
            font-size: 14px !important;
            line-height: 1 !important;
        }
        
        button.btn-kanban:hover {
            border-color: #58a6ff !important;
            background-color: #30363d !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- CONFIGURAÇÃO DA PLANILHA DO GOOGLE SHEETS ---
PLANILHA_ID = "COLE_O_ID_DA_ SUA_PLANILHA_AQUI"
url_cargas = f"https://docs.google.com/spreadsheets/d/{PLANILHA_ID}/gviz/tq?tqx=out:csv&sheet=cargas"
url_motoristas = f"https://docs.google.com/spreadsheets/d/{PLANILHA_ID}/gviz/tq?tqx=out:csv&sheet=motoristas"
url_ajudantes = f"https://docs.google.com/spreadsheets/d/{PLANILHA_ID}/gviz/tq?tqx=out:csv&sheet=ajudantes"

@st.cache_data(ttl=5)
def carregar_dados(colecao):
    try:
        if colecao == "cargas":
            df = pd.read_csv(url_cargas)
            return df.to_dict('records')
        elif colecao == "motoristas":
            df = pd.read_csv(url_motoristas)
            return df['nome'].dropna().tolist() if 'nome' in df.columns else ["Carlos Silva", "João Pereira", "Maurício", "Cícero Taveira"]
        elif colecao == "ajudantes":
            df = pd.read_csv(url_ajudantes)
            return df['nome'].dropna().tolist() if 'nome' in df.columns else ["Pedrinho", "Lucas Souza"]
    except Exception:
        if colecao == "motoristas":
            return ["Carlos Silva", "João Pereira", "Maurício", "Cícero Taveira"]
        elif colecao == "ajudantes":
            return ["Pedrinho", "Lucas Souza"]
        return []

def salvar_dado(colecao, dados, doc_id):
    st.info("Para salvar alterações permanentemente, gerencie diretamente na sua planilha aberta no Google Sheets.")

def adicionar_dado(colecao, dados, doc_id=None):
    st.info("Para cadastrar novos itens permanentemente, insira as linhas diretamente na sua planilha do Google Sheets.")

def excluir_dado(colecao, campo_filtro, valor):
    st.warning("A exclusão direta está desativada no modo de planilha pública. Gerencie os dados direto no Google Sheets.")

def formatar_data_br(data_str):
    if not data_str:
        return ""
    try:
        return datetime.date.fromisoformat(str(data_str)).strftime('%d/%m/%Y')
    except Exception:
        return str(data_str)

def preparar_dataframe(cargas_lista):
    df = pd.DataFrame(cargas_lista)
    if df.empty:
        return df
    
    if "ajudantes" in df.columns:
        df["ajudantes"] = df["ajudantes"].apply(lambda x: ", ".join(x) if isinstance(x, list) else str(x))
    
    for col in ["data_carga", "data_saida", "data_entrega"]:
        if col in df.columns:
            df[col] = df[col].apply(formatar_data_br)
    
    colunas_desejadas = ["id", "motorista", "destino", "observacoes", "ajudantes", "data_carga", "data_saida", "data_entrega", "status"]
    colunas_existentes = [col for col in colunas_desejadas if col in df.columns]
    df = df[colunas_existentes]
    return df

def gerar_excel_profissional(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatorio de Cargas"

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    data_font = Font(name="Arial", size=9)
    center_alignment = Alignment(horizontal="center", vertical="center")
    left_alignment = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = ["ID", "Motorista", "Destino", "Observações", "Ajudantes", "Data Carga", "Data Saída", "Data Entrega", "Status"]
    
    for col_num, header in enumerate(headers[:len(df.columns)], 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border
    
    ws.row_dimensions[1].height = 24

    for row_num, row_data in enumerate(df.values, 2):
        ws.row_dimensions[row_num].height = 20
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.value = value
            cell.font = data_font
            cell.border = thin_border
            if col_num in [1, 6, 7, 8]:
                cell.alignment = center_alignment
            else:
                cell.alignment = left_alignment

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

    wb.save(output)
    return output.getvalue()

def gerar_pdf(df):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.set_font("Arial", "B", 14)
    pdf.cell(190, 8, txt="Relatório de Cargas", ln=True, align="C")
    
    pdf.set_font("Arial", "", 9)
    pdf.cell(190, 5, txt=f"Data de geração: {datetime.date.today().strftime('%d/%m/%Y')}", ln=True, align="C")
    pdf.ln(4)

    pdf.set_font("Arial", "B", 8)
    pdf.set_fill_color(47, 117, 181)
    pdf.set_text_color(255, 255, 255)
    
    larguras = [10, 32, 38, 30, 25, 25, 30]
    nomes_colunas = ["ID", "Motorista", "Destino", "Ajudantes", "Saída", "Entrega", "Status"]
    
    for i, nome in enumerate(nomes_colunas):
        pdf.cell(larguras[i], 7, nome, 1, 0, "C", True)
    pdf.ln()

    pdf.set_font("Arial", "", 7)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        pdf.cell(larguras[0], 6, str(row.get("id", "")), 1, 0, "C")
        pdf.cell(larguras[1], 6, str(row.get("motorista", ""))[:18], 1, 0, "L")
        pdf.cell(larguras[2], 6, str(row.get("destino", ""))[:22], 1, 0, "L")
        pdf.cell(larguras[3], 6, str(row.get("ajudantes", ""))[:16], 1, 0, "L")
        pdf.cell(larguras[4], 6, str(row.get("data_saida", "")), 1, 0, "C")
        pdf.cell(larguras[5], 6, str(row.get("data_entrega", "")), 1, 0, "C")
        pdf.cell(larguras[6], 6, str(row.get("status", ""))[:15], 1, 0, "C")
        pdf.ln()

    return bytes(pdf.output(dest='S'))

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

st.info("💡 Conectado diretamente à planilha do Google Sheets. Visualize seu Kanban completo, cadastros e relatórios abaixo!")

menu = st.radio(
    "Menu Principal",
    [
        "📋 Painel (Kanban)",
        "➕ Nova Carga",
        "👥 Cadastros (Equipe)",
        "📊 Relatório Semanal",
    ],
    horizontal=True,
)

st.markdown("---")

motoristas_lista = carregar_dados("motoristas")
ajudantes_lista = carregar_dados("ajudantes")
cargas_lista = carregar_dados("cargas")

# ----------------------------------------------------
# 1. PAINEL (KANBAN)
# ----------------------------------------------------
if menu == "📋 Painel (Kanban)":
    st.subheader("Visão Geral das Cargas")

    col_f1, col_f2 = st.columns([2, 2])
    with col_f1:
        motoristas_filtro_opcoes = ["Todos os Motoristas"] + motoristas_lista
        motorista_selecionado = st.selectbox("Filtrar por Motorista", motoristas_filtro_opcoes)

    hoje = datetime.date.today()
    inicio_semana_atual = hoje - datetime.timedelta(days=hoje.weekday())
    fim_proxima_semana = inicio_semana_atual + datetime.timedelta(days=13)

    cargas_filtradas_periodo = []
    for c in cargas_lista:
        data_str = c.get("data_saida") or c.get("data_carga")
        if data_str:
            try:
                data_carga_obj = datetime.date.fromisoformat(str(data_str))
                if inicio_semana_atual <= data_carga_obj <= fim_proxima_semana:
                    cargas_filtradas_periodo.append(c)
            except Exception:
                cargas_filtradas_periodo.append(c)
        else:
            cargas_filtradas_periodo.append(c)

    if motorista_selecionado != "Todos os Motoristas":
        cargas_filtradas_periodo = [c for c in cargas_filtradas_periodo if c.get("motorista") == motorista_selecionado]

    colunas_status = [
        "Aguardando Carregamento",
        "Carregado / No Pátio",
        "Em Trânsito / Viagem Iniciada",
        "Entregue / Concluído",
    ]

    paleta_cores = ["#58a6ff", "#3fb950", "#d29922", "#bc8cff", "#f85149", "#39c5bb", "#f0883e", "#db61a2"]
    mapa_cores = {mot: paleta_cores[i % len(paleta_cores)] for i, mot in enumerate(motoristas_lista)}

    cols = st.columns(len(colunas_status))

    for idx, status in enumerate(colunas_status):
        with cols[idx]:
            st.markdown(
                f"<div class='kanban-header'>{status}</div>",
                unsafe_allow_html=True,
            )

            cargas_status_filtradas = [c for c in cargas_filtradas_periodo if str(c.get("status", "")).strip().lower() == status.lower()]

            if not cargas_status_filtradas:
                st.caption("Nenhuma carga neste status.")

            for carga in cargas_status_filtradas:
                carga_id = carga.get('id')
                motorista_atual = carga.get('motorista', '')
                cor_motorista = mapa_cores.get(motorista_atual, "#8b949e")

                saida_br = formatar_data_br(carga.get('data_saida'))
                entrega_br = formatar_data_br(carga.get('data_entrega'))

                with st.container():
                    st.markdown(f"""
                        <div style="border-left: 4px solid {cor_motorista}; padding-left: 8px; margin-bottom: 4px;">
                            <b style="font-size: 14px; color: #ffffff;">🚚 {motorista_atual}</b><br>
                            <span style="font-size: 13px; color: #8b949e;">Destino:</span> <span style="color: #c9d1d9; font-weight: 500;">{carga.get('destino')}</span><br>
                            <span style="font-size: 12px; color: #8b949e;">📅 Saída: {saida_br} | Entrega: {entrega_br}</span>
                        </div>
                    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 2. NOVA CARGA
# ----------------------------------------------------
elif menu == "➕ Nova Carga":
    st.subheader("Cadastrar Novo Agendamento de Carga")
    with st.form("form_nova_carga"):
        col1, col2 = st.columns(2)
        with col1:
            motorista = st.selectbox("Motorista Responsável", motoristas_lista if motoristas_lista else ["Nenhum cadastrado"])
            destino = st.text_input("Região / Cidades de Destino", placeholder="Ex: Uberaba, Araxá")
            observacoes = st.text_area("Observações / Rota")
        with col2:
            ajudantes = st.multiselect("Ajudantes da Viagem", ajudantes_lista)
            data_carga = st.date_input("Data do Carregamento")
            data_saida = st.date_input("Data de Saída")
            data_entrega = st.date_input("Data Prevista de Entrega")

        status_inicial = st.selectbox(
            "Status Inicial",
            [
                "Aguardando Carregamento",
                "Carregado / No Pátio",
                "Em Trânsito / Viagem Iniciada",
                "Entregue / Concluído",
            ],
        )
        if st.form_submit_button("Salvar e Agendar Carga"):
            st.success("Para registrar a nova carga, adicione a linha correspondente diretamente na sua planilha do Google Sheets!")

# ----------------------------------------------------
# 3. CADASTROS (EQUIPE)
# ----------------------------------------------------
elif menu == "👥 Cadastros (Equipe)":
    st.subheader("Gerenciamento de Motoristas e Ajudantes")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Motoristas Atuais")
        for m in motoristas_lista:
            st.text(f"👤 {m}")
    with col2:
        st.markdown("### Ajudantes Atuais")
        for a in ajudantes_lista:
            st.text(f"👥 {a}")
    st.info("Para adicionar ou remover motoristas e ajudantes, atualize as respectivas abas na sua planilha do Google Sheets.")

# ----------------------------------------------------
# 4. RELATÓRIO SEMANAL DE EXECUÇÃO E EXPORTAÇÃO
# ----------------------------------------------------
elif menu == "📊 Relatório Semanal":
    st.subheader("Relatório de Execução Semanal")

    if not cargas_lista:
        st.info("Nenhuma carga encontrada na planilha.")
    else:
        df = preparar_dataframe(cargas_lista)

        total_cargas = len(df)
        cargas_entregues = len(df[df["status"].str.contains("Entregue", case=False, na=False)]) if "status" in df.columns else 0

        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total de Cargas Registradas", total_cargas)
        col_m2.metric("Cargas Concluídas", cargas_entregues)
        col_m3.metric(
            "Taxa de Conclusão",
            f"{(cargas_entregues / total_cargas * 100):.1f}%" if total_cargas > 0 else "0%",
        )

        st.markdown("---")
        st.markdown("### Produtividade por Motorista")
        if "motorista" in df.columns:
            prod_motorista = df["motorista"].value_counts().reset_index()
            prod_motorista.columns = ["Motorista", "Total de Viagens Atribuídas"]
            st.dataframe(prod_motorista, use_container_width=True)

        st.markdown("### Detalhamento Geral de Todas as Cargas")
        st.dataframe(df, use_container_width=True)

        st.markdown("### Exportar Relatório de Cargas")
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            excel_data = gerar_excel_profissional(df)
            st.download_button(
                label="📥 Baixar Excel (.xlsx)",
                data=excel_data,
                file_name='relatorio_de_cargas.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        with col_exp2:
            pdf_bytes = gerar_pdf(df)
            st.download_button(
                label="📄 Baixar PDF",
                data=pdf_bytes,
                file_name='relatorio_de_cargas.pdf',
                mime='application/pdf',
            )
