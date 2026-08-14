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

FIREBASE_PROJECT_ID = "logistica-d6c14"

def formatar_data_br(data_str):
    if not data_str or str(data_str).lower() in ["nan", "none", ""]:
        return ""
    try:
        dt = datetime.datetime.strptime(str(data_str).split("T")[0], "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(data_str)

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
                    if "arrayValue" in v:
                        arr_vals = v["arrayValue"].get("values", [])
                        item[k] = [list(x.values())[0] for x in arr_vals]
                    else:
                        item[k] = list(v.values())[0]
                documentos.append(item)
            return documentos
    except Exception:
        pass
    return []

def salvar_documento(colecao, doc_id, dados):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}/{doc_id}"
    fields = {}
    for k, v in dados.items():
        if isinstance(v, list):
            fields[k] = {"arrayValue": {"values": [{"stringValue": str(x)} for x in v]}}
        else:
            fields[k] = {"stringValue": str(v)}
    try:
        requests.patch(url, json={"fields": fields}, timeout=5)
    except Exception as e:
        st.error(f"Erro ao salvar no Firebase: {e}")

def deletar_documento(colecao, doc_id):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/{colecao}/{doc_id}"
    try:
        requests.delete(url, timeout=5)
    except Exception as e:
        st.error(f"Erro ao excluir no Firebase: {e}")

if "cargas" not in st.session_state:
    st.session_state["cargas"] = carregar_colecao("cargas")
if "motoristas" not in st.session_state:
    st.session_state["motoristas"] = carregar_colecao("motoristas")
if "ajudantes" not in st.session_state:
    st.session_state["ajudantes"] = carregar_colecao("ajudantes")

cargas_lista = st.session_state["cargas"]
motoristas_lista = [m.get("nome", "") for m in st.session_state["motoristas"] if m.get("nome")]
ajudantes_lista = [a.get("nome", "") for a in st.session_state["ajudantes"] if a.get("nome")]

if not motoristas_lista:
    motoristas_lista = ["Carlos Silva", "João Pereira", "Maurício", "Cícero Taveira"]
if not ajudantes_lista:
    ajudantes_lista = ["Pedrinho", "Lucas Souza"]

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

st.markdown("---")

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

    headers = ["ID Planejamento", "Motorista", "Destino", "Observações", "Ajudantes", "Data Carga", "Data Saída", "Data Entrega", "Status"]
    
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
            cell.value = "" if str(value).lower() in ["nan", "none"] else value
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
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.set_font("Arial", "B", 14)
    pdf.cell(277, 8, txt="Relatório de Cargas", ln=True, align="C")
    
    pdf.set_font("Arial", "", 9)
    pdf.cell(277, 5, txt=f"Data de geração: {datetime.date.today().strftime('%d/%m/%Y')}", ln=True, align="C")
    pdf.ln(4)

    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(47, 117, 181)
    pdf.set_text_color(255, 255, 255)
    
    larguras = [25, 40, 50, 45, 28, 28, 28, 33]
    nomes_colunas = ["ID", "Motorista", "Destino", "Ajudantes", "Carga", "Saída", "Entrega", "Status"]
    
    for i, nome in enumerate(nomes_colunas):
        pdf.cell(larguras[i], 7, nome, 1, 0, "C", True)
    pdf.ln()

    pdf.set_font("Arial", "", 8)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        pdf.cell(larguras[0], 6, str(row.get("id", "")), 1, 0, "C")
        pdf.cell(larguras[1], 6, str(row.get("motorista", ""))[:22], 1, 0, "L")
        pdf.cell(larguras[2], 6, str(row.get("destino", ""))[:30], 1, 0, "L")
        pdf.cell(larguras[3], 6, str(row.get("ajudantes", ""))[:25], 1, 0, "L")
        pdf.cell(larguras[4], 6, str(row.get("data_carga", "")), 1, 0, "C")
        pdf.cell(larguras[5], 6, str(row.get("data_saida", "")), 1, 0, "C")
        pdf.cell(larguras[6], 6, str(row.get("data_entrega", "")), 1, 0, "C")
        pdf.cell(larguras[7], 6, str(row.get("status", ""))[:18], 1, 0, "C")
        pdf.ln()

    return bytes(pdf.output(dest='S'))

# ----------------------------------------------------
# 1. PAINEL (KANBAN)
# ----------------------------------------------------
if menu == "📋 Painel (Kanban)":
    st.subheader("Visão Geral das Cargas")

    col_f1, col_f2, col_f3 = st.columns([2, 1.5, 1.5])
    with col_f1:
        motoristas_filtro_opcoes = ["Todos os Motoristas"] + motoristas_lista
        motorista_selecionado = st.selectbox("Filtrar por Motorista", motoristas_filtro_opcoes)
    
    with col_f2:
        data_inicial_filtro = st.date_input("Data Inicial (Saída)", value=datetime.date.today() - datetime.timedelta(days=7))
    with col_f3:
        data_final_filtro = st.date_input("Data Final (Saída)", value=datetime.date.today() + datetime.timedelta(days=30))

    cargas_filtradas_periodo = []
    for c in cargas_lista:
        data_str = c.get("data_saida") or c.get("data_carga")
        incluir = True
        if data_str:
            try:
                dt_obj = datetime.date.fromisoformat(str(data_str).split("T")[0])
                if not (data_inicial_filtro <= dt_obj <= data_final_filtro):
                    incluir = False
            except Exception:
                pass
        
        if incluir:
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

            cargas_status_filtradas = [c for c in cargas_filtradas_periodo if c.get("status") == status]

            for carga in cargas_status_filtradas:
                carga_id = carga.get('id')
                motorista_atual = carga.get('motorista', '')
                cor_motorista = mapa_cores.get(motorista_atual, "#8b949e")

                saida_br = formatar_data_br(carga.get('data_saida'))
                entrega_br = formatar_data_br(carga.get('data_entrega'))

                with st.container():
                    c_info, c_btn = st.columns([3.5, 2.5])
                    
                    with c_info:
                        st.markdown(f"""
                            <div style="border-left: 4px solid {cor_motorista}; padding-left: 8px; margin-bottom: 2px;">
                                <b style="font-size: 13px; color: #58a6ff;">📌 ID: {carga_id}</b><br>
                                <b style="font-size: 14px; color: #ffffff;">🚚 {motorista_atual}</b><br>
                                <span style="font-size: 13px; color: #8b949e;">Destino:</span> <span style="color: #c9d1d9; font-weight: 500;">{carga.get('destino')}</span><br>
                                <span style="font-size: 11px; color: #8b949e;">📅 Saída: {saida_br} | Entrega: {entrega_br}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    
                    with c_btn:
                        col_e, col_d = st.columns(2)
                        with col_e:
                            if st.button("✏️", key=f"btn_edit_{carga_id}", help="Editar Carga"):
                                st.session_state[f"editando_{carga_id}"] = not st.session_state.get(f"editando_{carga_id}", False)
                        with col_d:
                            if st.button("🗑️", key=f"btn_del_{carga_id}", help="Excluir Carga"):
                                deletar_documento("cargas", carga_id)
                                st.session_state["cargas"] = [c for c in cargas_lista if c.get("id") != carga_id]
                                st.rerun()

                    novo_status = st.selectbox(
                        "Mover Status",
                        colunas_status,
                        index=colunas_status.index(status) if status in colunas_status else 0,
                        key=f"status_{carga_id}",
                    )
                    
                    if novo_status != carga.get("status"):
                        carga["status"] = novo_status
                        if novo_status == "Entregue / Concluído" and not carga.get("data_entrega"):
                            carga["data_entrega"] = str(datetime.date.today())
                        
                        salvar_documento("cargas", carga_id, carga)
                        st.rerun()

                    if st.session_state.get(f"editando_{carga_id}", False):
                        with st.form(key=f"form_edit_{carga_id}"):
                            st.markdown(f"**Editando Planejamento #{carga_id}**")
                            
                            mot_idx = motoristas_lista.index(carga.get("motorista")) if carga.get("motorista") in motoristas_lista else 0
                            novo_mot = st.selectbox("Motorista", motoristas_lista if motoristas_lista else [""], index=mot_idx)
                            novo_dest = st.text_input("Destino", value=carga.get("destino", ""))
                            
                            try:
                                dt_saida_val = datetime.date.fromisoformat(str(carga.get("data_saida")).split("T")[0]) if carga.get("data_saida") else datetime.date.today()
                            except:
                                dt_saida_val = datetime.date.today()
                                
                            try:
                                dt_ent_val = datetime.date.fromisoformat(str(carga.get("data_entrega")).split("T")[0]) if carga.get("data_entrega") else datetime.date.today()
                            except:
                                dt_ent_val = datetime.date.today()

                            nova_saida = st.date_input("Data Saída", value=dt_saida_val, key=f"saida_{carga_id}")
                            nova_entrega = st.date_input("Data Entrega", value=dt_ent_val, key=f"entrega_{carga_id}")

                            col_f_salvar, _ = st.columns(2)
                            with col_f_salvar:
                                salvar_edicao = st.form_submit_button("💾 Salvar")
                            
                            if salvar_edicao:
                                carga["motorista"] = novo_mot
                                carga["destino"] = novo_dest
                                carga["data_saida"] = str(nova_saida)
                                carga["data_entrega"] = str(nova_entrega)
                                
                                salvar_documento("cargas", carga_id, carga)
                                st.session_state[f"editando_{carga_id}"] = False
                                st.success("Atualizado!")
                                st.rerun()

# ----------------------------------------------------
# 2. NOVA CARGA
# ----------------------------------------------------
elif menu == "➕ Nova Carga":
    st.subheader("Cadastrar Novo Agendamento de Carga")

    with st.form("form_nova_carga"):
        col_id_manual, _ = st.columns([2, 2])
        with col_id_manual:
            id_planejamento = st.text_input("Número do Planejamento / ID da Carga", placeholder="Ex: 1042")

        col1, col2 = st.columns(2)

        with col1:
            motorista = st.selectbox("Motorista Responsável", motoristas_lista if motoristas_lista else ["Nenhum cadastrado"])
            destino = st.text_input("Região / Cidades de Destino", placeholder="Ex: Uberaba, Araxá (Múltiplas entregas)")
            observacoes = st.text_area("Observações / Rota", placeholder="Ex: Carga com entregas em lojas diferentes")

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

        submit = st.form_submit_button("Salvar e Agendar Carga")

        if submit:
            if id_planejamento and destino and motorista:
                ids_existentes = [str(c.get("id")) for c in cargas_lista]
                if str(id_planejamento) in ids_existentes:
                    st.error(f"Já existe uma carga cadastrada com o ID/Planejamento '{id_planejamento}'. Use outro número.")
                else:
                    data_ent_val = str(data_entrega)
                    if status_inicial == "Entregue / Concluído" and not data_ent_val:
                        data_ent_val = str(datetime.date.today())

                    nova_carga = {
                        "id": str(id_planejamento),
                        "motorista": motorista,
                        "destino": destino,
                        "observacoes": observacoes,
                        "ajudantes": ajudantes,
                        "data_carga": str(data_carga),
                        "data_saida": str(data_saida),
                        "data_entrega": data_ent_val,
                        "status": status_inicial,
                    }
                    salvar_documento("cargas", str(id_planejamento), nova_carga)
                    st.session_state["cargas"].append(nova_carga)
                    st.success("Carga cadastrada com sucesso!")
                    st.rerun()
            else:
                st.error("Preencha o Número do Planejamento, o Motorista e a Região de Destino.")

# ----------------------------------------------------
# 3. CADASTROS (EQUIPE)
# ----------------------------------------------------
elif menu == "👥 Cadastros (Equipe)":
    st.subheader("Gerenciamento de Motoristas e Ajudantes")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Motoristas")
        with st.form("form_cad_mot", clear_on_submit=True):
            novo_mot = st.text_input("Adicionar novo motorista")
            cad_mot_btn = st.form_submit_button("Cadastrar Motorista")
            if cad_mot_btn and novo_mot:
                doc_id = f"mot_{int(datetime.datetime.now().timestamp())}"
                dados_mot = {"id": doc_id, "nome": novo_mot}
                salvar_documento("motoristas", doc_id, dados_mot)
                st.session_state["motoristas"].append(dados_mot)
                st.success(f"Motorista {novo_mot} adicionado!")
                st.rerun()

        st.markdown("---")
        st.write("**Motoristas Atuais:**")
        for m_obj in st.session_state["motoristas"]:
            m_nome = m_obj.get("nome", "")
            m_id = m_obj.get("id", m_nome)
            c_mot1, c_mot2 = st.columns([4, 2])
            c_mot1.text(m_nome)
            if c_mot2.button("Excluir", key=f"del_mot_{m_id}"):
                deletar_documento("motoristas", m_id)
                st.session_state["motoristas"] = [m for m in st.session_state["motoristas"] if m.get("id") != m_id]
                st.rerun()

    with col2:
        st.markdown("### Ajudantes")
        with st.form("form_cad_aju", clear_on_submit=True):
            novo_aju = st.text_input("Adicionar novo ajudante")
            cad_aju_btn = st.form_submit_button("Cadastrar Ajudante")
            if cad_aju_btn and novo_aju:
                doc_id = f"aju_{int(datetime.datetime.now().timestamp())}"
                dados_aju = {"id": doc_id, "nome": novo_aju}
                salvar_documento("ajudantes", doc_id, dados_aju)
                st.session_state["ajudantes"].append(dados_aju)
                st.success(f"Ajudante {novo_aju} adicionado!")
                st.rerun()

        st.markdown("---")
        st.write("**Ajudantes Atuais:**")
        for a_obj in st.session_state["ajudantes"]:
            a_nome = a_obj.get("nome", "")
            a_id = a_obj.get("id", a_nome)
            c_aju1, c_aju2 = st.columns([4, 2])
            c_aju1.text(a_nome)
            if c_aju2.button("Excluir", key=f"del_aju_{a_id}"):
                deletar_documento("ajudantes", a_id)
                st.session_state["ajudantes"] = [a for a in st.session_state["ajudantes"] if a.get("id") != a_id]
                st.rerun()

# ----------------------------------------------------
# 4. RELATÓRIOS (EXPORTAÇÃO E TABELAS)
# ----------------------------------------------------
elif menu == "📈 Relatórios":
    st.subheader("📈 Relatórios e Exportação de Dados")

    if not cargas_lista:
        st.info("Nenhuma carga cadastrada para gerar relatórios.")
    else:
        df_tabela = preparar_dataframe(cargas_lista)

        st.markdown("### 📋 Detalhamento Geral de Todas as Cargas")
        st.dataframe(df_tabela, use_container_width=True)

        st.markdown("### Exportar Arquivos")
        col_exp1, col_exp2 = st.columns(2)

        with col_exp1:
            excel_data = gerar_excel_profissional(df_tabela)
            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=excel_data,
                file_name='relatorio_de_cargas.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

        with col_exp2:
            pdf_bytes = gerar_pdf(df_tabela)
            st.download_button(
                label="📄 Baixar Relatório em PDF",
                data=pdf_bytes,
                file_name='relatorio_de_cargas.pdf',
                mime='application/pdf',
            )
