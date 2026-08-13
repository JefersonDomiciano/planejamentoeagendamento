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
    </style>
""",
    unsafe_allow_html=True,
)

@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()

st.title("🚚 Painel de Controle de Cargas e Agendamentos")

# Bloco de diagnóstico visual para identificar eventuais falhas de conexão ou leitura
try:
    db = init_firebase()
    usar_firebase = True
    st.success("✅ Conectado ao Firebase com sucesso!")
except Exception as e:
    usar_firebase = False
    st.warning(f"⚠️ Aviso: Rodando em modo local (Firebase indisponível: {e})")

if not usar_firebase:
    if "motoristas" not in st.session_state:
        st.session_state.motoristas = ["Carlos Silva", "João Pereira", "Maurício", "Cícero Taveira"]
    if "ajudantes" not in st.session_state:
        st.session_state.ajudantes = ["Pedrinho", "Lucas Souza"]
    if "cargas" not in st.session_state:
        st.session_state.cargas = []

def carregar_dados(colecao):
    if usar_firebase:
        try:
            docs = db.collection(colecao).stream()
            return [doc.to_dict() for doc in docs]
        except Exception as e:
            st.error(f"Erro ao carregar {colecao} do Firebase: {e}")
            return []
    else:
        return st.session_state.get(colecao, [])

def salvar_dado(colecao, dados, doc_id):
    if usar_firebase:
        try:
            db.collection(colecao).document(str(doc_id)).set(dados)
        except Exception as e:
            st.error(f"Erro ao salvar no Firebase: {e}")
    else:
        if colecao == "cargas":
            encontrado = False
            for i, c in enumerate(st.session_state.cargas):
                if c.get("id") == doc_id:
                    st.session_state.cargas[i] = dados
                    encontrado = True
                    break
            if not encontrado:
                st.session_state.cargas.append(dados)

def adicionar_dado(colecao, dados, doc_id=None):
    if usar_firebase:
        try:
            if doc_id:
                db.collection(colecao).document(str(doc_id)).set(dados)
            else:
                db.collection(colecao).add(dados)
        except Exception as e:
            st.error(f"Erro ao adicionar no Firebase: {e}")
    else:
        st.session_state[colecao].append(dados)

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

try:
    motoristas_raw = carregar_dados("motoristas")
    ajudantes_raw = carregar_dados("ajudantes")
    cargas_lista = carregar_dados("cargas")
except Exception as e:
    st.error(f"Erro ao buscar coleções: {e}")
    motoristas_raw, ajudantes_raw, cargas_lista = [], [], []

motoristas_lista = [m.get("nome", m) if isinstance(m, dict) else m for m in motoristas_raw if m]
ajudantes_lista = [a.get("nome", a) if isinstance(a, dict) else a for a in ajudantes_raw if a]

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Painel (Kanban)", 
    "➕ Nova Carga", 
    "👥 Cadastros (Equipe)", 
    "📊 Relatório Semanal"
])

# ----------------------------------------------------
# 1. PAINEL (KANBAN)
# ----------------------------------------------------
with tab1:
    st.subheader("Visão Geral das Cargas Ativas")

    col_f1, col_f2 = st.columns([2, 2])
    with col_f1:
        motoristas_filtro_opcoes = ["Todos os Motoristas"] + motoristas_lista
        motorista_selecionado = st.selectbox("Filtrar por Motorista", motoristas_filtro_opcoes, key="filtro_kanban_mot")

    cargas_filtradas_periodo = [c for c in cargas_lista if c.get("status") != "Entregue / Concluído"]

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
            st.markdown(f"<div class='kanban-header'>{status}</div>", unsafe_allow_html=True)

            cargas_status_filtradas = [c for c in cargas_filtradas_periodo if c.get("status") == status]

            for carga in cargas_status_filtradas:
                carga_id = carga.get('id')
                motorista_atual = carga.get('motorista', '')
                cor_motorista = mapa_cores.get(motorista_atual, "#8b949e")

                saida_br = formatar_data_br(carga.get('data_saida'))
                entrega_br = formatar_data_br(carga.get('data_entrega'))

                with st.container():
                    c_info, c_edit, c_del = st.columns([6, 0.8, 0.8])
                    
                    with c_info:
                        st.markdown(f"""
                            <div style="border-left: 4px solid {cor_motorista}; padding-left: 8px; margin-bottom: 4px;">
                                <b style="font-size: 14px; color: #ffffff;">🚚 {motorista_atual}</b><br>
                                <span style="font-size: 13px; color: #8b949e;">Destino:</span> <span style="color: #c9d1d9; font-weight: 500;">{carga.get('destino')}</span><br>
                                <span style="font-size: 12px; color: #8b949e;">📅 Saída: {saida_br} | Entrega: {entrega_br}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    
                    with c_edit:
                        if st.button("✏️", key=f"btn_edit_{carga_id}", help="Editar"):
                            st.session_state[f"editando_{carga_id}"] = not st.session_state.get(f"editando_{carga_id}", False)
                            st.rerun()
                    
                    with c_del:
                        if st.button("🗑️", key=f"btn_del_{carga_id}", help="Excluir"):
                            if usar_firebase:
                                try:
                                    db.collection("cargas").document(str(carga_id)).delete()
                                except:
                                    pass
                            else:
                                st.session_state.cargas = [c for c in st.session_state.cargas if c["id"] != carga_id]
                            st.rerun()

                    novo_status = st.selectbox(
                        "Mover Status",
                        colunas_status,
                        index=colunas_status.index(status) if status in colunas_status else 0,
                        key=f"status_{carga_id}",
                    )
                    if novo_status != carga.get("status"):
                        carga["status"] = novo_status
                        salvar_dado("cargas", carga, carga_id)
                        st.rerun()

                    if st.session_state.get(f"editando_{carga_id}", False):
                        with st.form(key=f"form_edit_{carga_id}"):
                            st.markdown(f"**Editando Carga #{carga_id}**")
                            
                            mot_idx = motoristas_lista.index(carga.get("motorista")) if carga.get("motorista") in motoristas_lista else 0
                            novo_mot = st.selectbox("Motorista", motoristas_lista if motoristas_lista else [""], index=mot_idx)
                            novo_dest = st.text_input("Destino", value=carga.get("destino", ""))
                            
                            try:
                                dt_saida_val = datetime.date.fromisoformat(carga.get("data_saida")) if carga.get("data_saida") else datetime.date.today()
                            except:
                                dt_saida_val = datetime.date.today()
                                
                            try:
                                dt_ent_val = datetime.date.fromisoformat(carga.get("data_entrega")) if carga.get("data_entrega") else datetime.date.today()
                            except:
                                dt_ent_val = datetime.date.today()

                            nova_saida = st.date_input("Data Saída", value=dt_saida_val, key=f"saida_{carga_id}")
                            nova_entrega = st.date_input("Data Entrega", value=dt_ent_val, key=f"entrega_{carga_id}")

                            if st.form_submit_button("💾 Salvar Edição"):
                                carga["motorista"] = novo_mot
                                carga["destino"] = novo_dest
                                carga["data_saida"] = str(nova_saida)
                                carga["data_entrega"] = str(nova_entrega)
                                salvar_dado("cargas", carga, carga_id)
                                st.session_state[f"editando_{carga_id}"] = False
                                st.success("Atualizado!")
                                st.rerun()

# ----------------------------------------------------
# 2. NOVA CARGA
# ----------------------------------------------------
with tab2:
    st.subheader("Cadastrar Novo Agendamento de Carga")

    with st.form("form_nova_carga"):
        col1, col2 = st.columns(2)

        with col1:
            if motoristas_lista:
                motorista = st.selectbox("Motorista Responsável", motoristas_lista, key="nova_mot_sel")
            else:
                motorista = st.text_input("Motorista Responsável (Digite o nome)", placeholder="Ex: Carlos Silva")

            destino = st.text_input("Região / Cidades de Destino", placeholder="Ex: Uberaba, Araxá")
            observacoes = st.text_area("Observações / Rota", placeholder="Ex: Carga com entregas em lojas diferentes")

        with col2:
            ajudantes = st.multiselect("Ajudantes da Viagem", ajudantes_lista)
            data_carga = st.date_input("Data do Carregamento", key="nova_data_carg")
            data_saida = st.date_input("Data Saída", key="nova_data_said")
            data_entrega = st.date_input("Data Prevista de Entrega", key="nova_data_entr")

        status_inicial = st.selectbox(
            "Status Inicial",
            [
                "Aguardando Carregamento",
                "Carregado / No Pátio",
                "Em Trânsito / Viagem Iniciada",
                "Entregue / Concluído",
            ],
            key="nova_status_inic"
        )

        submit = st.form_submit_button("Salvar e Agendar Carga")

        if submit:
            if destino and motorista:
                if not motoristas_lista or motorista not in motoristas_lista:
                    adicionar_dado("motoristas", {"nome": motorista})

                ids_existentes = []
                for c in cargas_lista:
                    try:
                        ids_existentes.append(int(c.get("id", 0)))
                    except:
                        pass
                
                novo_id = max(ids_existentes, default=0) + 1
                
                nova_carga = {
                    "id": novo_id,
                    "motorista": motorista,
                    "destino": destino,
                    "observacoes": observacoes,
                    "ajudantes": ajudantes,
                    "data_carga": str(data_carga),
                    "data_saida": str(data_saida),
                    "data_entrega": str(data_entrega),
                    "status": status_inicial,
                }
                adicionar_dado("cargas", nova_carga, doc_id=novo_id)
                st.success(f"Carga #{novo_id} cadastrada com sucesso!")
                st.rerun()
            else:
                st.error("Preencha o destino e o motorista.")

# ----------------------------------------------------
# 3. CADASTROS (EQUIPE)
# ----------------------------------------------------
with tab3:
    st.subheader("Gerenciamento de Motoristas e Ajudantes")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Motoristas")
        novo_mot = st.text_input("Adicionar novo motorista", key="input_novo_mot")
        if st.button("Cadastrar Motorista", key="btn_cad_mot"):
            if novo_mot and novo_mot not in motoristas_lista:
                adicionar_dado("motoristas", {"nome": novo_mot})
                st.success(f"Motorista {novo_mot} adicionado!")
                st.rerun()

        st.markdown("---")
        st.write("**Motoristas Atuais:**")
        for m in motoristas_lista:
            c_mot1, c_mot2 = st.columns([4, 2])
            c_mot1.text(m)
            if c_mot2.button("Excluir", key=f"del_mot_{m}"):
                if usar_firebase:
                    try:
                        docs = db.collection("motoristas").where("nome", "==", m).stream()
                        for d in docs: d.reference.delete()
                    except:
                        pass
                else:
                    st.session_state.motoristas = [item for item in st.session_state.motoristas if item != m]
                st.rerun()

    with col2:
        st.markdown("### Ajudantes")
        novo_aju = st.text_input("Adicionar novo ajudante", key="input_novo_aju")
        if st.button("Cadastrar Ajudante", key="btn_cad_aju"):
            if novo_aju and novo_aju not in ajudantes_lista:
                adicionar_dado("ajudantes", {"nome": novo_aju})
                st.success(f"Ajudante {novo_aju} adicionado!")
                st.rerun()

        st.markdown("---")
        st.write("**Ajudantes Atuais:**")
        for a in ajudantes_lista:
            c_aju1, c_aju2 = st.columns([4, 2])
            c_aju1.text(a)
            if c_aju2.button("Excluir", key=f"del_aju_{a}"):
                if usar_firebase:
                    try:
                        docs = db.collection("ajudantes").where("nome", "==", a).stream()
                        for d in docs: d.reference.delete()
                    except:
                        pass
                else:
                    st.session_state.ajudantes = [item for item in st.session_state.ajudantes if item != a]
                st.rerun()

# ----------------------------------------------------
# 4. RELATÓRIO
# ----------------------------------------------------
with tab4:
    st.subheader("Relatório de Execução e Histórico de Cargas")

    if not cargas_lista:
        st.info("Nenhuma carga cadastrada para gerar relatório.")
    else:
        st.markdown("### 🔎 Filtros do Relatório")
        col_f1, col_f2 = st.columns(2)

        with col_f1:
            meses_opcoes = ["Todos os Meses", "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril", "05 - Maio", "06 - Junho", "07 - Julho", "08 - Agosto", "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"]
            mes_selecionado = st.selectbox("Filtrar por Mês (Data de Saída)", meses_opcoes, key="filtro_mes_rel")

        with col_f2:
            mot_rel_opcoes = ["Todos os Motoristas"] + motoristas_lista
            motorista_rel_selecionado = st.selectbox("Filtrar por Motorista", mot_rel_opcoes, key="filtro_mot_rel")

        cargas_filtradas_rel = cargas_lista

        if motorista_rel_selecionado != "Todos os Motoristas":
            cargas_filtradas_rel = [c for c in cargas_filtradas_rel if c.get("motorista") == motorista_rel_selecionado]

        if mes_selecionado != "Todos os Meses":
            mes_num = mes_selecionado.split(" - ")[0]
            cargas_temp = []
            for c in cargas_filtradas_rel:
                data_str = c.get("data_saida") or c.get("data_carga")
                if data_str:
                    try:
                        dt_obj = datetime.date.fromisoformat(data_str)
                        if f"{dt_obj.month:02d}" == mes_num:
                            cargas_temp.append(c)
                    except:
                        pass
            cargas_filtradas_rel = cargas_temp

        df = preparar_dataframe(cargas_filtradas_rel)

        total_cargas = len(df)
        cargas_entregues = len(df[df["status"] == "Entregue / Concluído"]) if "status" in df.columns else 0

        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Cargas Filtradas", total_cargas)
        col_m2.metric("Concluídas no Filtro", cargas_entregues)
        col_m3.metric(
            "Taxa de Conclusão",
            f"{(cargas_entregues / total_cargas * 100):.1f}%" if total_cargas > 0 else "0%",
        )

        st.markdown("---")
        st.markdown("### Detalhamento das Cargas Filtradas")
        if df.empty:
            st.warning("Nenhuma carga encontrada com os filtros selecionados.")
        else:
            st.dataframe(df, use_container_width=True)

            st.markdown("### Exportar Cargas Filtradas")
            col_exp1, col_exp2 = st.columns(2)

            with col_exp1:
                excel_data = gerar_excel_profissional(df)
                st.download_button(
                    label="📥 Baixar Excel (.xlsx)",
                    data=excel_data,
                    file_name='relatorio_cargas_filtrado.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

            with col_exp2:
                pdf_bytes = gerar_pdf(df)
                st.download_button(
                    label="📄 Baixar PDF",
                    data=pdf_bytes,
                    file_name='relatorio_cargas_filtrado.pdf',
                    mime='application/pdf',
                )
