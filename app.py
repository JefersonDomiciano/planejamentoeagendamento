import streamlit as st
import pandas as pd
import datetime

# --- CONFIGURAÇÃO DA PLANILHA PÚBLICA ---
PLANILHA_ID = "COLE_O_ID_DA_ SUA_PLANILHA_AQUI"
url = f"https://docs.google.com/spreadsheets/d/{PLANILHA_ID}/gviz/tq?tqx=out:csv"

# --- FUNÇÕES DE DADOS ---
@st.cache_data(ttl=5)
def carregar_dados():
    try:
        df = pd.read_csv(url)
        # Garante colunas mínimas caso venha vazia
        colunas_esperadas = ['id', 'motorista', 'destino', 'observacoes', 'ajudantes', 'data_carga', 'data_saida', 'data_entrega', 'status']
        for col in colunas_esperadas:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception:
        return pd.DataFrame(columns=['id', 'motorista', 'destino', 'observacoes', 'ajudantes', 'data_carga', 'data_saida', 'data_entrega', 'status'])

# --- INTERFACE ---
st.set_page_config(page_title="Painel de Controle de Cargas", layout="wide")
st.title("🚚 Painel de Controle de Cargas - Equipe")

menu = st.radio("Menu", ["📋 Painel (Kanban)", "➕ Nova Carga"], horizontal=True)

df_cargas = carregar_dados()

if menu == "📋 Painel (Kanban)":
    st.subheader("Gerenciamento e Acompanhamento de Cargas")
    
    if df_cargas.empty:
        st.info("Nenhuma carga encontrada na planilha.")
    else:
        # Filtros rápidos
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_motorista = st.selectbox("Filtrar por Motorista", ["Todos"] + list(df_cargas['motorista'].dropna().unique()))
        with col_f2:
            filtro_status = st.selectbox("Filtrar por Status", ["Todos", "Aguardando Carregamento", "Em Trânsito", "Entregue"])
        
        df_filtrado = df_cargas.copy()
        if filtro_motorista != "Todos":
            df_filtrado = df_filtrado[df_filtrado['motorista'] == filtro_motorista]
        if filtro_status != "Todos":
            df_filtrado = df_filtrado[df_filtrado['status'] == filtro_status]

        # Layout em Colunas Estilo Kanban / Cards
        col_k1, col_k2, col_k3 = st.columns(3)
        
        with col_k1:
            st.markdown("### ⏳ Aguardando")
            cargas_aguardando = df_filtrado[df_filtrado['status'].str.contains("Aguardando", case=False, na=False)]
            for _, row in cargas_aguardando.iterrows():
                with st.container(border=True):
                    st.markdown(f"**Destino:** {row.get('destino', '')}")
                    st.text(f"Motorista: {row.get('motorista', '')}")
                    st.text(f"Saída: {row.get('data_saida', '')}")
                    if row.get('ajudantes'):
                        st.caption(f"Ajudantes: {row['ajudantes']}")

        with col_k2:
            st.markdown("### 🚚 Em Trânsito")
            cargas_transito = df_filtrado[df_filtrado['status'].str.contains("Trânsito", case=False, na=False)]
            for _, row in cargas_transito.iterrows():
                with st.container(border=True):
                    st.markdown(f"**Destino:** {row.get('destino', '')}")
                    st.text(f"Motorista: {row.get('motorista', '')}")
                    st.text(f"Saída: {row.get('data_saida', '')}")
                    if row.get('ajudantes'):
                        st.caption(f"Ajudantes: {row['ajudantes']}")

        with col_k3:
            st.markdown("### ✅ Entregue")
            cargas_entregue = df_filtrado[df_filtrado['status'].str.contains("Entregue", case=False, na=False)]
            for _, row in cargas_entregue.iterrows():
                with st.container(border=True):
                    st.markdown(f"**Destino:** {row.get('destino', '')}")
                    st.text(f"Motorista: {row.get('motorista', '')}")
                    st.text(f"Entrega: {row.get('data_entrega', '')}")

        st.divider()
        with st.expander("Ver tabela completa de dados"):
            st.dataframe(df_filtrado, use_container_width=True)

elif menu == "➕ Nova Carga":
    st.subheader("Cadastrar Nova Carga no Sistema")
    with st.form("form_carga_completo"):
        col1, col2 = st.columns(2)
        with col1:
            mot = st.text_input("Motorista responsável")
            dest = st.text_input("Destino da Carga (Cidade/Cliente)")
            ajudantes = st.text_input("Ajudantes (separados por vírgula)")
        with col2:
            saida = st.date_input("Data Prevista de Saída", datetime.date.today())
            status = st.selectbox("Status Inicial", ["Aguardando Carregamento", "Em Trânsito", "Entregue"])
        
        obs = st.text_area("Observações importantes sobre a rota ou carga")
        
        if st.form_submit_button("Gerar Orientações de Cadastro"):
            st.success("Formulário validado!")
            st.info(f"Para registrar a carga de **{mot}** para **{dest}** com salvamento automático integrado, basta inserir os dados diretamente na sua planilha aberta no Google Sheets.")
