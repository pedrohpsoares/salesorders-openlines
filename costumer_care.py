import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
from icons import *
import altair as alt    

# configuração da página
st.set_page_config(
    layout="wide", 
    page_title="Consulta Linhas Abertas (Atendimento)", 
    initial_sidebar_state="expanded"
)

# caminho para o arquivo transformado (saída do transform_v5.py)
DATA_PATH = "data_transformed/data_costumer_care.xlsx"
# coluna chave de cliente
COLUNA_CLIENTE_DISPLAY = 'customer_name'

# criação de kpi's em html/css
def create_kpi_card(icon, title, value):
    """Gera o HTML/CSS para um card de métrica estilizado (Dark Mode)."""
    # O estilo é injetado para criar o visual de card e remover margens padrão.
    html_content = f"""
    <div style="
        background-color: #1E1E1E; /* Fundo Escuro */
        padding: 20px; 
        border-radius: 12px; 
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); 
        text-align: left;
        margin: 0;
        height: 100%;
    ">
        <p style="
            font-size: 14px; 
            color: #AAAAAA; 
            margin: 0; 
            padding: 0;
            display: flex;
            align-items: center;
            gap: 8px;
        ">
            <span style="font-weight: 600;">{icon} {title}</span>
        </p>
        <p style="
            font-size: 28px; 
            font-weight: 700; 
            color: #FFFFFF; 
            margin-top: 5px; 
            margin-bottom: 0;
            padding: 0;
        ">
            {value}
        </p>
    </div>
    """
    return html_content

# --- funções básicas ---

@st.cache_data(ttl=600) # adicionando TTL para garantir que ele atualize a cada 10 min
def load_data(path):
    """Carrega o DataFrame final transformado."""
    try:
        import os
        full_path = os.path.abspath(path)
        print(f"DEBUG: Tentando ler em: {full_path}")
        
        # Note: Estamos lendo um arquivo .xlsx
        df = pd.read_excel(path)
        
        # CHECAGEM 1: DataFrame Vazio Imediato
        if df.empty:
            st.error(f"❌ Erro de Leitura: Arquivo encontrado em '{path}', mas está vazio (0 linhas).")
            return pd.DataFrame()
            
        # 1. Garantir que datas e valores estejam corretos
        for col in ['order_date', 'picking_date', 'Data prevista para fatura']:
            if col in df.columns:
                # Usando format='mixed' para lidar com diferentes formatos de data
                df[col] = pd.to_datetime(df[col], errors='coerce', format='mixed') 
        
        # 2. Limpeza de Colunas Críticas
        if 'cust_account_id' in df.columns:
             df['cust_account_id'] = df['cust_account_id'].astype(str).str.strip().str.upper()
        
        if 'sales_amount' in df.columns:
             df['sales_amount'] = pd.to_numeric(df['sales_amount'], errors='coerce').fillna(0)
             
        if 'salesid' in df.columns:
             df['salesid'] = df['salesid'].astype(str).str.strip()
        
        # 3. CHECAGEM DE COLUNA DE FILTRO
        if COLUNA_CLIENTE_DISPLAY not in df.columns:
             st.error(f"❌ Erro de Coluna: A coluna de filtro '{COLUNA_CLIENTE_DISPLAY}' não foi encontrada no arquivo. Colunas disponíveis: {df.columns.tolist()}")
             return pd.DataFrame()
        
        # Limpeza da chave de cliente para filtro (customer_name)
        df[COLUNA_CLIENTE_DISPLAY] = df[COLUNA_CLIENTE_DISPLAY].astype(str).str.strip()
        
        # Retorna o DataFrame
        return df
    
    except FileNotFoundError:
        st.error(f"❌ Erro de Arquivo: O arquivo {path} não foi encontrado. Execute o script transform_v5.py primeiro.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erro de Leitura Inesperado: {e}")
        return pd.DataFrame()

@st.cache_data
def convert_df_to_excel(df_to_convert, sheet_name='Extrato_Aberto'):
    """Cria o arquivo Excel na memória para download."""
    output = io.BytesIO()
    
    # CORREÇÃO: Encurtando o nome da aba para garantir que tenha menos de 31 caracteres
    # O nome "Extrato_Picking_Ativo" é muito longo. Usamos "Extrato_Pick" + data.
    short_sheet_name = 'Extrato_Pick' 
    
    # Adiciona apenas a data no nome da sheet, para ficar Extrato_Pick_YYYYMMDD (22 caracteres)
    sheet_name_with_date = f"{short_sheet_name}_{datetime.now().strftime('%Y%m%d')}"
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_to_convert.to_excel(writer, index=False, sheet_name=sheet_name_with_date)
    return output.getvalue()

# 1. CARREGAMENTO E FILTRO DE CLIENTE

df = load_data(DATA_PATH)

# antigo titulo: st.title("📊 Sales Orders - Open Lines")
st.markdown(f"## {wallet_icon} Sales Orders - Open Lines", unsafe_allow_html=True)
# Atualizado para refletir o novo filtro mais restrito
st.caption("Visão focada em linhas em aberto, dentro do intervalo de dez/24 à data presente.")

# --- NOVO BLOCO DE INFORMAÇÕES GERAIS PARA DEBUG (Total de Registros) ---
total_registros_carregados = len(df)
total_clientes_distintos = df[COLUNA_CLIENTE_DISPLAY].nunique()

# ESTE É O NOVO BLOCO DE PARADA
if df.empty: 
    st.error("Nenhum dado válido carregado. Consulte os erros de leitura acima.")
    st.stop()

# --------------------------------------------------------------------------

# Usar a conta do cliente (Cust Account) como filtro principal
lista_clientes = sorted(df[COLUNA_CLIENTE_DISPLAY].dropna().unique())

cliente_selecionado = st.sidebar.selectbox(
    "Selecione o Cliente (Nome):",
    options=['Selecione um Cliente'] + lista_clientes
)

# --- DETERMINAÇÃO E FILTRO DE DATA (NOVO BLOCO) ---
date_col = 'order_date'
df_dates = df[date_col].dropna()

# Encontra a data mínima e máxima para usar nos widgets
if not df_dates.empty:
    min_date_available = df_dates.min().date()
    max_date_available = df_dates.max().date()
else:
    # Fallback caso não haja datas válidas
    min_date_available = date.today()
    max_date_available = date.today()

st.sidebar.markdown(f"## {calendar_icon} Período (criação da ordem de venda)", unsafe_allow_html=True)

data_inicial = st.sidebar.date_input(
    "Data inicial:",
    value=min_date_available,
    min_value=min_date_available,
    max_value=max_date_available,
    key='date_start'
)

data_final = st.sidebar.date_input(
    "Data final:",
    value=max_date_available,
    min_value=min_date_available,
    max_value=max_date_available,
    key='date_end'
)

# Aplica o filtro de data no DataFrame principal (df)
if data_inicial and data_final and data_inicial <= data_final:
    df = df[
        (df[date_col].dt.date >= data_inicial) & 
        (df[date_col].dt.date <= data_final)
    ].copy()
else:
    # Se a data inicial for maior que a final, exibe um aviso e para
    if data_inicial > data_final:
        st.error("A Data Inicial não pode ser posterior à Data Final. Por favor, ajuste o período.")
        st.stop()

# ------------------------------------------
st.sidebar.markdown("---") # Linha divisória
st.sidebar.caption(f"{list_icon} Total de registros na base: **{total_registros_carregados:,}**", unsafe_allow_html=True)
st.sidebar.caption(f"{user_icon} Total de clientes ativos: **{total_clientes_distintos}**", unsafe_allow_html=True)
# ------------------------------------------

# 2. VISUALIZAÇÃO E DOWNLOAD

if cliente_selecionado and cliente_selecionado != 'Selecione um Cliente':
    
    # Filtra a base (que já contem SÓ o que está aberto e com picking ativado) pelo cliente
    df_aberto_cliente = df[df[COLUNA_CLIENTE_DISPLAY] == cliente_selecionado].copy()

    if not df_aberto_cliente.empty:
        
        # Cálculo dos KPIs a partir da coluna 'sales_amount'
        total_aberto = df_aberto_cliente['sales_amount'].sum()
        num_linhas = len(df_aberto_cliente)
        num_ordens = df_aberto_cliente['salesid'].nunique()
        
        # Extração de dados do Cliente
        vendedor = df_aberto_cliente['sales_responsible'].iloc[0] if 'sales_responsible' in df_aberto_cliente.columns else 'N/A'
        nome_cliente = df_aberto_cliente['customer_name'].iloc[0] if 'customer_name' in df_aberto_cliente.columns else 'N/A'
        conta_cliente = df_aberto_cliente['cust_account_id'].iloc[0] if 'cust_account_id' in df_aberto_cliente.columns else 'N/A'

        st.markdown(f"## {company_icon} {nome_cliente} : {conta_cliente}", unsafe_allow_html=True)
        
        # Cards de KPIs
        col1, col2, col3, col4 = st.columns(4)
        # KPI 1: Valor Total em Aberto
        with col1:
            valor_formatado = f"R$ {total_aberto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            st.markdown(create_kpi_card(
                icon=money_icon,
                title="Valor total em aberto",
                value=valor_formatado
            ), unsafe_allow_html=True)
    
        # KPI 2: Total de Linhas (Itens) em Aberto
        with col2:
            st.markdown(create_kpi_card(
                icon=list_icon, # Usando o ícone do usuário
                title="Total de linhas",
                value=f"{num_linhas:,}".replace(",", ".")
            ), unsafe_allow_html=True)
        
        # KPI 3: Total de Ordens de Venda
        with col3:
            st.markdown(create_kpi_card(
                icon=sales_icon, # Usando o ícone do usuário
                title="Total de ordens de venda",
                value=f"{num_ordens:,}".replace(",", ".")
            ), unsafe_allow_html=True)

        # KPI 4: Vendedor Responsável
        with col4:
            # Reduzindo o tamanho da fonte para o nome do vendedor, que é um texto
            vendedor_value = f"<span style='font-size: 24px; color: #FFFFFF;'>{vendedor}</span>"
            st.markdown(create_kpi_card(
                icon=user_icon, # Usando o ícone do usuário
                title="Vendedor responsável",
                value=vendedor_value), unsafe_allow_html=True)

        st.markdown("---")
        
        # Tabela Detalhada (Visão do Atendente)
        
        df_display = df_aberto_cliente.copy()
        
        # Cria um status de estoque para facilitar
        df_display['Status Estoque'] = df_display['stock_available'].apply(
            lambda x: '🟩 OK' if x > 0 else '🟥 Sem Estoque'
        )
        
        # Formatação do Valor
        df_display['Valor Aberto (R$)'] = df_display['sales_amount'].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        
        # Criando coluna de previsao de fatura em texto
        df_display['Previsão Fatura'] = df_display['Data prevista para fatura'].dt.strftime('%d/%m/%Y').fillna('N/A')

        # Seleção e Ordem final das colunas para visualização
        COLUNAS_FINAL_DISPLAY = [
            'order_date', 'salesid', 'itemid', 'open_qty_order', 
            'Valor Aberto (R$)', 'status_logistica', 'picking_route', 
            'picking_date', 'Status Estoque', 'Chegada Importação', 'Previsão Fatura', 'customer_group'
        ]
        
        df_display_final = df_display[[col for col in COLUNAS_FINAL_DISPLAY if col in df_display.columns]].copy()
        
        # Renomeação amigável (Apenas para o display)
        df_display_final.columns = [
            'Data Criação Ordem', 'Ordem Venda', 'Item ID', 'Qtd Aberta', 
            'Valor Aberto (R$)', 'Status', 'Nº Picking', 
            'Data Criação Picking', 'Status Estoque', 'Chegada Importação', 'Previsão Fatura', 'Grupo Cliente'
        ]
        
        # Exibição
        st.dataframe(
            df_display_final.sort_values(by='Data Criação Ordem', ascending=True), 
            use_container_width=True, 
            hide_index=True
        )
        
        
        # Botão de Download
        st.subheader("Gerar extrato para o atendimento")
        excel_data = convert_df_to_excel(df_aberto_cliente, sheet_name='Extrato_Picking_Ativo')
        
        st.download_button(
            label=f"💾 Baixar o extrato de linhas em aberto de {nome_cliente} (Excel)",
            data=excel_data,
            file_name=f'OpenLines_{conta_cliente}_{pd.Timestamp("today").strftime("%Y%m%d")}.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
    else:
        # Mensagem de sucesso ajustada para o novo filtro
        st.success(f"✅ Excelente! O cliente {cliente_selecionado} não possui linhas em aberto ou faturamento pendente'.")

# Caso o cliente não tenha sido selecionado
else:
    st.info("Por favor, selecione um cliente na barra lateral.")
