import streamlit as st
from g_drive_service import GoogleDriveService
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
from googleapiclient.http import MediaIoBaseDownload
import io

st.set_page_config(page_title="Google Drive Explorer", layout="wide")
st.title("🗂️ Google Drive Explorer")
st.write("Aplicativo para visualizar e baixar arquivos Excel (.xlsx) do Google Drive via API")

@st.cache_data(ttl=3600)  # Cache por 1 hora
def getFileListFromGDrive():
    """
    Obtém lista de arquivos Excel (.xlsx) do Google Drive.
    
    Returns:
        dict: Dicionário contendo a lista de arquivos
    """
    selected_fields = "files(id, name, mimeType, webViewLink, createdTime, modifiedTime, size)"
    g_drive_service = GoogleDriveService().build()
    
    try:
        # Query para filtrar apenas arquivos .xlsx
        query = "name contains '.xlsx'"
        list_file = g_drive_service.files().list(
            fields=selected_fields,
            pageSize=100,
            q=query
        ).execute()
        return {"files": list_file.get("files", [])}
    except Exception as e:
        st.error(f"Erro ao acessar o Google Drive: {e}")
        return {"files": []}

@st.cache_data(ttl=3600)  # Cache por 1 hora
def download_file(file_id, file_name):
    """
    Baixa um arquivo do Google Drive.
    
    Args:
        file_id (str): ID do arquivo no Google Drive
        file_name (str): Nome do arquivo
        
    Returns:
        BytesIO: Conteúdo do arquivo em memória
    """
    try:
        g_drive_service = GoogleDriveService().build()
        request = g_drive_service.files().get_media(fileId=file_id)
        
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        
        done = False
        with st.spinner(f"Baixando {file_name}..."):
            while not done:
                status, done = downloader.next_chunk()
                
        file.seek(0)
        return file
    except Exception as e:
        st.error(f"Erro ao baixar arquivo: {e}")
        return None

@st.cache_data(ttl=3600)  # Cache por 1 hora
def read_file_to_dataframe(file_content, file_name):
    """
    Converte o conteúdo do arquivo Excel em um DataFrame
    
    Args:
        file_content (BytesIO): Conteúdo do arquivo
        file_name (str): Nome do arquivo
        
    Returns:
        pd.DataFrame: DataFrame com os dados do arquivo
    """
    try:
        # Para arquivos Excel, podemos tentar ler múltiplas abas
        excel_file = pd.ExcelFile(file_content)
        sheet_names = excel_file.sheet_names
        
        # Se tiver múltiplas abas, retorna um dicionário com um DataFrame para cada aba
        if len(sheet_names) > 1:
            dataframes = {}
            for sheet in sheet_names:
                dataframes[sheet] = pd.read_excel(file_content, sheet_name=sheet)
            return dataframes
        else:
            return pd.read_excel(file_content)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo {file_name}: {e}")
        return None

def generate_metrics(df):
    """
    Gera métricas descritivas para um DataFrame
    
    Args:
        df (pd.DataFrame): DataFrame para gerar métricas
    """
    # Mostrar dimensões e informações básicas
    st.write(f"**Dimensões:** {df.shape[0]} linhas x {df.shape[1]} colunas")
    
    # Encontrar colunas numéricas para métricas
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns
    
    if len(numeric_cols) > 0:
        # Criar métricas para até 4 colunas numéricas
        cols = st.columns(min(4, len(numeric_cols)))
        for i, col_name in enumerate(numeric_cols[:4]):
            with cols[i % 4]:
                median = df[col_name].median()
                mean = df[col_name].mean()
                st.metric(
                    label=f"{col_name}",
                    value=f"{mean:.2f}",
                    delta=f"{mean - median:.2f} da mediana"
                )

def generate_auto_graphs(df, tab_name=""):
    """
    Gera gráficos automáticos baseados nos dados do DataFrame
    
    Args:
        df (pd.DataFrame): DataFrame para gerar visualizações
        tab_name (str): Nome da aba/planilha para contexto
    """
    if df is None or df.empty:
        st.warning("Não há dados para visualizar.")
        return
    
    st.subheader(f"📊 Visualização de Dados {tab_name}")
    
    # Gerar métricas básicas
    generate_metrics(df)
    
    # Mostrar resumo estatístico
    with st.expander("📝 Resumo estatístico"):
        st.dataframe(df.describe())
    
    # Análises por tipo de dados
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
    date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    
    # Criar tabs para diferentes tipos de visualizações
    if numeric_cols or categorical_cols or date_cols:
        graph_tabs = st.tabs(["📈 Distribuição", "🔢 Correlações", "📊 Categorias"])
        
        # Tab para distribuições
        with graph_tabs[0]:
            if numeric_cols:
                col_hist = st.selectbox(
                    "Selecione uma coluna para visualizar distribuição:",
                    numeric_cols,
                    key=f"dist_{tab_name}"
                )
                fig = px.histogram(
                    df, x=col_hist,
                    title=f"Distribuição de {col_hist}",
                    marginal="box"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Mostrar estatísticas específicas da coluna
                q1 = df[col_hist].quantile(0.25)
                q3 = df[col_hist].quantile(0.75)
                iqr = q3 - q1
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Média", f"{df[col_hist].mean():.2f}")
                col2.metric("Mediana", f"{df[col_hist].median():.2f}")
                col3.metric("Desvio Padrão", f"{df[col_hist].std():.2f}")
        
        # Tab para correlações
        with graph_tabs[1]:
            if len(numeric_cols) >= 2:
                # Matriz de correlação
                st.subheader("Matriz de Correlação")
                corr = df[numeric_cols].corr()
                fig = px.imshow(
                    corr, 
                    text_auto=True,
                    color_continuous_scale='RdBu_r',
                    title="Correlação entre variáveis numéricas"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Gráfico de dispersão
                st.subheader("Gráfico de Dispersão")
                col_x = st.selectbox(
                    "Selecione a coluna para o eixo X:",
                    numeric_cols,
                    key=f"x_{tab_name}"
                )
                col_y = st.selectbox(
                    "Selecione a coluna para o eixo Y:", 
                    [col for col in numeric_cols if col != col_x],
                    key=f"y_{tab_name}"
                )
                
                # Adicionar cor opcional se temos categorias
                if categorical_cols:
                    color_col = st.selectbox(
                        "Colorir por categoria (opcional):", 
                        [None] + categorical_cols,
                        key=f"color_{tab_name}"
                    )
                    fig = px.scatter(
                        df, x=col_x, y=col_y, 
                        color=color_col, 
                        trendline="ols",
                        title=f"{col_x} vs {col_y}"
                    )
                else:
                    fig = px.scatter(
                        df, x=col_x, y=col_y,
                        trendline="ols", 
                        title=f"{col_x} vs {col_y}"
                    )
                
                st.plotly_chart(fig, use_container_width=True)
        
        # Tab para dados categóricos
        with graph_tabs[2]:
            if categorical_cols:
                col_cat = st.selectbox(
                    "Selecione uma coluna categórica:", 
                    categorical_cols,
                    key=f"cat_{tab_name}"
                )
                
                # Contar valores únicos
                value_counts = df[col_cat].value_counts().reset_index()
                value_counts.columns = [col_cat, 'Contagem']
                
                # Limitar a 15 categorias mais frequentes
                if len(value_counts) > 15:
                    value_counts = value_counts.head(15)
                    st.info("Mostrando apenas as 15 categorias mais frequentes")
                
                # Gráfico de barras horizontal para melhor visualização
                fig = px.bar(
                    value_counts, y=col_cat, x='Contagem',
                    orientation='h',
                    title=f"Distribuição de {col_cat}",
                    text='Contagem',
                    color='Contagem'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Se houver colunas numéricas, permitir gráfico de caixa ou violin
                if numeric_cols:
                    num_col = st.selectbox(
                        "Selecione uma coluna numérica para análise por categoria:", 
                        numeric_cols,
                        key=f"numcat_{tab_name}"
                    )
                    
                    plot_type = st.radio(
                        "Tipo de gráfico:", 
                        ["Boxplot", "Violin Plot"],
                        horizontal=True,
                        key=f"plottype_{tab_name}"
                    )
                    
                    if plot_type == "Boxplot":
                        fig = px.box(
                            df, x=col_cat, y=num_col,
                            title=f"{num_col} por {col_cat}"
                        )
                    else:
                        fig = px.violin(
                            df, x=col_cat, y=num_col,
                            title=f"{num_col} por {col_cat}",
                            box=True
                        )
                    
                    st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Não foram encontradas colunas apropriadas para gerar gráficos automaticamente.")

# Barra lateral com opções
with st.sidebar:
    st.header("Opções")
    refresh_button = st.button("🔄 Atualizar Lista")
    
    st.markdown("---")
    st.markdown("Desenvolvido com ❤️ e Streamlit")

# Corpo principal
try:
    files_data = getFileListFromGDrive()
    
    if not files_data["files"]:
        st.info("Nenhum arquivo Excel (.xlsx) encontrado no Google Drive.")
    else:
        # Converter para DataFrame para facilitar a exibição
        files_df = pd.DataFrame(files_data["files"])
        
        # Formatar as datas
        if 'createdTime' in files_df.columns:
            files_df['createdTime'] = pd.to_datetime(files_df['createdTime']).dt.strftime('%d/%m/%Y %H:%M')
        if 'modifiedTime' in files_df.columns:
            files_df['modifiedTime'] = pd.to_datetime(files_df['modifiedTime']).dt.strftime('%d/%m/%Y %H:%M')
        
        # Adicionar formatação para o tamanho
        if 'size' in files_df.columns:
            files_df['size'] = files_df['size'].astype(float) / 1024
            files_df['size'] = files_df['size'].round(2).astype(str) + ' KB'
        
        # Adicionar coluna de ações
        if 'webViewLink' not in files_df.columns:
            files_df['webViewLink'] = None
        
        # Criar tabs para navegação
        tab1, tab2 = st.tabs(["📁 Arquivos Excel", "📊 Visualização de Dados"])
        
        with tab1:
            # Exibir arquivos em uma tabela interativa
            st.subheader(f"Arquivos Excel ({len(files_df)})")
            
            # Expande o dataframe para mostrar as colunas mais importantes
            cols_to_show = ['name', 'mimeType', 'modifiedTime']
            cols_to_show = [col for col in cols_to_show if col in files_df.columns]
            
            st.dataframe(files_df[cols_to_show], height=400)
            
            # Seleção de arquivos para visualizar ou baixar
            st.subheader("🔍 Visualizar ou Baixar Arquivo")
            
            selected_file = st.selectbox(
                "Selecione um arquivo Excel",
                options=files_df['name'].tolist(),
                index=None
            )
            
            if selected_file:
                file_info = files_df[files_df['name'] == selected_file].iloc[0]
                file_id = file_info['id']
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if file_info['webViewLink']:
                        st.markdown(f"[🔗 Abrir no Google Drive]({file_info['webViewLink']})")
                
                with col2:
                    download_btn = st.download_button(
                        label="📥 Baixar Arquivo",
                        data=download_file(file_id, selected_file),
                        file_name=selected_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                if st.button("📊 Visualizar Dados"):
                    # Armazenar o arquivo selecionado na sessão
                    file_content = download_file(file_id, selected_file)
                    if file_content:
                        st.session_state.selected_file = {
                            'name': selected_file,
                            'content': file_content,
                            'id': file_id
                        }
                        # Mudar para a segunda tab automaticamente
                        st.query_params(tab='visualizacao')
                        st.experimental_rerun()
        
        with tab2:
            if st.query_params.get('tab') == ['visualizacao'] and hasattr(st.session_state, 'selected_file'):
                selected_file = st.session_state.selected_file
                st.subheader(f"📄 {selected_file['name']}")
                
                # Tentar ler o arquivo
                df_data = read_file_to_dataframe(selected_file['content'], selected_file['name'])
                
                if df_data is not None:
                    # Verificar se o resultado é um dicionário (múltiplas abas) ou um DataFrame único
                    if isinstance(df_data, dict):
                        # Criar tabs dinâmicas para cada aba da planilha
                        sheet_tabs = st.tabs(list(df_data.keys()))
                        
                        for i, sheet_name in enumerate(df_data.keys()):
                            with sheet_tabs[i]:
                                st.subheader(f"Aba: {sheet_name}")
                                df = df_data[sheet_name]
                                st.dataframe(df, height=300)
                                generate_auto_graphs(df, tab_name=f"({sheet_name})")
                    else:
                        # Mostrar um único DataFrame
                        st.dataframe(df_data, height=300)
                        generate_auto_graphs(df_data)
                else:
                    st.warning("Não foi possível ler o arquivo Excel.")
            else:
                # Carregar todos os arquivos Excel automaticamente
                st.info("Selecione um arquivo Excel na aba 'Arquivos Excel' para visualizar seus dados, ou aguarde enquanto carregamos todos os arquivos disponíveis.")
                
                # Mostrar visualização de todos os arquivos Excel encontrados
                if len(files_df) > 0:
                    # Criar tabs dinâmicas para cada arquivo Excel
                    excel_file_tabs = st.tabs([name for name in files_df['name']])
                    
                    for i, (idx, file_info) in enumerate(files_df.iterrows()):
                        with excel_file_tabs[i]:
                            st.subheader(f"📄 {file_info['name']}")
                            
                            # Verificar se já temos os dados em cache (usando o ID do arquivo como chave)
                            file_cache_key = f"file_data_{file_info['id']}"
                            
                            if file_cache_key not in st.session_state:
                                with st.spinner("Carregando dados..."):
                                    file_content = download_file(file_info['id'], file_info['name'])
                                    if file_content:
                                        df_data = read_file_to_dataframe(file_content, file_info['name'])
                                        st.session_state[file_cache_key] = df_data
                                    else:
                                        st.session_state[file_cache_key] = None
                            
                            df_data = st.session_state[file_cache_key]
                            
                            if df_data is not None:
                                # Similar ao código anterior, mas para cada arquivo
                                if isinstance(df_data, dict):
                                    sheet_tabs = st.tabs(list(df_data.keys()))
                                    
                                    for j, sheet_name in enumerate(df_data.keys()):
                                        with sheet_tabs[j]:
                                            df = df_data[sheet_name]
                                            st.dataframe(df, height=300)
                                            generate_auto_graphs(df, tab_name=f"({sheet_name})")
                                else:
                                    st.dataframe(df_data, height=300)
                                    generate_auto_graphs(df_data)
                            else:
                                st.warning("Não foi possível ler este arquivo Excel.")

except Exception as e:
    st.error(f"Ocorreu um erro: {e}")
    st.exception(e)  # Adiciona detalhes do erro para facilitar diagnósticos