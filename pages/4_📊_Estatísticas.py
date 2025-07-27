import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Importações das suas utilidades de autenticação e banco de dados
from utils.auth import get_auth_status
from utils.db import get_user_projects, get_user_stats # Assumindo que get_user_stats e get_user_projects retornam dados

# --- Configuração da Página ---
st.set_page_config(page_title="Estatísticas", page_icon="📊", layout="wide")

# --- Verificação de Autenticação ---
# Assume que get_auth_status() também pode retornar 'role' para consistência
# Se não retornar, ajuste para buscar 'role' separadamente ou remova a checagem de admin se não for aplicável aqui.

auth_status, username, user_id, role, is_active = get_auth_status()
if not auth_status:
    st.warning("Por favor, faça login para acessar esta página.")
    st.info("Redirecionando para a página de Login...")
    # Substitua "pages/0_🔑_Login.py" pelo caminho real da sua página de login
    st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Login")
    st.stop() # Interrompe a execução do script se não estiver autenticado

# --- Título da Página ---
st.title(f"📊 Estatísticas de {username}")

# --- Sidebar de Navegação ---
with st.sidebar:
    #st.page_link("pages/1_🏠_Trabalhos.py", label="🏠 Meus Trabalhos", icon="🏠")
    #st.page_link("pages/4_📊_Estatísticas.py", label="📊 Minhas Estatísticas", icon="📊")
    
    # Exemplo: Mostrar opções de admin apenas para usuários com role 'admin'
    # Você precisa garantir que a variável 'role' seja corretamente obtida.
    if 'role' in locals() and role == 'admin': # Verifica se 'role' existe e é 'admin'
        st.divider()
        st.page_link("pages/3_👨‍💻_Admin.py", label="👨Painel Admin", icon="👨‍💻")

    st.divider()
    if st.button("Sair", key="logout_stats_sidebar"):
        # Lógica de logout (por exemplo, limpar st.session_state e redirecionar)
        st.session_state.clear()
        st.warning("Você foi desconectado(a).")
        st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Login")
        st.stop()


# --- Seção 1: Estatísticas Gerais ---
st.header("📈 Visão Geral")
try:
    stats = get_user_stats(user_id)
    if stats:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Projetos", stats.get('total_projects', 0))
        col2.metric("Projetos Ativos", stats.get('active_projects', 0))
        
        last_date_str = stats.get('last_project_date')
        if last_date_str:
            try:
                # Tenta converter para datetime e formatar
                last_project_dt = datetime.fromisoformat(last_date_str)
                col3.metric("Último Projeto", last_project_dt.strftime("%d/%m/%Y"))
            except ValueError:
                col3.metric("Último Projeto", "Data Inválida")
        else:
            col3.metric("Último Projeto", "N/A")
    else:
        st.info("Nenhuma estatística geral disponível para seus projetos.")
except Exception as e:
    st.error(f"Erro ao carregar estatísticas gerais: {e}")
    st.info("Verifique a implementação de `get_user_stats`.")


# --- Seção 2: Evolução Temporal ---
st.header("🕒 Evolução Temporal")
try:
    projects_data = get_user_projects(user_id) # Renomeado para evitar conflito com 'projects' do pd.DataFrame
    
    if projects_data:
        # Converter para DataFrame, assumindo que projects_data é uma lista de dicionários
        # Exemplo de keys esperadas: 'id', 'name', 'created_at', 'updated_at'
        df = pd.DataFrame(projects_data)
        
        # Renomear colunas para consistência, se necessário
        # Ex: Se suas colunas forem 'id', 'nome', 'criacao', 'atualizacao'
        df = df.rename(columns={
            'id': 'ID',
            'name': 'Nome',
            'created_at': 'Criação',
            'updated_at': 'Atualização'
        })

        # Converter para datetime
        df['Criação'] = pd.to_datetime(df['Criação'], errors='coerce')
        df['Atualização'] = pd.to_datetime(df['Atualização'], errors='coerce')
        
        # Remover linhas onde a data de criação é NaT (Not a Time) após a conversão
        df.dropna(subset=['Criação'], inplace=True)

        if not df.empty:
            # Agrupar por mês e ano para garantir ordem correta
            df['Mês_Ano'] = df['Criação'].dt.strftime('%Y-%m') # Formato 'AAAA-MM' para ordenação
            monthly_stats = df.groupby('Mês_Ano').size().reset_index(name='Projetos')
            
            # Gráfico de linhas
            fig = px.line(
                monthly_stats,
                x='Mês_Ano',
                y='Projetos',
                title='Projetos Criados por Mês',
                markers=True,
                labels={'Mês_Ano': 'Mês/Ano', 'Projetos': 'Número de Projetos'},
                line_shape="spline" # Opcional: para suavizar a linha
            )
            fig.update_xaxes(dtick="M1", tickformat="%b-%Y") # Formato de tick para meses
            st.plotly_chart(fig, use_container_width=True)
            
            # Mostrar projetos recentes
            st.subheader("Projetos Recentes")
            # Ordena por 'Atualização' se existir, caso contrário por 'Criação'
            sort_column = 'Atualização' if 'Atualização' in df.columns else 'Criação'
            
            # Garante que as colunas existem antes de configurar
            column_cfg = {}
            if 'ID' in df.columns: column_cfg["ID"] = None
            if 'Criação' in df.columns: column_cfg["Criação"] = st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm")
            if 'Atualização' in df.columns: column_cfg["Atualização"] = st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm")
            
            st.dataframe(
                df.sort_values(sort_column, ascending=False).head(5).reset_index(drop=True), # reset_index para evitar index antigo
                hide_index=True,
                column_config=column_cfg
            )
        else:
            st.info("Dados de projeto insuficientes para gerar a evolução temporal.")
    else:
        st.info("Você ainda não possui projetos salvos para análise temporal.")
except Exception as e:
    st.error(f"Erro ao gerar a evolução temporal dos projetos: {e}")
    st.info("Verifique a estrutura dos dados retornados por `get_user_projects` e a conversão de datas.")


# --- Seção 3: Análise de Conteúdo (se aplicável) ---
st.header("🔍 Análise de Conteúdo")
try:
    if projects_data: # Usar projects_data do get_user_projects
        # Verifica se pelo menos um projeto tem a chave 'data' e dentro dela a chave 'tipo'
        # Isso evita IndexError em projects_data[0] se a lista estiver vazia
        # e KeyError se 'data' ou 'tipo' não existirem em todos os projetos.
        
        # Filtra projetos que possuem a chave 'data' e a chave 'tipo' dentro de 'data'
        projects_with_type = [p for p in projects_data if 'data' in p and 'tipo' in p['data']]

        if projects_with_type:
            tipos = [p['data'].get('tipo', 'Não especificado') for p in projects_with_type]
            tipo_counts = pd.Series(tipos).value_counts().reset_index()
            tipo_counts.columns = ['Tipo', 'Quantidade']
            
            fig_pie = px.pie(
                tipo_counts,
                names='Tipo',
                values='Quantidade',
                title='Distribuição por Tipo de Projeto',
                color_discrete_sequence=px.colors.qualitative.Pastel # Cores mais agradáveis
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Dados de 'tipo' de projeto não disponíveis para análise de conteúdo. Verifique se seus projetos salvos incluem a chave 'data' com um 'tipo'.")
    else:
        st.info("Nenhum projeto salvo para análise de conteúdo.")
except Exception as e:
    st.error(f"Não foi possível analisar os conteúdos dos projetos: {e}")
    st.info("Verifique a estrutura dos dados na chave 'data' dos seus projetos salvos.")



       # Rodapé
st.markdown("""
    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p style="font-size: 0.8em;">© 2025 AcadêmicoPro - Todos os direitos reservados</p>
    </div>
    """, unsafe_allow_html=True)