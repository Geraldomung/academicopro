import streamlit as st
from utils.db import update_user_password
from utils.db import get_all_users, get_all_projects, update_user_is_active_status # Assumindo que esta função existe ou será criada

# ✅ Verificação de acesso
if st.session_state.get("role") != "Admin":
    st.error("❌ Acesso negado. Esta página é restrita a administradores.")
    st.stop()

# ✅ Conteúdo da página Admin (se passou na verificação)
st.title("👨‍💻 Painel de Administração")

st.header("📊 Estatísticas do Sistema")
# Aqui você pode adicionar mais estatísticas, se houver
st.write("Estatísticas futuras virão aqui...")

st.header("👥 Gestão de Usuários")
users = get_all_users()

# --- PARA DEPURAR: Verifique as colunas disponíveis ---
#st.write("Colunas no DataFrame de Usuários:", users.columns.tolist()) # Mostra as colunas como uma lista

# Correção: Use .empty para verificar se o DataFrame está vazio
if not users.empty:
    # Verifique se a coluna 'is_active' existe no DataFrame
    if 'is_active' not in users.columns:
        st.error("Erro: A coluna 'is_active' não foi encontrada no DataFrame de usuários. Por favor, verifique sua função get_all_users() e a estrutura do banco de dados.")
        st.stop() # Para a execução para evitar mais erros

    st.write("Atualize o status 'Ativo' dos usuários marcando ou desmarcando a caixa.")
    # Criar uma tabela interativa para cada usuário
    for index, user in users.iterrows():
        user_id = user['id']
        username = user['email']
        current_is_active = user['is_active']

        col1, col2, col3, col4, col5 = st.columns([0.4, 0.2, 0.2, 0.5, 0.5])

        with col1:
            st.write(f"**Usuário:** {username}")
        with col2:
            st.write(f"**ID:** {user_id}")
        with col3:
            # Impede que o usuário desative sua própria conta
            is_self = st.session_state.get("user_id") == user_id
            new_is_active = st.checkbox(
                "Ativo?",
                value=bool(current_is_active),
                key=f"active_status_{user_id}",
                disabled=is_self
            )
        with col4:
            st.write(f"**Último Login:** {user['last_login']}")
            # Se o status foi alterado, chame a função de atualização no banco de dados
            if bool(new_is_active) != bool(current_is_active): 
                if update_user_is_active_status(user_id, new_is_active):
                    st.success(f"Status de '{username}' atualizado para {'Ativo' if new_is_active else 'Inativo'}.")
                    st.rerun()
                else:
                    st.error(f"Falha ao atualizar o status de '{username}'.")
        with col5:
            # Campo para atualizar senha
            new_password = st.text_input(
                f"Nova senha para {username}",
                type="password",
                key=f"new_password_{user_id}"
            )
            if st.button(f"Atualizar senha de {username}", key=f"update_password_{user_id}"):
                if new_password:
                    # Importe e use a função para atualizar senha no banco de dados
                    if update_user_password(user_id, new_password):
                        st.success(f"Senha de '{username}' atualizada com sucesso.")
                    else:
                        st.error(f"Falha ao atualizar a senha de '{username}'.")
                else:
                    st.warning("Digite uma nova senha antes de atualizar.")
        st.markdown("---")
        with col4:
            st.write(f"**Último Login:** {user['last_login']}")
            # Se o status foi alterado, chame a função de atualização no banco de dados
            # Compare como booleanos para evitar problemas se current_is_active for 0/1
            if bool(new_is_active) != bool(current_is_active): 
                if update_user_is_active_status(user_id, new_is_active):
                    st.success(f"Status de '{username}' atualizado para {'Ativo' if new_is_active else 'Inativo'}.")
                    st.rerun() # Recarrega a página para refletir a mudança
                else:
                    st.error(f"Falha ao atcualizar o status de '{username}'.")
        st.markdown("---") # Separador para cada usuário
else:
    st.info("Nenhum usuário encontrado.")

with st.sidebar:
    #st.page_link("pages/1_🏠_Trabalhos.py", label="🏠 Meus Trabalhos", icon="🏠")
    #st.page_link("pages/4_📊_Estatísticas.py", label="📊 Minhas Estatísticas", icon="📊")
    
    # Exemplo: Mostrar opções de admin apenas para usuários com role 'admin'
    # Você precisa garantir que a variável 'role' seja corretamente obtida.
  

    st.divider()
    if st.button("Sair", key="logout_stats_sidebar"):
        # Lógica de logout (por exemplo, limpar st.session_state e redirecionar)
        st.session_state.clear()
        st.warning("Você foi desconectado(a).")
        st.page_link("pages/2_🔐_Login.py", label="Ir para Login")
        st.stop()

st.header("📂 Todos os Projectos")
projects = get_all_projects()
# Correção: Use .empty para verificar se o DataFrame está vazio
if not projects.empty:
    st.dataframe(projects)
else:
    st.info("Nenhum projeto encontrado.")

st.markdown("""
    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p style="font-size: 0.8em;">© 2025 AcadêmicoPro - Todos os direitos reservados</p>
    </div>
""", unsafe_allow_html=True)