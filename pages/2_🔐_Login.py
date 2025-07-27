import datetime
import streamlit as st
from utils.auth import authenticate, register_user
# Importe a nova função update_last_login
from utils.db import get_user_id, get_user_role, get_user_status, update_last_login
import time

def validate_password(password):
    """Valida a complexidade da senha"""
    if len(password) < 8:
        return "A senha deve ter pelo menos 8 caracteres"
    if not any(char.isdigit() for char in password):
        return "A senha deve conter pelo menos 1 número"
    if not any(char.isupper() for char in password):
        return "A senha deve conter pelo menos 1 letra maiúscula"
    return None

def login_page():
    """Página de login com abas para login e registro"""
    st.title("🔐 Área de Login")
    
    # Verifica se já está logado (redirecionamento)
    if st.session_state.get('logged_in'):
        st.switch_page("pages/1_🗂️_Projectos.py")
    
    tab1, tab2 = st.tabs(["Login", "Registrar"])

    with tab1:
        with st.form("login_form"):
            username = st.text_input("Nome de usuário", key="login_username")
            password = st.text_input("Senha", type="password", key="login_password")
            
            if st.form_submit_button("Entrar", type="primary"):
                with st.spinner("Autenticando..."):
                    time.sleep(0.5)  # Simula processamento
                    if authenticate(username, password):
                        user_id = get_user_id(username)
                        role = get_user_role(user_id)
                        is_active = get_user_status(user_id) # Função que verifica se o usuário está ativo
                        if not is_active:
                            st.error("❌ Sua conta está desactivada. Entre em contacto com o administrador (geraldomung@outlook.com / 846182830).")
                            return
                        if user_id:
                            # Obter a data e hora atual
                            current_time = datetime.datetime.now()
                            
                            # Chamar a função para atualizar o last_login no banco de dados
                            if update_last_login(user_id, current_time):
                                st.session_state.update({
                                    'logged_in': True,
                                    'username': username,
                                    'user_id': user_id,
                                    'role': role,
                                    'is_active': is_active,
                                    'last_login': current_time # Atualiza o session_state com a mesma data e hora
                                })
                                st.success("Login bem-sucedido!")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error("Erro ao atualizar o último login. Tente novamente.")
                        else:
                            st.error("Erro ao obter informações do usuário")
                    else:
                        st.error("Credenciais inválidas")

    with tab2:
        with st.form("register_form"):
            new_username = st.text_input("Novo nome de usuário", key="reg_username")
            new_email = st.text_input("Email", key="reg_email")
            new_password = st.text_input("Nova senha", type="password", key="reg_password")
            confirm_password = st.text_input("Confirmar senha", type="password", key="reg_confirm")
            
            if st.form_submit_button("Criar conta", type="primary"):
                if new_password != confirm_password:
                    st.error("As senhas não coincidem")
                elif not new_username or not new_password:
                    st.error("Nome de usuário e senha são obrigatórios")
                else:
                    password_error = validate_password(new_password)
                    if password_error:
                        st.error(password_error)
                    else:
                        with st.spinner("Criando conta..."):
                            time.sleep(0.5)
                            if register_user(new_username, new_email, new_password):
                                user_id = get_user_id(new_username)
                                if user_id:
                                    st.session_state.update({
                                        'logged_in': True,
                                        'username': new_username,
                                        'user_id': user_id,
                                        'role': 'user',  # Default role
                                    })
                                    st.success("Conta criada e login realizado com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.success("Conta criada com sucesso! Por favor faça login.")
                            else:
                                st.error("Erro ao criar conta. Nome de usuário já existe.")

    # Rodapé
    st.markdown("""
    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p style="font-size: 0.8em;">© 2025 AcadêmicoPro - Todos os direitos reservados</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    login_page()