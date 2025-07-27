import streamlit as st
import sqlite3
from sqlite3 import Error
import hashlib
import datetime
import os
import uuid
from utils.db import DatabaseManager

db_manager = DatabaseManager()

def get_auth_status():
    """Retorna (auth_status, username, user_id, role, is_active)"""
    if st.session_state.get('logged_in'):
        return (
            True,
            st.session_state.get('username'),
            st.session_state.get('user_id'),
            st.session_state.get('role', 'user'),  # Default role
            st.session_state.get('is_active', False)
        )
    return False, None, None, None, False

def login_required(func):
    """Decorator para proteger páginas que requerem login"""
    def wrapper(*args, **kwargs):
        auth_status, _, _, _, is_active = get_auth_status()
        if not auth_status:
            st.warning("🔒 Você precisa estar logado para acessar esta página")
            st.session_state['redirect'] = st.runtime.scriptrunner.script_run_context.get().script_path
            st.switch_page("pages/2_🔐_Login.py")
            return
        if not is_active:
            st.error("❌ Sua conta está desativada. Entre em contato com o administrador.")
            st.stop()
        return func(*args, **kwargs)
    return wrapper

def role_required(required_role):
    """Decorator para verificar roles de usuário"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            _, _, _, role, is_active = get_auth_status()
            if not is_active:
                st.error("❌ Sua conta está desativada.")
                st.stop()
            if role != required_role:
                st.error("⛔ Acesso não autorizado")
                st.stop()
            return func(*args, **kwargs)
        return wrapper
    return decorator

def redirect_if_logged_in():
    """Redireciona usuários logados para a página principal"""
    auth_status, _, _, _, is_active = get_auth_status()
    if auth_status and is_active:
        redirect_to = st.session_state.pop('redirect', 'pages/1_🗂️_Projectos.py')
        st.switch_page(redirect_to)

def authenticate(username, password):
    """Verifica credenciais"""
    conn = sqlite3.connect('academicopro.db')
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, password_hash, salt, role, is_active 
            FROM users WHERE username = ?
        ''', (username,))
        result = cursor.fetchone()
        
        if result:
            user_id, stored_hash, salt_str, role, is_active = result
            if not is_active:
                st.error("❌ Conta desativada. Entre em contato com o administrador.")
                return False
            
            salt_bytes = bytes.fromhex(salt_str)
            
            input_hash = hashlib.pbkdf2_hmac(
                'sha256',
                password.encode('utf-8'),
                salt_bytes,
                100000
            ).hex()
            
            if input_hash == stored_hash:
                # Atualiza a sessão
                st.session_state['logged_in'] = True
                st.session_state['username'] = username
                st.session_state['user_id'] = user_id
                st.session_state['role'] = role
                st.session_state['is_active'] = is_active
                return True
    except Error as e:
        st.error(f"Erro de autenticação: {e}")
    finally:
        conn.close()
    return False

def register_user(username, email, password, role='user'):
    """Registra novo usuário"""
    # Verifica se o usuário já existe
    conn = sqlite3.connect('academicopro.db')
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM users WHERE username = ? OR email = ?', (username, email))
        if cursor.fetchone():
            st.error("Usuário ou email já cadastrado")
            return False

        # Cria hash seguro da senha
        salt = os.urandom(32)
        key = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt,
            100000
        ).hex()
        
        # Insere novo usuário
        salt_str = salt.hex()
        cursor.execute('''
            INSERT INTO users (username, email, password_hash, salt, role, is_active)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (username, email, key, salt_str, role, True))
        conn.commit()
        
        st.success("✅ Usuário registrado com sucesso!")
        return True
    except Error as e:
        st.error(f"Erro ao registrar: {e}")
        return False
    finally:
        conn.close()

def admin_required(func):
    """Decorator para restringir acesso a administradores"""
    def wrapper(*args, **kwargs):
        auth_status, _, _, role, is_active = get_auth_status()
        if not auth_status:
            st.error("⚠️ Você precisa estar logado para acessar esta página.")
            st.stop()
        if not is_active:
            st.error("❌ Acesso negado. Sua conta está desativada.")
            st.stop()
        if role != "Admin":
            st.error("❌ Acesso negado. Esta página é restrita a administradores.")
            st.stop()
        return func(*args, **kwargs)
    return wrapper

def logout():
    """Realiza logout do usuário"""
    keys_to_remove = ['logged_in', 'username', 'user_id', 'role', 'is_active']
    for key in keys_to_remove:
        if key in st.session_state:
            del st.session_state[key]
    st.success("👋 Você foi desconectado com sucesso!")
    redirect_to = st.session_state.pop('redirect', 'pages/2_🔐_Login.py')
    st.switch_page(redirect_to)
