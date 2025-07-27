# init_db.py
import sqlite3
from sqlite3 import Error
import streamlit as st
import hashlib
import os

def create_connection():
    """Cria uma conexão com o banco de dados SQLite"""
    conn = None
    try:
        conn = sqlite3.connect('academicopro.db')
        return conn
    except Error as e:
        st.error(f"Erro ao conectar ao banco de dados: {e}")
    return conn

def create_tables(conn):
    """Cria as tabelas necessárias no banco de dados"""
    try:
        cursor = conn.cursor()
        
        # Tabela de usuários
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            role TEXT DEFAULT 'user',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_login TIMESTAMP
        )
        ''')
        
        # Tabela de projetos
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
        ''')
        
        # Tabela de sessões (opcional para controle de sessões)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
        ''')
        
        conn.commit()
        st.success("Tabelas criadas com sucesso!")
        
        # Cria um usuário admin padrão se não existir
        create_default_admin(conn)
        
    except Error as e:
        st.error(f"Erro ao criar tabelas: {e}")

def create_default_admin(conn):
    """Cria um usuário admin padrão se não existir"""
    try:
        cursor = conn.cursor()
        
        # Verifica se já existe um admin
        cursor.execute("SELECT id FROM users WHERE username = 'admin'")
        if cursor.fetchone() is None:
            # Cria um salt aleatório
            salt = os.urandom(32)
            password = "Admin@123"  # Senha padrão - deve ser alterada após o primeiro login
            password_hash = hashlib.pbkdf2_hmac(
                'sha256',
                password.encode('utf-8'),
                salt,
                100000
            ).hex()
            
            cursor.execute('''
            INSERT INTO users (username, password_hash, salt, role)
            VALUES (?, ?, ?, ?)
            ''', ('admin', password_hash, salt.hex(), 'admin'))
            
            conn.commit()
            st.warning("Usuário admin padrão criado. Por favor, altere a senha após o primeiro login.")
    
    except Error as e:
        st.error(f"Erro ao criar usuário admin: {e}")

def main():
    st.title("🔧 Inicialização do Banco de Dados")
    
    if st.button("Criar/Atualizar Banco de Dados"):
        conn = create_connection()
        if conn is not None:
            create_tables(conn)
            conn.close()

if __name__ == "__main__":
    main()