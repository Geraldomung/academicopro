import streamlit as st
import hashlib
import os
import sqlite3
from sqlite3 import Error
import json
#from turtle import st
import uuid
from datetime import datetime
import pandas as pd
from typing import Dict, List, Optional, Tuple, Union

# Configuração do banco de dados
DATABASE_NAME = "academicopro.db"

class DatabaseManager:
    """Classe para gerenciar todas as operações do banco de dados"""
    
    def __init__(self):
        self.conn = None
        self.init_db()  # Alterado de _initialize_db para init_db
    
    def init_db(self):  # Renomeado de _initialize_db para init_db
        """Inicializa o banco de dados criando as tabelas necessárias"""
        try:
            self.conn = self._get_connection()
            cursor = self.conn.cursor()
            
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
                    last_login TIMESTAMP,
                    is_active BOOLEAN DEFAULT 1
                )
            ''')
            
            # Tabela de projetos
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS projects (
                    id TEXT PRIMARY KEY,
                    user_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    data TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
                    UNIQUE(user_id, name)
                )
            ''')
            
            # Índices para melhor performance
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_projects_user ON projects(user_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_projects_updated ON projects(updated_at)')
            
            # Cria usuário admin padrão se não existir
            cursor.execute("SELECT id FROM users WHERE username = 'admin'")
            if cursor.fetchone() is None:
                self._create_admin_user()
                
            self.conn.commit()
            
        except Error as e:
            print(f"Erro ao inicializar banco de dados: {e}")
            raise
        finally:
            if self.conn:
                self.conn.close()
    
    def _create_admin_user(self):
        """Cria o usuário admin padrão"""
        salt = os.urandom(32)
        password = "Admin@123"
        password_hash = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt,
            100000
        ).hex()
        
        self.conn.execute('''
            INSERT INTO users (username, password_hash, salt, role)
            VALUES (?, ?, ?, ?)
        ''', ('admin', password_hash, salt.hex(), 'Admin'))
    
    def _get_connection(self):
        """Obtém uma conexão com o banco de dados"""
        conn = sqlite3.connect(DATABASE_NAME)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")  # Ativa chaves estrangeiras
        return conn
    
    def _execute_query(self, query: str, params: tuple = (), commit: bool = False) -> Optional[sqlite3.Cursor]:
        """Executa uma query genérica com tratamento de erros"""
        try:
            self.conn = self._get_connection()
            cursor = self.conn.cursor()
            cursor.execute(query, params)
            if commit:
                self.conn.commit()
            return cursor
        except Error as e:
            print(f"Erro na query: {query} | Params: {params} | Erro: {e}")
            return None
        except Exception as e:
            print(f"Erro inesperado: {e}")
            return None
        # Note: A conexão é fechada pelo chamador quando necessário
    def update_user_password(self, user_id: int, new_password: str) -> bool:
        """Atualiza a senha de um usuário"""
        try:
            salt = os.urandom(32)
            password_hash = hashlib.pbkdf2_hmac(
                'sha256',
                new_password.encode('utf-8'),
                salt,
                100000
            ).hex()
            cursor = self._execute_query(
                "UPDATE users SET password_hash = ?, salt = ? WHERE id = ?",
                (password_hash, salt.hex(), user_id),
                commit=True
            )
            if cursor and cursor.connection:
                cursor.connection.close()
            return bool(cursor and cursor.rowcount > 0)
        except Exception as e:
            print(f"Erro ao atualizar senha do usuário: {e}")
            return False
    # Operações de usuário
    def get_user_by_username(self, username: str) -> Optional[Dict]:
        """Obtém um usuário pelo username"""
        cursor = self._execute_query(
            "SELECT * FROM users WHERE username = ?",
            (username,)
        )
        if cursor:
            result = cursor.fetchone()  # Chamada única
            if cursor.connection:
                cursor.connection.close()
            return dict(result) if result else None
        return None
    
    def register_user(self, username: str, email: str, password: str) -> bool:
        """Registra um novo usuário com senha criptografada"""
        if self.get_user_by_username(username):
            return False
            
        salt = os.urandom(32)
        password_hash = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt,
            100000
        ).hex()
        
        query = '''
            INSERT INTO users (username, email, password_hash, salt)
            VALUES (?, ?, ?, ?)
        '''
        cursor = self._execute_query(query, (username, email, password_hash, salt.hex()), commit=True)
        if cursor and cursor.connection:
            cursor.connection.close()
        return bool(cursor)
    
    def authenticate_user(self, username: str, password: str) -> Optional[Dict]:
        user = self.get_user_by_username(username)
        if not user:
            return None

        try:
            salt = user['salt'] if isinstance(user['salt'], bytes) else bytes.fromhex(user['salt'])
        except Exception as e:
            print(f"Erro ao converter salt: {e}")
            return None

        password_hash = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt,
            100000
        ).hex()

        if password_hash == user['password_hash']:
            self._execute_query(
                "UPDATE users SET last_login = CURRENT_TIMESTAMP WHERE id = ?",
                (user['id'],),
                commit=True
            )
            return user
        return None
    
    def get_user_role(self, user_id: int) -> str:
        """Obtém o role/perfil do usuário"""
        cursor = self._execute_query(
            "SELECT role FROM users WHERE id = ?",
            (user_id,)
        )
        if cursor:
            result = cursor.fetchone()
            if cursor.connection:
                cursor.connection.close()
            return result['role'] if result else 'user'
        return 'user'

    def get_user_status(self, user_id: int) -> bool:
        """Retorna True se o usuário estiver ativo"""
        cursor = self._execute_query(
            "SELECT is_active FROM users WHERE id = ?",
            (user_id,)
        )
        if cursor:
            result = cursor.fetchone()
            if cursor.connection:
                cursor.connection.close()
            return result["is_active"] == 1 if result else False
        return False
    

    def update_user_is_active_status(self, user_id, is_active):
        try:
            self.conn = self._get_connection()
            cursor = self.conn.cursor()
            # Converte booleano Python para 0 ou 1 para o banco de dados se necessário
            active_value = 1 if is_active else 0 
            cursor.execute("UPDATE users SET is_active = ? WHERE id = ?", (active_value, user_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Erro ao atualizar status do usuário: {e}")
            return False
        finally:
            if cursor and cursor.connection:
                cursor.connection.close()
            return bool(cursor)


    def update_last_login(self, user_id, login_time):
    
        
        try:
            self.conn = self._get_connection()
            cursor = self.conn.cursor()
            # Formate o datetime para string se o seu banco de dados exigir (SQLite geralmente aceita objetos datetime diretamente)
            # Para SQLite, 'YYYY-MM-DD HH:MM:SS' é o formato recomendado para TIMESTAMP
            login_time_str = login_time.strftime("%Y-%m-%d %H:%M:%S") 
            cursor.execute("UPDATE users SET last_login = ? WHERE id = ?", (login_time_str, user_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Erro ao atualizar o último login para o usuário {user_id}: {e}")
            return False
        finally:
            if cursor and cursor.connection:
                cursor.connection.close()
    # Operações de projetos
   
        
    def update_project(self, project_id: str, project_data: Dict) -> bool:
        """Atualiza um projeto existente"""
        query = '''
            UPDATE projects 
            SET data = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        '''
        cursor = self._execute_query(
            query,
            (json.dumps(project_data), project_id),
            commit=True
        )
        if cursor and cursor.connection:
            cursor.connection.close()
        return bool(cursor)
    
    def save_project(self, user_id: str, project_data: dict) -> tuple[bool, int | None]:
        """
        Salva um novo projeto usando o tema do project_data como nome.
        Retorna (success, project_id) ou (False, None) em caso de falha.
        """
        try:
            # Validação e preparação do tema
            tema = project_data.get('tema', '').strip()
            if not tema:
                tema = f"Projeto_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                project_data['tema'] = tema  # Atualiza os dados

            # Conversão e salvamento
            data_json = json.dumps(project_data, ensure_ascii=False, indent=2)
            
            cursor = self._execute_query(
                """INSERT INTO projects (user_id, name, data) 
                VALUES (?, ?, ?)""",
                (user_id, tema, data_json),
                commit=True
            )
            
            return (True, cursor.lastrowid) if cursor.lastrowid else (False, None)

        except Exception as e:
            st.error(f"Database Error: {str(e)}")
            return False, None




    
    
    def update_project(self, project_id: str, project_data: dict) -> bool:
        """Atualiza um projeto existente"""
        try:
            # Converter dados para JSON
            data_json = json.dumps(project_data, ensure_ascii=False)
            
            cursor = self._execute_query(
                "UPDATE projects SET data = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (data_json, project_id),
                commit=True
            )
            
            return bool(cursor and cursor.rowcount > 0)
        except Exception as e:
            st.error(f"Erro ao atualizar projeto: {str(e)}")
            return False
    
    
    def get_user_projects(self, user_id: str) -> list[dict]:
        """Retorna todos os projetos do usuário no formato padrão"""
        try:
            cursor = self._execute_query(
                "SELECT id, name, data, created_at, updated_at FROM projects WHERE user_id = ?",
                (user_id,)
            )
            
            projects = []
            for row in cursor.fetchall():
                projects.append({
                    'id': row[0],          # ID gerado pelo auto incremento
                    'name': row[1],       # Nome do projeto
                    'data': json.loads(row[2]),  # Dados do projeto (convertidos de JSON)
                    'created_at': row[3],  # Timestamp de criação
                    'updated_at': row[4]   # Timestamp de atualização
                })
            
            return projects
        
        except Exception as e:
            st.error(f"Erro ao carregar projetos: {str(e)}")
            return []    
    
    def load_project(self, project_id: str) -> dict:
        """Carrega um projeto específico"""
        try:
            cursor = self._execute_query(
                "SELECT id, name, data, created_at, updated_at FROM projects WHERE id = ?",
                (project_id,)
            )
            
            if cursor:
                row = cursor.fetchone()
                if row:
                    return {
                        'id': row[0],
                        'name': row[1],
                        'data': json.loads(row[2]),
                        'created_at': row[3],
                        'updated_at': row[4]
                    }
            return None
        except Exception as e:
         st.error(f"Erro ao carregar projeto: {str(e)}")
        return None
    
    def delete_project(self, project_id: str) -> bool:
        """Remove um projeto do banco de dados"""
        cursor = self._execute_query(
            "DELETE FROM projects WHERE id = ?",
            (project_id,),
            commit=True
        )
        return bool(cursor and cursor.rowcount > 0)
    
    def get_user_projects(self, user_id: int) -> List[Dict]:
        """Lista todos projetos de um usuário"""
        cursor = self._execute_query(
            '''
            SELECT id, name, created_at, updated_at 
            FROM projects 
            WHERE user_id = ?
            ORDER BY updated_at DESC
            ''',
            (user_id,)
        )
        if cursor:
            projects = [dict(row) for row in cursor.fetchall()]
            if cursor.connection:
                cursor.connection.close()
            return projects
        return []
    
    def get_project_metadata(self, project_id: str) -> Optional[Dict]:
        """Obtém metadados de um projeto"""
        cursor = self._execute_query(
            "SELECT id, name, created_at, updated_at FROM projects WHERE id = ?",
            (project_id,)
        )
        if cursor:
            result = cursor.fetchone()
            if cursor.connection:
                cursor.connection.close()
            return dict(result) if result else None
        return None
    
    # Operações administrativas
    def get_all_users(self) -> pd.DataFrame:
        """Retorna todos os usuários cadastrados"""
        try:
            conn = self._get_connection()
            query = '''
                SELECT id, username, email, role, created_at, last_login, is_active 
                FROM users 
                ORDER BY created_at DESC
            '''
            df = pd.read_sql(query, conn)
            return df
        except Error as e:
            print(f"Erro ao buscar usuários: {e}")
            return pd.DataFrame()
        finally:
            if conn:
                conn.close()
    
    def get_all_projects(self) -> pd.DataFrame:
        """Retorna todos os projetos do sistema"""
        try:
            conn = self._get_connection()
            query = '''
                SELECT p.id, p.name,u.email, u.username, p.created_at, p.updated_at 
                FROM projects p
                JOIN users u ON p.user_id = u.id
                ORDER BY p.updated_at DESC
            '''
            df = pd.read_sql(query, conn)
            return df
        except Error as e:
            print(f"Erro ao buscar projetos: {e}")
            return pd.DataFrame()
        finally:
            if conn:
                conn.close()
    
    def get_user_stats(self, user_id: int) -> Dict[str, Union[int, str]]:
        """Obtém estatísticas do usuário"""
        stats = {
            'total_projects': 0,
            'active_projects': 0,
            'last_project_date': "N/A"
        }
        
        # Total de projetos
        cursor = self._execute_query(
            "SELECT COUNT(*) FROM projects WHERE user_id = ?",
            (user_id,)
        )
        if cursor:
            stats['total_projects'] = cursor.fetchone()[0]
            if cursor.connection:
                cursor.connection.close()
        
        # Projetos ativos (últimos 30 dias)
        cursor = self._execute_query(
            '''
            SELECT COUNT(*) FROM projects 
            WHERE user_id = ? 
            AND updated_at >= datetime('now', '-30 days')
            ''',
            (user_id,)
        )
        if cursor:
            stats['active_projects'] = cursor.fetchone()[0]
            if cursor.connection:
                cursor.connection.close()
        
        # Data do último projeto
        cursor = self._execute_query(
            '''
            SELECT MAX(created_at) FROM projects 
            WHERE user_id = ?
            ''',
            (user_id,)
        )
        if cursor:
            result = cursor.fetchone()
            if result and result[0]:
                stats['last_project_date'] = result[0][:10]
            if cursor.connection:
                cursor.connection.close()
        
        return stats

# Inicializa o banco de dados quando o módulo é carregado

def delete_project(project_id: str) -> bool:
    """Remove um projeto do banco de dados"""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM projects WHERE id = ?", (project_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"Erro ao deletar projeto: {e}")
        return False
    finally:
        if conn:
            conn.close()

db_manager = DatabaseManager()            
# Funções de interface para compatibilidade com código existente
def init_db():
    """Função de inicialização para compatibilidade"""
    pass  # Já é feito automaticamente pelo DatabaseManager

def get_db_connection():
    """Obtém conexão com o banco (para compatibilidade)"""
    return db_manager._get_connection()

def register_user(username, email, password):
    return db_manager.register_user(username, email, password)

def save_project(user_id, project_data, project_name=None):
    return db_manager.save_project(user_id, project_data, project_name)

def load_project(project_id):
    return db_manager.load_project(project_id)

def get_user_projects(user_id):
    return db_manager.get_user_projects(user_id)

def get_user_stats(user_id):
    return db_manager.get_user_stats(user_id)

def get_user_id(username):
    user = db_manager.get_user_by_username(username)
    return user['id'] if user else None

def get_user_role(user_id):
    return db_manager.get_user_role(user_id)

def get_all_users():
    return db_manager.get_all_users()

def get_all_projects():
    return db_manager.get_all_projects()
def get_user_status(user_id):
    return db_manager.get_user_status(user_id)
def update_user_is_active_status(user_id, is_active):
    return db_manager.update_user_is_active_status(user_id, is_active)

def update_last_login(user_id, login_time):
    return db_manager.update_last_login(user_id, login_time)
def update_user_password(user_id, new_password):
    return db_manager.update_user_password(user_id, new_password)
# Adicione no final do db.py
__all__ = [
    'DatabaseManager',
    'save_project',
    'load_project',
    'load_complete_project',
    'delete_project',
    'get_user_projects',
    'get_user_stats',
    'init_db',
    'get_db_connection',
    'register_user',
    'get_user_id',
    'get_user_role',
    'get_all_users',
    'get_all_projects',
    'get_user_status',
    'update_last_login',
    'update_user_password'
]