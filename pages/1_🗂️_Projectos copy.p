import streamlit as st
from datetime import datetime
import json
from utils.auth import get_auth_status
from utils.db import DatabaseManager
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import sys
import re
import os
from dotenv import load_dotenv
import string
import google.generativeai as genai
import serpapi
import requests
from bs4 import BeautifulSoup
#from serpapi import GoogleSearch
import serpapi
print(serpapi.__file__)  # Verifica o caminho do pacote

# --- Configuração inicial ---
load_dotenv()
# Configuração do SerpAPI
SERPAPI_KEY = os.getenv("SERPAPI_KEY") or st.secrets.get("SERPAPI_KEY")
if not SERPAPI_KEY:
    st.warning("⚠️ Chave SerpAPI não encontrada. Buscas no Google Scholar serão desativadas.")
sys.path.append('./utils')


GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or st.secrets.get("GOOGLE_API_KEY")

if not GOOGLE_API_KEY:
    st.error("ERRO: Chave GOOGLE_API_KEY não encontrada.")
    st.stop()

try:
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    st.error(f"Erro ao inicializar o modelo Gemini: {str(e)}")
    st.stop()

# --- Autenticação ---
#auth_status, username, user_id, role = get_auth_status()
auth_status, username, user_id, role, is_active = get_auth_status()
if not auth_status:
    st.warning("Por favor, faça login para acessar esta página.")
    st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Login")
    st.stop()

with st.sidebar:
    if st.button("Sair", key="logout_stats_sidebar"):
        # Lógica de logout (por exemplo, limpar st.session_state e redirecionar)
        st.session_state.clear()
        st.warning("Você foi desconectado(a).")
        st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Login")
        st.stop()


# --- Inicialização do banco de dados ---
#db_manager = DatabaseManager()
# --- Inicialização do banco de dados (usando st.cache_resource para inicializar uma vez) ---

@st.cache_data(ttl=86400)  # Cache por 24h para economizar requisições
def buscar_referencias_serpapi(tema, max_resultados=5):
    if not SERPAPI_KEY:
        st.error("Chave SerpAPI não configurada.")
        return []

    try:
        params = {
            "engine": "google_scholar",
            "q": tema,
            "api_key": SERPAPI_KEY,
            "num": max_resultados,
            "hl": "pt"  # Idioma português
        }

        resultados = serpapi.search(params).get("organic_results", [])
        
        referencias_formatadas = []
        for item in resultados:
            titulo = item.get("title", "Sem título")
            autores = item.get("publication_info", {}).get("authors", "Autores desconhecidos")
            ano = item.get("publication_info", {}).get("year", "Sem ano")
            link = item.get("link", "#")
            
            referencias_formatadas.append({
                "titulo": titulo,
                "autores": autores,
                "ano": ano,
                "link": link
            })
        
        return referencias_formatadas

    except Exception as e:
        st.error(f"Erro na busca: {str(e)}")
        return []

@st.cache_resource
def get_database_manager():
    return DatabaseManager()

db_manager = get_database_manager()

# Carregar projetos do usuário ao iniciar a sessão (se o usuário estiver autenticado)
if 'user_projects' not in st.session_state and auth_status:
    st.session_state['user_projects'] = db_manager.get_user_projects(user_id)

def initialize_session_state():
    """Inicializa todos os estados da sessão necessários"""
    defaults = {
        'passo': 1,
        'tema': "",
        'curso': "",
        'problema_pesquisa': "",
        'problema_pesquisa_sugestoes': [],
        'problematizacao': "",
        'tipo_inquerito_pesquisa': "Perguntas de Pesquisa",
        'perguntas_pesquisa': [],
        'hipoteses': [],
        'objetivo_geral': "",
        'objetivos_especificos': [],
        'justificativa_pessoal': "",
        'justificativa_academica': "",
        'justificativa_social': "",
        'resultados_esperados': [],
        'introducao': "",
        'conclusao': "",
        'metodologia_sugestoes': {
            "natureza": "",
            "natureza_justificativa": "",
            "abordagem": "",
            "abordagem_justificativa": "",
            "objetivos_pesquisa": "",
            "fundamentacao_teorica": "",
            "visao_pesquisador": "",
            "procedimentos_tecnicos": [],
            "universo_amostra": "",
            "populacao": "",
            "amostra": "",
            "tipo_amostra": "",
            "instrumentos_coleta": [],
            "analise_de_dados": [],
            "consideracoes_eticas": ""
        },
        'referencial_teorico_sugestoes': [],
        'autores_citados': [],
        'norma_bibliografica': "APA",
        'referencias_bibliograficas': "",
        'cronograma_results_sugestoes': {
            'cronograma': {},
            'tabela_md': "",
            'resultados_esperados': []
        },
        'meses_cronograma': [],
        'instrumento_gerado': "",
        'tipo_instrumento_display': "",
        'nivel_academico': "Licenciatura",
        'local_foco': "",
        'current_project_id': None,
        'user_projects': [],
        'tipo_instrumento_selecionado':{},
        'todos_instrumentos_gerados':{},
        'instrumento_counter': 0
        
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value



def buscar_com_serpapi(tema, api_key):
    params = {
        "engine": "google_scholar",
        "q": tema,
        "api_key": api_key,
        "num": 5
    }
    
    try:
        resultados = serpapi.search(params)
        return resultados.get("organic_results", [])
    
    except Exception as e:
        st.error(f"Erro na API: {str(e)}")
        return []
# Inicializa o estado da sessão se ainda não estiver definido
def get_project_data():
    """Consolida todos os dados do projeto atual em um dicionário padronizado"""
    return {
        'tema': st.session_state.tema,
        'curso': st.session_state.curso,
        'local_foco': st.session_state.local_foco,
        'problema_pesquisa': st.session_state.problema_pesquisa,
        'problematizacao': st.session_state.problematizacao,
        'tipo_inquerito_pesquisa': st.session_state.tipo_inquerito_pesquisa,
        'perguntas_pesquisa': st.session_state.perguntas_pesquisa,
        'hipoteses': st.session_state.hipoteses,
        'objetivo_geral': st.session_state.objetivo_geral,
        'objetivos_especificos': st.session_state.objetivos_especificos,
        'justificativa_pessoal': st.session_state.justificativa_pessoal,
        'justificativa_academica': st.session_state.justificativa_academica,
        'justificativa_social': st.session_state.justificativa_social,
        'referencial_teorico_sugestoes': st.session_state.referencial_teorico_sugestoes,
        'autores_citados': st.session_state.autores_citados,
        'norma_bibliografica': st.session_state.norma_bibliografica,
        'referencias_bibliograficas': st.session_state.referencias_bibliograficas,
        'metodologia_sugestoes': st.session_state.metodologia_sugestoes,
        'cronograma_results_sugestoes': st.session_state.cronograma_results_sugestoes,
        'meses_cronograma': st.session_state.meses_cronograma,
        'introducao': st.session_state.introducao,
        'conclusao': st.session_state.conclusao,
        'instrumento_gerado': st.session_state.instrumento_gerado,
        'tipo_instrumento_display': st.session_state.tipo_instrumento_display,
        'nivel_academico': st.session_state.nivel_academico,
        'passo': st.session_state.passo,
        'palavras_chave': st.session_state.palavras_chave,
        'area_interesse': st.session_state.area_interesse,
        'tipo_instrumento_selecionado': st.session_state.tipo_instrumento_selecionado,
        'todos_instrumentos_gerados': st.session_state.todos_instrumentos_gerados,
        'instrumento_counter': st.session_state.instrumento_counter
    }

def load_project_data(project_data):
    """Carrega todos os dados do projeto na sessão"""

    # ⚠️ Primeiro definimos o tipo de inquérito (antes de qualquer widget ser renderizado)
    tipo = project_data.get('tipo_inquerito_pesquisa', 'Perguntas de Pesquisa')
    if 'tipo_inquerito_pesquisa' not in st.session_state:
        st.session_state['tipo_inquerito_pesquisa'] = tipo

    # ⚠️ Agora já podemos definir o conteúdo correspondente
    if tipo == "Hipóteses":
        st.session_state['conteudo_inquerito'] = project_data.get('hipoteses', [])
    else:
        st.session_state['conteudo_inquerito'] = project_data.get('perguntas_pesquisa', [])
       # ⚠️ Garante que campos essenciais existam mesmo que venham ausentes

       # Define os campos principais um por um
    campos_simples = [
        'tema',
        'curso',
        'area_interesse',
        'palavras_chave',
        'nivel_academico',
        'local_foco'
    ]

    for campo in campos_simples:
        valor = project_data.get(campo, "")
        if valor and st.session_state.get(campo) != valor:
            st.session_state[campo] = valor

    # Carrega os demais campos genéricos se existirem
    for key, value in project_data.items():
        if key not in st.session_state or st.session_state.get(key) != value:
            try:
                st.session_state[key] = value
            except st.StreamlitAPIException:
                pass  # ignora campos que já estão sendo usados por widgets
 
def save_current_project() -> None:
    """Gerencia o salvamento/atualização do projeto atual"""
    project_data = get_project_data()
    
    # Validação crítica
    if not project_data.get('tema'):
        st.error("⚠️ Tema não definido")
        st.info("Defina um tema antes de salvar")
        return

    try:
        # Fluxo de Atualização
        if st.session_state.get('current_project_id'):
            success = db_manager.update_project(
                project_id=st.session_state.current_project_id,
                project_data=project_data
            )
            action = "atualizado"
        
        # Fluxo de Criação
        else:
            success, project_id = db_manager.save_project(
                user_id=st.session_state.user_id,
                project_data=project_data
            )
            
            if success:
                st.session_state.current_project_id = project_id
                action = "salvo"
            else:
                raise RuntimeError("Falha ao gerar ID no banco")

        # Pós-operacao
        if success:
            st.session_state.user_projects = db_manager.get_user_projects(st.session_state.user_id)
            st.toast(
                f"✅ Projeto '{project_data['tema']}' {action} (ID: {st.session_state.current_project_id})",
                icon="✅"
            )
        else:
            raise RuntimeError("Operação não afetou registros")

    except Exception as e:
        st.error(f"🚨 Falha crítica: {str(e)}")
        st.code(f"Debug Data:\n{json.dumps(project_data, indent=2)}", language='json')


# --- Funções auxiliares ---
# Função auxiliar para detectar problemas em seções
import streamlit as st
import re

def avaliar_secao(texto, secao_nome):
    problemas = []
    if not texto or len(texto.strip()) < 100:
        problemas.append("🔴 Conteúdo muito curto ou ausente.")
    if any(p in texto.lower() for p in [
        "este trabalho tem como objetivo", "a presente pesquisa visa",
        "de acordo com a literatura", "é importante ressaltar que"
    ]):
        problemas.append("🟡 Frases genéricas comuns em textos gerados por IA.")
    palavras = texto.lower().split()
    if len(set(palavras)) < len(palavras) * 0.5:
        problemas.append("🟡 Repetição excessiva de palavras.")
    if texto.count('.') < len(palavras) / 30:
        problemas.append("🔶 Poucos pontos finais: possível falta de estrutura frasal.")
    return problemas

def avaliar_qualidade_documento(projeto):
    secoes = {
        "Introdução": projeto.get('introducao', ''),
        "Problema de Pesquisa": projeto.get('problema_pesquisa', ''),
        "Problematização": projeto.get('problematizacao', ''),
        "Hipóteses": '\n'.join(projeto.get('hipoteses', [])),
        "Perguntas de Pesquisa": '\n'.join(projeto.get('perguntas_pesquisa', [])),
        "Objetivo Geral": projeto.get('objetivo_geral', ''),
        "Objetivos Específicos": '\n'.join(projeto.get('objetivos_especificos', [])),
        "Justificativa Pessoal": projeto.get('justificativa_pessoal', ''),
        "Justificativa Acadêmica": projeto.get('justificativa_academica', ''),
        "Justificativa Social": projeto.get('justificativa_social', ''),
        "Referencial Teórico": '\n'.join(projeto.get('referencial_teorico_sugestoes', [])),
        "Metodologia": '\n'.join(str(v) for v in projeto.get('metodologia_sugestoes', {}).values() if v),
        "Resultados Esperados": '\n'.join(projeto.get('cronograma_results_sugestoes', {}).get('resultados_esperados', [])),
        "Conclusão": projeto.get('conclusao', ''),
        "Referências Bibliográficas": projeto.get('referencias_bibliograficas', '')
    }

    for nome_secao, conteudo in secoes.items():
        problemas = avaliar_secao(conteudo, nome_secao)
        if problemas:
            with st.expander(f"⚠️ Problemas detectados em {nome_secao}"):
                for p in problemas:
                    st.markdown(f"- {p}")
                if st.button(f"🔁 Regenerar {nome_secao}", key=f"regen_{nome_secao}"):
                    st.session_state['secao_a_regenerar'] = nome_secao

def regenerar_secao(secao_nome, projeto):
    if secao_nome == "Introdução":
        return generate_introducao_gemini(
            projeto['tema'], projeto['problema_pesquisa'], projeto['problematizacao'],
            projeto['objetivo_geral'], projeto['objetivos_especificos'],
            projeto['metodologia_sugestoes'].get("fundamentacao_teorica", "")
        )
    elif secao_nome == "Conclusão":
        return generate_conclusao_gemini(
            projeto['tema'], projeto['problema_pesquisa'], projeto['objetivo_geral'],
            projeto['objetivos_especificos'], projeto['cronograma_results_sugestoes'].get("resultados_esperados", []),
            contexto={}
        )
    elif secao_nome == "Problema de Pesquisa":
        return generate_problema_suggestions_gemini(
            projeto['tema'], projeto['curso'], projeto['local_foco'], projeto['objetivo_geral']
        )[0]
    elif secao_nome == "Problematização":
        return generate_problematization_gemini(projeto['problema_pesquisa'])
    elif secao_nome == "Hipóteses" or secao_nome == "Perguntas de Pesquisa":
        return generate_inquiry_gemini(
            projeto['problema_pesquisa'], projeto['objetivo_geral'], projeto['objetivos_especificos'], projeto['tipo_inquerito_pesquisa']
        )
    elif secao_nome == "Objetivo Geral" or secao_nome == "Objetivos Específicos":
        return generate_objetivos_suggestions_gemini(
            projeto['tema'], projeto['problema_pesquisa'], projeto['curso'], projeto['local_foco']
        )
    elif secao_nome.startswith("Justificativa"):
        return generate_justificativa_suggestions_gemini(
            projeto['tema'], projeto['problema_pesquisa'], projeto['objetivo_geral'],
            projeto['objetivos_especificos'], projeto['curso'], projeto['local_foco']
        )
    elif secao_nome == "Referencial Teórico":
        return generate_referencial_teorico_suggestions_gemini(
            projeto['tema'], projeto['problema_pesquisa'], projeto['curso'], projeto['local_foco'],
            projeto['objetivo_geral'], projeto['objetivos_especificos']
        )
    elif secao_nome == "Metodologia":
        return generate_metodologia_suggestions_gemini(
            projeto['objetivo_geral'], projeto['objetivos_especificos'], projeto['tema'],
            projeto['problema_pesquisa'], projeto['curso'], projeto['local_foco']
        )
    elif secao_nome == "Resultados Esperados":
        return generate_cronograma_results_gemini(
            projeto['tema'], projeto['objetivos_especificos'], projeto['curso']
        )
    elif secao_nome == "Referências Bibliográficas":
        full_text = "\n".join([
            projeto.get('introducao', ''), projeto.get('problematizacao', ''), projeto.get('referencial_teorico_sugestoes', ''),
            projeto.get('metodologia_sugestoes', ''), projeto.get('conclusao', '')
        ])
        return generate_references_gemini(
            full_text, projeto.get('norma_bibliografica', 'APA'), projeto['tema'], projeto['curso']
        )
    return "[Função de regeneração não definida para esta seção]"

# Exemplo de uso:
# if st.button("🧪 Avaliar Qualidade do Documento"):
#     avaliar_qualidade_documento(get_project_data())
# 
# if 'secao_a_regenerar' in st.session_state:
#     nova = regenerar_secao(st.session_state['secao_a_regenerar'], get_project_data())
#     st.success(f"{st.session_state['secao_a_regenerar']} regenerada com sucesso!")


def extract_block(text, start_marker, end_marker=None, include_marker=False):
    """Extrai um bloco de texto entre marcadores"""
    start_index = text.find(start_marker)
    if start_index == -1:
        return ""

    if not include_marker:
        start_index += len(start_marker)

    if end_marker:
        end_index = text.find(end_marker, start_index)
        if end_index == -1:
            return text[start_index:].strip()
        return text[start_index:end_index].strip()
    return text[start_index:].strip()
def create_instrument_word_document(title, instrument_text):
    """Cria um documento Word formatado com estrutura semântica"""
    try:
        document = Document()
        
        # Configuração de estilos
        style = document.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        paragraph_format = style.paragraph_format
        paragraph_format.line_spacing = 1.5
        paragraph_format.space_after = Pt(6)

        # Título principal
        title_heading = document.add_heading(title, level=1)
        title_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_page_break()

        # Processamento inteligente do conteúdo
        current_table = []
        in_table = False
        list_level = 0

        for line in instrument_text.split('\n'):
            line = line.strip()
            
            # Ignora linhas vazias
            if not line:
                if in_table:
                    current_table.append([""])
                continue
                
            # Detecção de tabelas
            if "|" in line and "-|-" not in line:
                if not in_table:
                    in_table = True
                row = [cell.strip() for cell in line.split("|") if cell.strip()]
                current_table.append(row)
                continue
            
            # Finalização de tabela
            if in_table and "|" not in line:
                in_table = False
                if current_table:
                    table = document.add_table(rows=1, cols=len(current_table[0]))
                    table.style = 'Table Grid'
                    
                    # Cabeçalho
                    hdr_cells = table.rows[0].cells
                    for i, header in enumerate(current_table[0]):
                        hdr_cells[i].text = header
                    
                    # Linhas de dados
                    for row in current_table[1:]:
                        row_cells = table.add_row().cells
                        for i, cell in enumerate(row):
                            row_cells[i].text = cell
                
                current_table = []
            
            # Listas numeradas
            if re.match(r'^\d+\.', line):
                list_level = 1
                p = document.add_paragraph(style='List Number')
                p.add_run(line.split('.', 1)[1].strip())
                continue
                
            # Listas com marcadores
            if line.startswith('- ') or line.startswith('* '):
                list_level = 1
                p = document.add_paragraph(style='List Bullet')
                p.add_run(line[2:].strip())
                continue
                
            # Subtítulos
            if line.startswith('## '):
                document.add_heading(line[3:].strip(), level=2)
                continue
                
            if line.startswith('### '):
                document.add_heading(line[4:].strip(), level=3)
                continue
                
            # Texto formatado
            p = document.add_paragraph()
            
            # Negrito
            if '**' in line:
                parts = line.split('**')
                for i, part in enumerate(parts):
                    run = p.add_run(part.strip())
                    if i % 2 == 1:
                        run.bold = True
            # Itálico
            elif '*' in line:
                parts = line.split('*')
                for i, part in enumerate(parts):
                    run = p.add_run(part.strip())
                    if i % 2 == 1:
                        run.italic = True
            else:
                p.add_run(line)

        # Salva em buffer
        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"Erro ao formatar documento: {str(e)}")
        return None

def extract_metodologia_block(text, start_marker, end_marker=None, as_list=False):
    """Extrai um bloco de metodologia"""
    block = extract_block(text, start_marker, end_marker)
    if as_list:
        items = [line.strip().replace('- ', '') for line in block.split('\n') if line.strip().startswith('- ')]
        return items if items else [block.strip()] if block.strip() else []
    return block

def extract_authors_from_referencial(referencial_list):
    """Extrai autores do referencial teórico"""
    authors = set()
    for item in referencial_list:
        match = re.search(r':\s*(.*)', item)
        if match:
            authors_str = match.group(1)
            current_authors = [a.strip() for a in authors_str.split(',') if a.strip()]
            for author in current_authors:
                author = re.sub(r'\s*(et al\.|e outros|teoria de|de)\s*', '', author, flags=re.IGNORECASE)
                parts = author.split()
                if len(parts) > 0:
                    extracted_name = parts[-1] if len(parts) > 1 else parts[0]
                    authors.add(extracted_name)
    
    filtered_authors = {
        name for name in authors 
        if not re.search(r'(Teoria|Modelo|Princípio|Conceito)\s*(da|do|de|das|dos)?', name, re.IGNORECASE) and 
        len(name) > 1
    }
    return sorted(list(filtered_authors))

def extract_citations_from_full_text(full_text):
    """Extrai citações do texto completo"""
    citations = set()
    patterns = [
        r'\(([A-Z][a-z]+(?: [A-Z][a-z]+)*), \d{4}\)',
        r'([A-Z][a-z]+(?: [A-Z][a-z]+)*) \((\d{4})\)',
        r'segundo ([A-Z][a-z]+(?: [A-Z][a-z]+)*)',
        r'conforme ([A-Z][a-z]+(?: [A-Z][a-z]+)*)',
        r'([A-Z][a-z]+(?: [A-Z][a-z]+)*) e ([A-Z][a-z]+(?: [A-Z][a-z]+)*)',
        r'([A-Z][a-z]+(?: [A-Z][a-z]+)*) et al\.',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, full_text)
        for match in matches:
            for group in match.groups():
                if group and not group.isdigit():
                    if ' e ' in group:
                        for author in group.split(' e '):
                            citations.add(author.strip())
                    else:
                        citations.add(group.strip())
    
    return sorted(list(citations))

# --- Funções de geração com IA ---
@st.cache_data(show_spinner="Gerando sugestões de tema com IA...")
def generate_tema_suggestions_gemini(area_interesse, curso, palavras_chave, nivel_academico):
    prompt = f"""
    Com base na área de interesse '{area_interesse}', para um curso de {curso},
    e considerando as palavras-chave '{palavras_chave}', em nível {nivel_academico},
    sugira 10 (dez) ideias de tema para um pré-projeto de pesquisa.
    As sugestões devem ser claras, concisas e relevantes para a área e o nível.
    Apresente as sugestões em formato de lista numerada. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
    Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
    Nunca use conteudos detectados de IA.
    Humanize todas a costrução frásica. 
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip().split('\n')
    except Exception as e:
        st.error(f"Falha na geração de temas: {str(e)}")
        return ["Sugestão 1: [Falha na Geração]", "Sugestão 2: [Falha na Geração]", "Sugestão 3: [Falha na Geração]", 
                "Sugestão 4: [Falha na Geração]", "Sugestão 5: [Falha na Geração]", "Sugestão 6: [Falha na Geração]",
                "Sugestão 7: [Falha na Geração]", "Sugestão 8: [Falha na Geração]", "Sugestão 9: [Falha na Geração]", 
                "Sugestão 10: [Falha na Geração]"]

@st.cache_data(show_spinner="Gerando problemas de pesquisa com IA...")
def generate_problema_suggestions_gemini(tema, curso, local_foco, objetivo_geral):
    prompt = f"""
    Para o tema de pesquisa '{tema}' no curso de {curso}, com foco em {local_foco},
    e tendo como Objetivo Geral: '{objetivo_geral}',
    sugira 10 (dez) problemas de pesquisa concisos e relevantes em formato de pergunta, outras em forma de declarações.
    Apresente as sugestões em formato de lista numerada. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
 
    Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
    Nunca use conteudos detectados de IA.
    Humanize todas a costrução frásica. 
    """
    try:
        response = model.generate_content(prompt)
        return [re.sub(r'^\d+\.\s*', '', line).strip() for line in response.text.strip().split('\n') if line.strip()]
    except Exception as e:
        st.error(f"Falha na geração de problemas de pesquisa: {str(e)}")
        return ["[Falha na Geração de Problema 1]", "[Falha na Geração de Problema 2]", "[Falha na Geração de Problema 3]",
                "[Falha na Geração de Problema 4]", "[Falha na Geração de Problema 5]", "[Falha na Geração de Problema 6]",
                "[Falha na Geração de Problema 7]", "[Falha na Geração de Problema 8]", "[Falha na Geração de Problema 9]",
                "[Falha na Geração de Problema 10]"]

@st.cache_data(show_spinner="Problematizando a pesquisa com IA...")
def generate_problematization_gemini(problema_pesquisa):
    prompt = f"""
    Com base no seguinte problema de pesquisa:
    "{problema_pesquisa}"

    Desenvolva uma problematização detalhada, abordando os seguintes pontos:
    - Contexto do problema: Onde ele se insere? Qual a situação atual?
    - Relevância: Por que este problema é importante? Quais são as lacunas de conhecimento existentes?
    - Implicações: Quais são as possíveis consequências ou impactos se o problema não for abordado?
    - Quem é afetado e como?
    - Justificativa da escolha do problema: Por que este problema específico merece ser investigado?

    Apresente a problematização de forma dissertativa, com parágrafos bem estruturados e coesos. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
    Mantenha a coerência com o restante do projeto
    Humanize o texto, incluindo elementos pessoais quando necessário
    Mantenha a clareza e a concisão
    Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
    Nunca use conteudos detectados de IA.
    Humanize todas a costrução frásica. 
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Falha ao gerar a problematização: {str(e)}")
        return "Não foi possível gerar a problematização. Por favor, tente novamente ou insira manualmente."

@st.cache_data(show_spinner="Gerando perguntas de pesquisa/hipóteses com IA...")
def generate_inquiry_gemini(problema_pesquisa, objetivo_geral, objetivos_especificos, tipo_inquerito):
    objetivos_str = "\n- " + "\n- ".join(objetivos_especificos) if objetivos_especificos else "Não fornecidos."
    
    if tipo_inquerito == "Perguntas de Pesquisa":
        prompt = f"""
        Com base no Problema de Pesquisa: "{problema_pesquisa}",
        no Objetivo Geral: "{objetivo_geral}",
        e nos Objetivos Específicos:
        {objetivos_str}

        Gere 3 a 5 perguntas de pesquisa claras, concisas e diretamente alinhadas aos objetivos específicos e ao problema.
        As perguntas devem ser formuladas de modo que, ao serem respondidas, contribuam para resolver o problema de pesquisa e alcançar os objetivos.
        Apresente-as em formato de lista numerada. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
        Mantenha a coerência com o restante do projeto
        Humanize o texto, incluindo elementos pessoais quando necessário
        Mantenha a clareza e a concisão
        Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
        Nunca use conteudos detectados de IA.
        Humanize todas a costrução frásica. 
        """
    else: # "Hipóteses"
        prompt = f"""
        Com base no Problema de Pesquisa: "{problema_pesquisa}",
        no Objetivo Geral: "{objetivo_geral}",
        e nos Objetivos Específicos:
        {objetivos_str}

        Gere 3 a 5 hipóteses testáveis e diretamente alinhadas aos objetivos específicos.
        As hipóteses devem ser declarações sobre a relação esperada entre variáveis, formuladas de forma a serem provadas ou refutadas pela pesquisa.
        Apresente-as em formato de lista numerada. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
        Mantenha a coerência com o restante do projeto
        Humanize o texto, incluindo elementos pessoais quando necessário
        Mantenha a clareza e a concisão

        Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
        Nunca use conteudos detectados de IA.
        Humanize todas a costrução frásica. 
        """
    try:
        response = model.generate_content(prompt)
        return [re.sub(r'^\d+\.\s*', '', line).strip() for line in response.text.strip().split('\n') if line.strip()]
    except Exception as e:
        st.error(f"Falha ao gerar {tipo_inquerito.lower()}: {str(e)}")
        return [f"[Falha na Geração de {tipo_inquerito} 1]", f"[Falha na Geração de {tipo_inquerito} 2]", f"[Falha na Geração de {tipo_inquerito} 3]"]

@st.cache_data(show_spinner="Gerando objetivos com IA...")
def generate_objetivos_suggestions_gemini(tema, problema, curso, local_foco):
    prompt = f"""
    Considerando o tema '{tema}', o problema de pesquisa '{problema}', para o curso de {curso} com foco em {local_foco},
    sugira um **Objetivo Geral** claro e direto e 3 a 5 **Objetivos Específicos** que desdobrem o objetivo geral. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
    Formate a resposta EXATAMENTE assim:

    OBJETIVO GERAL: [Seu objetivo geral aqui]

    OBJETIVOS ESPECÍFICOS:
    - [Objetivo Específico 1]
    - [Objetivo Específico 2]
    - [Objetivo Específico 3]
    - [Objetivo Específico 4 (Opcional)]
    - [Objetivo Específico 5 (Opcional)]

    Use verbos no infinitivo
    Humanize o texto, sempre com tons humanos
    Cada verbo deve corrresponder a um objectivo, comecar a frase dos objectivos com um verbo no infinitivo
    EX:. Identificar, Caracterizar, Avaliar, etc.
    """
    try:
        response = model.generate_content(prompt)
        output = response.text.strip()
        
        objetivo_geral_match = re.search(r"OBJETIVO GERAL: (.*)", output)
        objetivos_especificos_block = re.search(r"OBJETIVOS ESPECÍFICOS:\n(.*)", output, re.DOTALL)
        
        objetivo_geral = objetivo_geral_match.group(1).strip() if objetivo_geral_match else ""
        objetivos_especificos = []
        if objetivos_especificos_block:
            especificos_text = objetivos_especificos_block.group(1).strip()
            objetivos_especificos = [line.strip().replace('- ', '') for line in especificos_text.split('\n') if line.strip().startswith('- ')]
            
        return objetivo_geral, objetivos_especificos
    except Exception as e:
        st.error(f"Falha na geração de objetivos: {str(e)}")
        return "[Falha na Geração do Objetivo Geral]", ["[Falha na Geração do Objetivo Específico 1]", "[Falha na Geração do Objetivo Específico 2]", "[Falha na Geração do Objetivo Específico 3]"]

@st.cache_data(show_spinner="Gerando sugestões de justificativa com IA...")
def generate_justificativa_suggestions_gemini(tema, problema, objetivo_geral, objetivos_especificos, curso, local_foco):
    """Gera uma justificativa completa para o projeto de pesquisa usando IA"""
    prompt = f"""
    Para o pré-projeto de pesquisa com tema '{tema}', problema '{problema}',
    Objetivo Geral '{objetivo_geral}' e Objetivos Específicos '{', '.join(objetivos_especificos)}',
    no curso de {curso}, com foco em {local_foco},
    gere uma justificativa completa. Use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).

    A justificativa deve ser separada em três seções principais obrigatórias: **Pessoal**, **Acadêmica** e **Social**. Cada uma destas seções deve ter no mínimo 3 parágrafos, explicando detalhadamente a relevância e contribuição do estudo sob sua respectiva ótica.

    **Além dessas, avalie criteriosamente se o tema, problema ou objetivos do projeto sugerem a necessidade de outras justificativas relevantes e distintas (por exemplo, Econômica, Ambiental, Metodológica, Política, Ética, Tecnológica, Cultural, etc.). Se identificar que alguma dessas dimensões é pertinentemente impactada ou justificada pelo projeto, crie uma seção adicional específica para cada uma delas.** Cada seção adicional deve ter no mínimo 2 parágrafos ricos em informação, detalhando o impacto e a contribuição específica dessa relevância para o estudo.

    Formate a resposta **EXATAMENTE** assim, usando Markdown para as seções:

    ### JUSTIFICATIVA PESSOAL:
    [Parágrafo 1 - Mínimo 3 parágrafos]
    [Parágrafo 2]
    [Parágrafo 3]
    [...]

    ### JUSTIFICATIVA ACADÊMICA:
    [Parágrafo 1 - Mínimo 3 parágrafos]
    [Parágrafo 2]
    [Parágrafo 3]
    [...]

    ### JUSTIFICATIVA SOCIAL:
    [Parágrafo 1 - Mínimo 3 parágrafos]
    [Parágrafo 2]
    [Parágrafo 3]
    [...]

    [OPCIONAL: ### JUSTIFICATIVA [TIPO ADICIONAL 1 - Ex: ECONÔMICA]:
    [Parágrafo 1 - Mínimo 2 parágrafos]
    [Parágrafo 2]
    [...] ]

    [OPCIONAL: ### JUSTIFICATIVA [TIPO ADICIONAL 2 - Ex: AMBIENTAL]:
    [Parágrafo 1 - Mínimo 2 parágrafos]
    [Parágrafo 2]
    [...] ]

    [Adicione outras seções adicionais se forem relevantes, seguindo o formato acima.]
    """
    try:
        response = model.generate_content(prompt)
        output = response.text.strip()
        
        justificativa_pessoal = extract_block(output, "### JUSTIFICATIVA PESSOAL:", "### JUSTIFICATIVA ACADÊMICA:")
        justificativa_academica = extract_block(output, "### JUSTIFICATIVA ACADÊMICA:", "### JUSTIFICATIVA SOCIAL:")
        justificativa_social = extract_block(output, "### JUSTIFICATIVA SOCIAL:")

        return justificativa_pessoal, justificativa_academica, justificativa_social
    except Exception as e:
        st.error(f"Falha na geração da justificativa: {str(e)}")
        return "[Falha na Geração da Justificativa Pessoal]", "[Falha na Geração da Justificativa Acadêmica]", "[Falha na Geração da Justificativa Social]"

# @st.cache_data(show_spinner="Gerando sugestões de referencial teórico com IA...")
# def generate_referencial_teorico_suggestions_gemini(tema, problema, curso, local_foco, objetivo_geral, objetivos_especificos):

#     prompt = f"""
#     Com base no tema de pesquisa '{tema}', no curso de {curso}, com foco em {local_foco},
#     considerando o problema de pesquisa '{problema}', o objetivo geral '{objetivo_geral}' e os objetivos específicos '{objetivos_especificos}',
#     sugira no mínimo 15 (quinze) tópicos chave ou seções para um referencial teórico robusto, use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).

#     Para cada tópico, faça o seguinte:
#     1.  Forneça uma **breve definição ou contextualização** do tópico.
#     2.  Dentro dessa definição, **cite diretamente 2 (dois) autores ou teorias relevantes**, explicando brevemente sua contribuição ou perspectiva sobre o tópico.
#     3.  Se o tópico exigir dados, regulamentações ou informações específicas de instituições (governamentais, acadêmicas, setoriais, etc.), **sugira até 2 (duas) fontes institucionais pertinentes** que o usuário possa consultar para aprofundar. Caso contrário, não inclua este ponto. Prefira fontes confiáveis e reconhecidas na área de pesquisa, incluindo autores brasileiros e outros do país da área {local_foco}.

#     4. Respeitar a construção frásica correcta, parágrafos e períodos bem estruturados e conjugações verbais correctos.
#     5. A construção frásica deve ser humanizada.
#     6. Evite no máximo conteúdos detectados de IA.

#     Formate a resposta **EXATAMENTE** assim, usando Markdown para facilitar a leitura:

#     ---
#     ### Tópicos para o Referencial Teórico:

#     **1. [Nome do Tópico 1]**
#     [Definição ou contextualização do Tópico 1. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **2. [Nome do Tópico 2]**
#     [Definição ou contextualização do Tópico 2. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **3. [Nome do Tópico 3]**
#     [Definição ou contextualização do Tópico 3. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **4. [Nome do Tópico 4]**
#     [Definição ou contextualização do Tópico 4. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **5. [Nome do Tópico 5]**
#     [Definição ou contextualização do Tópico 5. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **6. [Nome do Tópico 6]**
#     [Definição ou contextualização do Tópico 6. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **7. [Nome do Tópico 7]**
#     [Definição ou contextualização do Tópico 7. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **8. [Nome do Tópico 8]**
#     [Definição ou contextualização do Tópico 8. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **9. [Nome do Tópico 9]**
#     [Definição ou contextualização do Tópico 9. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]

#     **10. [Nome do Tópico 10]**
#     [Definição ou contextualização do Tópico 10. Inclua a citação do Autor/Teoria 1 (Ex: Segundo [Autor 1, Ano], "[citação/ideia]"). Em seguida, insira a citação do Autor/Teoria 2 (Ex: [Autor 2, Ano] aponta que "[citação/ideia]").]
#     [OPCIONAL: **Fontes Institucionais Sugeridas:** [Nome da Instituição 1], [Nome da Instituição 2]]
#     ---
#     """
#     try:
#         response = model.generate_content(prompt)
#         return [re.sub(r'^\d+\.\s*', '', line).strip() for line in response.text.strip().split('\n') if line.strip()]
#     except Exception as e:
#         st.error(f"Falha na geração do referencial teórico: {str(e)}")
#         return ["[Falha na Geração de Tópico 1]", "[Falha na Geração de Tópico 2]"]

@st.cache_data(show_spinner="Gerando referencial teórico com IA e dados reais...")
def generate_referencial_teorico_suggestions_gemini(tema, problema, curso, local_foco, objetivo_geral, objetivos_especificos):
    """Gera referencial teórico combinando IA e buscas reais"""
    
    # 1. Busca no Google Scholar via SerpAPI (se configurado)
    referencias_reais = []
    if SERPAPI_KEY:
        try:
            referencias_reais = buscar_referencias_serpapi(tema)
            st.toast(f"🔍 Encontradas {len(referencias_reais)} referências no Google Scholar")
        except Exception as e:
            st.error(f"Falha na busca SerpAPI: {str(e)}")
    
    # 2. Geração de sugestões pelo Gemini
    with st.spinner("🧠 Gerando sugestões teóricas com IA..."):
        sugestoes_ia = generate_referencial_teorico_suggestions_gemini(
            tema, problema, curso, local_foco, objetivo_geral, objetivos_especificos
        )
    
    # 3. Combinação inteligente dos resultados
    referencial_completo = []
    
    # Adiciona as referências reais encontradas
    for ref in referencias_reais:
        referencial_completo.append(
            f"📚 **{ref['titulo']}**\n"
            f"- Autores: {ref['autores']}\n"
            f"- Ano: {ref['ano']}\n"
            f"- [Acessar]({ref['link']})"
        )
    
    # Adiciona as sugestões da IA (evitando duplicatas)
    autores_reais = {ref['autores'] for ref in referencias_reais}
    for sugestao in sugestoes_ia:
        if not any(autor in sugestao for autor in autores_reais):
            referencial_completo.append(f"💡 {sugestao}")
    
    return referencial_completo

@st.cache_data(show_spinner="Gerando referências bibliográficas com IA...")
def generate_references_gemini(full_text, norma_bibliografica, tema, curso):
    """
    Gera referências bibliográficas analisando todo o texto do pré-projeto.
    """
    citations = extract_citations_from_full_text(full_text)
    
    if not citations:
        return "Nenhuma citação foi detectada no texto para gerar referências."
    
    prompt = f"""
    Com base no pré-projeto sobre '{tema}' para o curso de {curso},
    gere referências bibliográficas completas para todas as citações encontradas no texto.
    
    Norma bibliográfica: {norma_bibliografica}
    
    Citações detectadas: {', '.join(citations)}
    
    Instruções:
    1. Para cada autor detectado, gere uma referência completa.
    2. Inclua todos os elementos necessários conforme a norma especificada.
    3. Quando possível, inferir o tipo de obra (livro, artigo, etc.) com base no contexto.
    4. Organize as referências em ordem alfabética.
    5. Se houver múltiplas obras do mesmo autor, diferencie pelos anos.
    
    Exemplo para ABNT (livro):
    SOBRENOME, Nome. Título da obra. Edição. Local: Editora, ano.
    
    Exemplo para APA (artigo):
    Sobrenome, A., & Sobrenome, B. (ano). Título do artigo. Nome do Periódico, volume(número), páginas.
    
    Retorne apenas as referências formatadas, uma por linha.
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Falha ao gerar referências bibliográficas: {str(e)}")
        return "Não foi possível gerar as referências bibliográficas. Por favor, tente novamente."

@st.cache_data(show_spinner="Gerando sugestões de metodologia com IA...")
def generate_metodologia_suggestions_gemini(objetivo_geral, objetivos_especificos, tema, problema, curso, local_foco):
    prompt = f"""
    Para o projeto de pesquisa com tema '{tema}' no curso de {curso}, com foco em {local_foco},
    e considerando o Problema de pesquisa: '{problema}', Objetivo Geral: '{objetivo_geral}' e
    Objetivos Específicos: {', '.join(objetivos_especificos)},
    elabore uma metodologia completa e detalhada seguindo EXATAMENTE este formato, incluindo justificativas e fundamentação teórica, use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
    Use a construção correcta, respeitando parágrafos, períodos, orações e pontuação certa.
    Evite frases detectadas de IA.
    Humanize o texto todo com expressões e tons humanos

    ### NATUREZA DA PESQUISA:
    [Descreva se a pesquisa é de natureza **Básica** ou **Aplicada**. **Inclua uma breve definição do tipo escolhido, citando pelo menos 2 (dois) autores ou teóricos que fundamentem essa classificação (Ex: Segundo [Autor 1, Ano], pesquisa básica é definida como... Por outro lado, [Autor 2, Ano] aponta que a pesquisa aplicada busca...).**]

    ### JUSTIFICATIVA DA NATUREZA DA PESQUISA:
    [Apresente a justificativa detalhada para a natureza da pesquisa, explicando a escolha. ]

    ### ABORDAGEM DA PESQUISA:
    [Descreva a abordagem da pesquisa (**Qualitativa**, **Quantitativa** ou **Mista**). **Inclua uma breve definição da abordagem escolhida, citando pelo menos 2 (dois) autores ou teóricos que fundamentem essa classificação (Ex: Segundo [Autor 1, Ano], a pesquisa qualitativa busca... Por outro lado, [Autor 2, Ano] descreve a pesquisa quantitativa como...).**]

    ### JUSTIFICATIVA DA ABORDAGEM DA PESQUISA:
    [Apresente a justificativa detalhada para a abordagem da pesquisa, explicando a escolha.]

    ### OBJETIVOS DA PESQUISA (TIPO):
    [Descreva se a pesquisa aos objectivos **Exploratória** ou **Descritiva**  ou **Explicativa** ou **Outra**. **Inclua uma breve definição do tipo escolhido, citando pelo menos 2 (dois) autores ou teóricos que fundamentem essa classificação (Ex: Segundo [Autor 1, Ano], pesquisa básica é definida como... Por outro lado, [Autor 2, Ano] aponta que a pesquisa aplicada busca...).**]
   
    ### FUNDAMENTAÇÃO TEÓRICA DA METODOLOGIA:
    [Apresente a fundamentação teórica que suporta as escolhas metodológicas, citando autores e conceitos relevantes.]

    ### VISÃO DO PESQUISADOR (SUA PERSPECTIVA METODOLÓGICA):
    [Descreva brevemente a sua perspectiva ou crença em relação à metodologia proposta, como pesquisador.]

    ### PROCEDIMENTOS TÉCNICOS:
    - Procedimento 1: detalhe
    - Procedimento 2: detalhe
    - Procedimento 3: detalhe
    - Procedimento 4: detalhe (se aplicável)
    - Procedimento 5: detalhe (se aplicável)

    ### UNIVERSO E AMOSTRA:
    [Descreva o universo de estudo e a população alvo da pesquisa, com foco em {local_foco}.]
    POPULAÇÃO: [Detalhes específicos da população.]
    AMOSTRA: [Descreva a amostra, critérios de seleção e tamanho amostral.]
    TIPO DE AMOSTRA: [Explique o tipo de amostra (ex: probabilística, não probabilística, por conveniência, etc.) e sua justificativa.]

    ### INSTRUMENTOS DE COLETA DE DADOS:
    - Instrumento 1: detalhe
    - Instrumento 2: detalhe
    - Instrumento 3: detalhe (se aplicável)

    ### ANÁLISE DE DADOS:
    [Descreva o método escolhido, sua definição teórica e explique como ele será aplicado especificamente no contexto da sua pesquisa. Fundamente com referências, por exemplo:
    “A Análise de Conteúdo, conforme proposta por Bardin (2011), será utilizada para interpretar qualitativamente os dados textuais obtidos nas entrevistas, organizando-os em categorias temáticas...”]
    - [Se necessário, adicione mais métodos de análise de dados, seguindo o mesmo padrão de descrição e fundamentação teórica.]
 
    ### CONSIDERAÇÕES ÉTICAS:
    [Descreva as considerações éticas relevantes para a pesquisa.]
    """
    try:
        response = model.generate_content(prompt)
        output = response.text.strip()

        suggestions = {
            "natureza": extract_metodologia_block(output, "### NATUREZA DA PESQUISA:", "### JUSTIFICATIVA DA NATUREZA DA PESQUISA:"),
            "natureza_justificativa": extract_metodologia_block(output, "### JUSTIFICATIVA DA NATUREZA DA PESQUISA:", "### ABORDAGEM DA PESQUISA:"),
            "abordagem": extract_metodologia_block(output, "### ABORDAGEM DA PESQUISA:", "### JUSTIFICATIVA DA ABORDAGEM DA PESQUISA:"),
            "abordagem_justificativa": extract_metodologia_block(output, "### JUSTIFICATIVA DA ABORDAGEM DA PESQUISA:", "### OBJETIVOS DA PESQUISA (TIPO):"),
            "objetivos_pesquisa": extract_metodologia_block(output, "### OBJETIVOS DA PESQUISA (TIPO):", "### FUNDAMENTAÇÃO TEÓRICA DA METODOLOGIA:"),
            "fundamentacao_teorica": extract_metodologia_block(output, "### FUNDAMENTAÇÃO TEÓRICA DA METODOLOGIA:", "### VISÃO DO PESQUISADOR (SUA PERSPECTIVA METODOLÓGICA):"),
            "visao_pesquisador": extract_metodologia_block(output, "### VISÃO DO PESQUISADOR (SUA PERSPECTIVA METODOLÓGICA):", "### PROCEDIMENTOS TÉCNICOS:"),
            "procedimentos_tecnicos": extract_metodologia_block(output, "### PROCEDIMENTOS TÉCNICOS:", "### UNIVERSO E AMOSTRA:", as_list=True),
            
            # Extração mais granular de Universo e Amostra
            "universo_amostra_geral": extract_metodologia_block(output, "### UNIVERSO E AMOSTRA:", "POPULAÇÃO:"),
            "populacao": extract_metodologia_block(output, "POPULAÇÃO:", "AMOSTRA:"),
            "amostra": extract_metodologia_block(output, "AMOSTRA:", "TIPO DE AMOSTRA:"),
            "tipo_amostra": extract_metodologia_block(output, "TIPO DE AMOSTRA:", "### INSTRUMENTOS DE COLETA DE DADOS:"),

            "instrumentos_coleta": extract_metodologia_block(output, "### INSTRUMENTOS DE COLETA DE DADOS:", "### ANÁLISE DE DADOS:", as_list=True),
            "analise_de_dados": extract_metodologia_block(output, "### ANÁLISE DE DADOS:", "### CONSIDERAÇÕES ÉTICAS:", as_list=True),
            "consideracoes_eticas": extract_metodologia_block(output, "### CONSIDERAÇÕES ÉTICAS:"),
        }

        # Preenchimento mínimo de segurança e fallback para listas
        if not suggestions["procedimentos_tecnicos"]: suggestions["procedimentos_tecnicos"] = ["Procedimento 1", "Procedimento 2", "Procedimento 3"]
        if not suggestions["instrumentos_coleta"]: suggestions["instrumentos_coleta"] = ["Instrumento 1", "Instrumento 2"]
        if not suggestions["analise_de_dados"]: suggestions["analise_de_dados"] = ["Análise 1", "Análise 2"]

        # Consolidar universo_amostra se os campos granulares estiverem vazios
        if not suggestions["populacao"] and not suggestions["amostra"] and not suggestions["tipo_amostra"]:
            suggestions["universo_amostra"] = suggestions.get('universo_amostra_geral', f"População de {local_foco}. Amostra conforme critérios a definir.")
        else:
            combined_ua = ""
            if suggestions["universo_amostra_geral"]: combined_ua += suggestions["universo_amostra_geral"] + "\n"
            if suggestions["populacao"]: combined_ua += f"POPULAÇÃO: {suggestions['populacao']}\n"
            if suggestions["amostra"]: combined_ua += f"AMOSTRA: {suggestions['amostra']}\n"
            if suggestions["tipo_amostra"]: combined_ua += f"TIPO DE AMOSTRA: {suggestions['tipo_amostra']}\n"
            suggestions["universo_amostra"] = combined_ua.strip()

        return suggestions

    except Exception as e:
        st.error(f"Falha na geração da metodologia: {str(e)}")
        return {
            "natureza": "Pesquisa aplicada", "natureza_justificativa": "Busca resolver problemas práticos.",
            "abordagem": "Qualitativa", "abordagem_justificativa": "Visa compreender fenômenos em profundidade.",
            "objetivos_pesquisa": "Exploratória e descritiva",
            "fundamentacao_teorica": "Baseada em [Autor X] (Ano), [Autor Y] (Ano).",
            "visao_pesquisador": "Visão crítica e reflexiva sobre o objeto de estudo.",
            "procedimentos_tecnicos": ["Revisão bibliográfica", "Coleta de dados de campo", "Análise de dados"],
            "universo_amostra": f"População: [Descrever]. Amostra: [Descrever]. Tipo de Amostra: [Descrever].",
            "populacao": f"[Descrever população]", "amostra": "[Descrever amostra]", "tipo_amostra": "[Descrever tipo]",
            "instrumentos_coleta": ["Entrevistas", "Questionários"],
            "analise_de_dados": ["Análise de conteúdo", "Análise estatística descritiva"],
            "consideracoes_eticas": "Respeito à privacidade, consentimento informado."
        }

@st.cache_data(show_spinner="Gerando cronograma e resultados esperados com IA...")
def generate_cronograma_results_gemini(tema, objetivos_especificos, curso):
    prompt = f"""
    Para o pré-projeto de pesquisa com tema '{tema}', no curso de {curso},
    e com os seguintes Objetivos Específicos: {', '.join(objetivos_especificos)},
    elabore um cronograma de atividades para um período de 6 meses e sugira 3 resultados esperados, use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).

    Formate a resposta EXATAMENTE assim:

    ### CRONOGRAMA (6 MESES):
    Mês 1: [Atividades do Mês 1]
    Mês 2: [Atividades do Mês 2]
    Mês 3: [Atividades do Mês 3]
    Mês 4: [Atividades do Mês 4]
    Mês 5: [Atividades do Mês 5]
    Mês 6: [Atividades do Mês 6]

    ### RESULTADOS ESPERADOS:
    1. [Resultado Esperado 1]
    2. [Resultado Esperado 2]
    3. [Resultado Esperado 3]
    """
    try:
        response = model.generate_content(prompt)
        output = response.text.strip()
        
        cronograma_block = extract_block(output, "### CRONOGRAMA (6 MESES):", "### RESULTADOS ESPERADOS:")
        resultados_block = extract_block(output, "### RESULTADOS ESPERADOS:")
        
        cronograma = {}
        for i in range(1, 7):
            month_match = re.search(rf"Mês {i}: (.*)", cronograma_block)
            if month_match:
                cronograma[f"Mês {i}"] = month_match.group(1).strip()
            else:
                cronograma[f"Mês {i}"] = f"Atividades do Mês {i} (a definir)"

        resultados_esperados = [re.sub(r'^\d+\.\s*', '', line).strip() for i, line in enumerate(resultados_block.split('\n')) if line.strip()]

        return cronograma, resultados_esperados
    except Exception as e:
        st.error(f"Falha na geração do cronograma e resultados: {str(e)}")
        return {f"Mês {i}": "[Falha na Geração]" for i in range(1, 7)}, ["[Falha na Geração de Resultado 1]"]

@st.cache_data(show_spinner="Gerando introdução com IA...")
def generate_introducao_gemini(tema, problema_pesquisa, problematizacao, objetivo_geral, objetivos_especificos, metodologia_resumo):
    prompt = f"""
    Com base nos seguintes elementos do pré-projeto de pesquisa:
    
    Tema: {tema}
    Problema de Pesquisa: {problema_pesquisa}
    Problematização: {problematizacao}
    Objetivo Geral: {objetivo_geral}
    Objetivos Específicos: {', '.join(objetivos_especificos)}
    Metodologia (resumo): {metodologia_resumo}

    Escreva uma introdução completa para o pré-projeto de pesquisa, seguindo esta estrutura:
    
    1. Contextualização do tema: Apresente o tema geral da pesquisa e sua relevância.
    2. Problema de pesquisa: Introduza o problema específico que será investigado.
    3. Objetivos: Apresente o objetivo geral e os específicos de forma integrada.
    4. Justificativa breve: Explique rapidamente por que este estudo é importante.
    5. Metodologia: Descreva brevemente a abordagem metodológica.
    6. Organização do trabalho:(ex:. O presente trabalho de pesquisa tem como estrutura as seções de introdução, Metodologia, Referencial Teórico e Conclusão., 7. Inclua todos os tópicos de heading 1 do documento, exceto Referencias Biblograficas, apêndices e anexos, que deve ser escrita separadamente.)

    8. A introdução deve ter no mínimo 5 parágrafos bem redigidos, com linguagem acadêmica formal e coesão entre as ideias.
    9. Use a construção frasica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
    10. Nunca use conteudos detectados de IA.
    11. Humanize todas a costrução frásica. 

    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Falha ao gerar a introdução: {str(e)}")
        return "Não foi possível gerar a introdução automaticamente. Por favor, escreva manualmente."

@st.cache_data(show_spinner="Gerando conclusão com IA...")
def generate_conclusao_gemini(tema, problema, objetivo_geral, objetivos_espec, resultados, contexto):
    """
    Gera uma conclusão acadêmica completa.
    """
    nivel_academico = st.session_state.get('nivel_academico', 'Graduação')
    prompt = f"""
    Com base no pré-projeto com os seguintes elementos:
    
    TEMA: {tema}
    PROBLEMA: {problema}
    OBJETIVO GERAL: {objetivo_geral}
    OBJETIVOS ESPECÍFICOS: {', '.join(objetivos_espec)}
    RESULTADOS ESPERADOS: {', '.join(resultados)}
    
    CONTEXTO ADICIONAL:
    - Contribuições esperadas: {contexto.get('contribuicoes', 'Não especificado')}
    - Limitações: {contexto.get('limitacoes', 'Não especificado')}
    - Pesquisas futuras: {contexto.get('pesquisas_futuras', 'Não especificado')}
    
    Escreva uma conclusão acadêmica bem estruturada contendo:
    
    1. Síntese dos objetivos e metodologia (1 parágrafo)
    2. Contribuições acadêmicas e práticas (1-2 parágrafos)
    3. Limitações do estudo (1 parágrafo)
    4. Sugestões para pesquisas futuras (1 parágrafo)
    
    Nível acadêmico: {nivel_academico}
    Formato: Linguagem acadêmica formal, 4-6 parágrafos coesos, use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção)

    Use a construção frásica correcta respeitando paragrafos, periodos, orações e pontuacao correcta, usando tons humanos.
    Nunca use conteúdos detectados de IA.
    Humanize todas a costrução frásica. 
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Erro ao gerar conclusão: {str(e)}")
        return "Não foi possível gerar a conclusão automaticamente. Por favor, escreva manualmente."

@st.cache_data(show_spinner="Gerando o instrumento de coleta de dados com IA...")
def generate_data_collection_instrument_gemini(tema, problema, objetivo_geral, objetivos_especificos, tipo_instrumento, curso, local_foco):
    """
    Gera um guia de entrevista ou questionário com base nos detalhes do projeto.
    """
    objetivos_str = "- " + "\n- ".join(objetivos_especificos) if objetivos_especificos else "Não fornecidos."

    prompt = f"""
    Com base no pré-projeto de pesquisa:

    Tema: {tema}
    Problema de Pesquisa: {problema}
    Objetivo Geral: {objetivo_geral}
    Objetivos Específicos:
    {objetivos_str}
    Curso: {curso}
    Local de Foco: {local_foco}

    Crie um **{tipo_instrumento}** detalhado para a coleta de dados.

    Se for um **Guia de Entrevista**, inclua:
    - Uma breve introdução para o entrevistado.
    - Seções temáticas com perguntas abertas e claras.
    - Perguntas que ajudem a alcançar cada objetivo específico.
    - Perguntas sobre o problema de pesquisa.
    - Uma saudação de encerramento.

    Se for um **Questionário**, inclua:
    - Uma breve introdução para o respondente e instruções claras.
    - Seções (ex: dados demográficos, perguntas sobre o tema).
    - Tipos de perguntas variados (múltipla escolha, escala Likert, abertas) quando apropriado.
    - Perguntas que ajudem a alcançar cada objetivo específico.
    - Perguntas sobre o problema de pesquisa.
    - Uma mensagem de agradecimento.

    O instrumento deve ser relevante, claro, objetivo e adequado ao contexto acadêmico.
    Apresente-o de forma estruturada e fácil de ler, use sempre o pré-acordo ortográfico nas palavras (Ex: objetivo deve ser objectivo, correção deve ser correcção).
    """

    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        st.error(f"Falha ao gerar o {tipo_instrumento.lower()}: {str(e)}")
        return f"Não foi possível gerar o {tipo_instrumento.lower()} automaticamente. Por favor, tente novamente."

def add_table_of_contents(document, doc_content):
    """
    Adiciona um índice automático COMPLETO ao documento Word, sem depender do campo TOC.
    
    Args:
        document: O documento Word (objeto Document do python-docx)
        doc_content: Lista de dicionários com a estrutura do documento
        
    Exemplo de doc_content:
    [
        {"title": "Introdução", "level": 1, "page": 2},
        {"title": "Problema de Pesquisa", "level": 1, "page": 3},
        {"title": "Justificativa", "level": 1, "page": 4},
        {"title": "Justificativa Pessoal", "level": 2, "page": 4},
        ...
    ]
    """
    # Adiciona página do índice
    document.add_page_break()
    
    # Título do índice
    title = document.add_heading('ÍNDICE', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph()  # Espaço
    
    # Estilo para os itens do índice
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    # Adiciona cada item do índice
    for item in doc_content:
        para = document.add_paragraph(style='Normal')
        
        # Nível de indentação
        indent = 0.5 * (item['level'] - 1)  # 0.5 polegadas por nível
        
        # Texto do item
        run = para.add_run(" " * 4 * (item['level'] - 1) + item['title'])
        
        # Tabulação para alinhar o número da página à direita
        para.paragraph_format.tab_stops.add_tab_stop(Inches(6), WD_ALIGN_PARAGRAPH.RIGHT)
        
        # Adiciona tabulação e número da página
        para.add_run("\t" + str(item['page']))
        
        # Alinhamento justificado (texto à esquerda, número à direita)
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.left_indent = Inches(indent)
    
    document.add_paragraph()  # Espaço final
    return document
    
 
def create_word_document(data):
    try:
        document = Document()
        # Estilo padrão
        style = document.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)

        # Título do Pré-Projeto
        document.add_heading('Pré-Projeto de Pesquisa', level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph().add_run(f"Tema: {data.get('tema', 'Não Definido')}").bold = True
        document.add_paragraph().add_run(f"Curso: {data.get('curso', 'Não Definido')}").bold = True
        document.add_paragraph().add_run(f"Local de Foco: {data.get('local_foco', 'Não Definido')}").bold = True
        document.add_paragraph(f"Data de Geração: {datetime.today().date().strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_page_break()

      
        doc_structure = []
    
    # 1. Capa (não entra no índice)
       
        current_page = 3  # Começa na página 3 (capa e índice ocupam 1-2)
    
    # Adiciona cada seção ao doc_structure
        doc_structure.append({"title": "Introdução", "level": 1, "page": current_page})
    # Continue para todas as seções...
    
    # Cria o índice com as informações coletadas
        document = add_table_of_contents(document, doc_structure)
        
        # Restante do documento...
        document.add_page_break()

        # 1. Introdução
        document.add_heading('1. Introdução', level=1)
        document.add_paragraph(data.get('introducao', 'Conteúdo da introdução aqui. Esta seção será preenchida com uma visão geral do seu projeto.'))
        document.add_page_break()
        # 2. Problema de Pesquisa
        document.add_heading('1.1. Problema de Pesquisa', level=2)
        if data.get('problema_pesquisa'):
            document.add_paragraph(data['problema_pesquisa'])
        else:
            document.add_paragraph("Problema de Pesquisa não definido.")

        # 3. Problematização
        document.add_heading('1.2. Problematização', level=2)
        if data.get('problematizacao'):
            document.add_paragraph(data['problematizacao'])
        else:
            document.add_paragraph("Problematização não definida.")
        document.add_page_break()
        # 4. Perguntas de Pesquisa / Hipóteses
        document.add_heading('1.3. Perguntas de Pesquisa / Hipóteses', level=2)
        if data.get('tipo_inquerito_pesquisa') == "Perguntas de Pesquisa" and data.get('perguntas_pesquisa'):
            document.add_heading('1.3.1 Perguntas de Pesquisa', level=3)
            for i, pergunta in enumerate(data['perguntas_pesquisa']):
                document.add_paragraph(f"{i+1}. {pergunta}")
        elif data.get('tipo_inquerito_pesquisa') == "Hipóteses" and data.get('hipoteses'):
            document.add_heading('1.3.2 Hipóteses', level=3)
            for i, hipotese in enumerate(data['hipoteses']):
                document.add_paragraph(f"{i+1}. {hipotese}")
        else:
            document.add_paragraph("Perguntas de Pesquisa ou Hipóteses não definidas.")
        document.add_page_break()
        # 5. Objetivos
        document.add_heading('1.4. Objetivos', level=2)
        document.add_heading('1.4.1 Objetivo Geral', level=3)
        document.add_paragraph(data.get('objetivo_geral', 'Objetivo Geral não definido.'))
        document.add_heading('1.4.2 Objetivos Específicos', level=3)
        for obj in data.get('objetivos_especificos', []):
            document.add_paragraph(f"- {obj}")
        document.add_page_break()
        # 6. Justificativa
        document.add_heading('1.5. Justificativa', level=2)
        if data.get('justificativa_pessoal'):
            document.add_heading('1.5.1 Justificativa Pessoal', level=3)
            document.add_paragraph(data['justificativa_pessoal'])
        if data.get('justificativa_academica'):
            document.add_heading('1.5.2 Justificativa Acadêmica', level=3)
            document.add_paragraph(data['justificativa_academica'])
        if data.get('justificativa_social'):
            document.add_heading('1.5.3 Justificativa Social', level=3)
            document.add_paragraph(data['justificativa_social'])
        document.add_page_break()          
       
        # 8. Metodologia
        document.add_heading('2. Metodologia', level=1)
        metodologia_data = data.get('metodologia_sugestoes', {})

        if metodologia_data.get('natureza'):
            document.add_heading('2.1 Natureza da Pesquisa', level=2)
            document.add_paragraph(metodologia_data['natureza'])
        if metodologia_data.get('natureza_justificativa'):
            document.add_heading('2.1.1 Justificativa da Natureza', level=3)
            document.add_paragraph(metodologia_data['natureza_justificativa'])

        if metodologia_data.get('abordagem'):
            document.add_heading('2.2 Abordagem da Pesquisa', level=2)
            document.add_paragraph(metodologia_data['abordagem'])
        if metodologia_data.get('abordagem_justificativa'):
            document.add_heading('2.2.1 Justificativa da Abordagem', level=3)
            document.add_paragraph(metodologia_data['abordagem_justificativa'])

        if metodologia_data.get('objetivos_pesquisa'):
            document.add_heading('2.3 Objetivos da Pesquisa (Tipo)', level=2)
            document.add_paragraph(metodologia_data['objetivos_pesquisa'])

        if metodologia_data.get('fundamentacao_teorica'):
            document.add_heading('2.4 Fundamentação Teórica da Metodologia', level=2)
            document.add_paragraph(metodologia_data['fundamentacao_teorica'])

        if metodologia_data.get('visao_pesquisador'):
            document.add_heading('2.5 Visão do Pesquisador', level=2)
            document.add_paragraph(metodologia_data['visao_pesquisador'])

        if metodologia_data.get('procedimentos_tecnicos'):
            document.add_heading('2.6 Procedimentos Técnicos', level=2)
            for proc in metodologia_data['procedimentos_tecnicos']:
                document.add_paragraph(f"- {proc}")

        if metodologia_data.get('universo_amostra'):
            document.add_heading('2.7 Universo e Amostra', level=2)
            document.add_paragraph(metodologia_data['universo_amostra'])

        if metodologia_data.get('analise_de_dados'):
            document.add_heading('2.8 Análise de Dados', level=2)
            for analise in metodologia_data['analise_de_dados']:
                document.add_paragraph(f"- {analise}")

        if metodologia_data.get('consideracoes_eticas'):
            document.add_heading('2.9 Considerações Éticas', level=2)
            document.add_paragraph(metodologia_data['consideracoes_eticas'])

        document.add_page_break()


 # 3. Referencial Teórico
        document.add_heading('3. Referencial Teórico', level=1)
        if data.get('referencial_teorico_sugestoes'):
            for i, topico in enumerate(data['referencial_teorico_sugestoes']):
                document.add_paragraph(f"- {topico}")
        else:
            document.add_paragraph("Tópicos do referencial teórico aqui.")
        document.add_page_break()



        document.add_heading('4 Resultados Esperados', level=1)
        if data.get('cronograma_results_sugestoes', {}).get('resultados_esperados'):
            for i, resultado in enumerate(data['cronograma_results_sugestoes']['resultados_esperados']):
                document.add_paragraph(f"{i+1}. {resultado}")
        else:
            document.add_paragraph("Resultados esperados não definidos.")
        document.add_page_break()
        # 10. Conclusão
        document.add_heading('5. Conclusão', level=1)
        document.add_paragraph(data.get('conclusao', 'Conclusão provisória do pré-projeto.'))
        document.add_page_break()

        # 11. Referências Bibliográficas
        document.add_heading('6. Referências Bibliográficas', level=1)
        if data.get('referencias_bibliograficas'):
            for ref_line in data['referencias_bibliograficas'].split('\n'):
                if ref_line.strip():
                    document.add_paragraph(ref_line.strip())
        else:
            document.add_paragraph("Referências bibliográficas não geradas ou vazias.")
        document.add_page_break()
        # 12. Anexos (se houver)
        
     
    

        # 9. Cronograma e Resultados Esperados
        document.add_heading('7. Cronograma', level=1)
#document.add_heading('9.1 Cronograma', level=2)
        cronograma_data = data.get('cronograma_results_sugestoes', {}).get('cronograma', {})
        if cronograma_data:
            table = document.add_table(rows=1, cols=len(cronograma_data) + 1)
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Atividade'
            meses_sorted = sorted(cronograma_data.keys(), key=lambda x: int(x.split(' ')[1]))
            for i, mes in enumerate(meses_sorted):
                hdr_cells[i+1].text = mes

            atividades_unicas = set()
            for mes, atividades_str in cronograma_data.items():
                for atividade in atividades_str.split(';'):
                    if atividade.strip():
                        atividades_unicas.add(atividade.strip())
            
            for atividade in sorted(list(atividades_unicas)):
                row_cells = table.add_row().cells
                row_cells[0].text = atividade
                for i, mes in enumerate(meses_sorted):
                    if atividade in cronograma_data[mes]:
                        row_cells[i+1].text = 'X'
                    else:
                        row_cells[i+1].text = ''
        else:
            document.add_paragraph("Cronograma não gerado.")




            current_table = []
            in_table = False

            for line in conteudo.split('\n'):
                line = line.strip()
                if not line:
                    if in_table:
                        current_table.append([""])
                    else:
                        document.add_paragraph('')
                    continue

                # Tabela Markdown
                if "|" in line and "-|-" not in line:
                    in_table = True
                    row = [cell.strip() for cell in line.split("|") if cell.strip()]
                    current_table.append(row)
                    continue

                if in_table and "|" not in line:
                    in_table = False
                if current_table:
                    try:
                        table = document.add_table(rows=1, cols=len(current_table[0]))
                        table.style = 'Table Grid'
                        hdr_cells = table.rows[0].cells
                        for j, header in enumerate(current_table[0]):
                            hdr_cells[j].text = header
                        for row in current_table[1:]:
                            row_cells = table.add_row().cells
                            for j, cell in enumerate(row):
                                if j < len(row_cells):
                                    row_cells[j].text = cell
                    except Exception as e:
                        st.error(f"Erro ao criar tabela: {str(e)}")
                        for row in current_table:
                            document.add_paragraph(" | ".join(row))
                current_table = []

                # Listas numeradas
                if re.match(r'^\d+\.', line):
                    p = document.add_paragraph(style='List Number')
                    p.add_run(line.split('.', 1)[1].strip())
                    continue

                # Listas com marcadores
                if line.startswith('- ') or line.startswith('* '):
                    p = document.add_paragraph(style='List Bullet')
                    p.add_run(line[2:].strip())
                    continue

                # Subtítulos markdown
                for level, prefix in [(4, '#### '), (3, '### '), (2, '## '), (1, '# ')]:
                    if line.startswith(prefix):
                        document.add_heading(line[len(prefix):].strip(), level=level)
                        break
                else:
                # Negrito/Itálico markdown
                    p = document.add_paragraph()
                    bold_italic_pattern = r'(\*\*\*.+?\*\*\*|\*\*.+?\*\*|\*.+?\*)'
                    last_idx = 0
                for match in re.finditer(bold_italic_pattern, line):
                    if match.start() > last_idx:
                        p.add_run(line[last_idx:match.start()])
                    text = match.group(0)
                    run = p.add_run(text.strip('*'))
                    if text.startswith('***'):
                        run.bold = True
                        run.italic = True
                    elif text.startswith('**'):
                        run.bold = True
                    elif text.startswith('*'):
                        run.italic = True
                    last_idx = match.end()
                if last_idx < len(line):
                    p.add_run(line[last_idx:])
       

        # --- Adiciona instrumentos gerados como apêndices ---

        instrumentos = []
        # Suporte tanto para lista quanto para dict (compatibilidade)
        if isinstance(st.session_state.get("todos_instrumentos_gerados", None), dict):
            for k, v in st.session_state["todos_instrumentos_gerados"].items():
                instrumentos.append({"titulo": k, "conteudo": v})
        elif isinstance(st.session_state.get("instrumentos_gerados", None), list):
            instrumentos = st.session_state["instrumentos_gerados"]

        if instrumentos:
            document.add_page_break()
            document.add_heading("8. APÊNDICES ", level=1)

            for i, inst in enumerate(instrumentos):
                letra = string.ascii_uppercase[i % 26]  # A, B, C, ...
            titulo = inst.get("titulo", "Instrumento de Coleta")
            conteudo = inst.get("conteudo", "")

            # Título do apêndice
            # document.add_heading(f"APÊNDICE {letra} – {titulo}", level=2)
            # document.add_paragraph("Este apêndice contém o instrumento de coleta de dados utilizado na pesquisa.")
            document.add_heading(titulo, level=3)

            document.add_heading('9. Anexos', level=1) 
        if data.get('anexos'):
            for i, anexo in enumerate(data['anexos']):
                document.add_paragraph(f"{i+1}. {anexo}")   
        else:
            document.add_paragraph("Nenhum anexo adicionado.")
        # 13. Coleta de Dados
        document.add_page_break()


   # 14. Revisão Final

        bio = BytesIO()
        document.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        st.error(f"Erro ao criar documento Word: {str(e)}")
        return None

# --- Interface principal ---
def main():
    initialize_session_state()
    st.set_page_config(page_title="AcadêmicoPro", page_icon="🎓", layout="wide")
    st.title("🎓 AcadêmicoPro")
    

    passos_labels = {
    1: "1. Início (Tema)",
    2: "2. Problema",
    3: "3. Problematização",
    4: "4. Objectivos",
    5: "5. Hipóteses",
    6: "6. Justificativa",
    7: "7. Referencial Teórico",
    8: "8. Metodologia",
    9: "9. Cronograma",
    10: "10. Introdução",
    11: "11. Conclusão",
    12: "12. Referências",
    13: "13. Colecta de Dados",
    14: "14. Revisão Final",
    15: "15. Download Final"
}
    with st.sidebar:
        st.title("Navegação")

        # Garante que 'passo' exista no session_state, começando pelo passo 1
        if "passo" not in st.session_state:
            st.session_state.passo = 1

        # Cria uma lista dos rótulos para usar com st.radio
        opcoes_radio = list(passos_labels.values())

        # Determina o índice da opção atualmente selecionada para st.radio
        # Isso garante que o radio button correto seja pré-selecionado ao recarregar a página
        indice_selecionado = 0
        if st.session_state.passo in passos_labels:
            label_atual = passos_labels[st.session_state.passo]
            if label_atual in opcoes_radio:
                indice_selecionado = opcoes_radio.index(label_atual)

        # Usa st.radio para criar os botões de rádio
        secao_selecionada_label = st.radio(
            "Selecione a Etapa:",
            options=opcoes_radio,
            index=indice_selecionado,
            key="navegacao_radio" # Uma chave única é sempre boa prática
        )

        # Encontra o número do passo correspondente ao rótulo selecionado
        # e atualiza st.session_state.passo se a seleção mudar
        for passo_num, label in passos_labels.items():
            if label == secao_selecionada_label:
                if st.session_state.passo != passo_num:
                    st.session_state.passo = passo_num
                    st.rerun() # Recarrega a página para refletir a nova seção

        st.markdown("---")
        st.info("Preencha as informações em cada etapa. A IA o(a) ajudará com sugestões!")

    # Exemplo de como você usaria o st.session_state.passo para exibir o conteúdo da seção
    st.write(f"Você está na seção: **{passos_labels[st.session_state.passo]}**")
    # Aqui você adicionaria a lógica para mostrar o conteúdo específico de cada passo
    # Por exemplo, um 'if' statement ou uma função que renderiza o conteúdo com base em st.session_state.passo
    
    if st.session_state.passo == 1:
        st.subheader("1. 🎯 Tema da Pesquisa")

        col1, col2 = st.columns(2)
        with col1:
            area_interesse = st.text_input(
                "Área de interesse: (Ex:.Saúde, Educação, Tecnologia, etc.)",
                key="area_interesse_input",
                value=st.session_state.get("area_interesse", "")
            )

            curso = st.text_input(
                "Seu curso/formação:",
                key="curso_input",
                value=st.session_state.get("curso", "")
            )
        with col2:
            palavras_chave = st.text_input(
                "Palavras-chave (separe por vírgulas):",
                key="palavras_chave_input",
                value=st.session_state.get("palavras_chave", "")
            )
            nivel_academico = st.selectbox(
                "Nível Acadêmico:",
                ["Licenciatura", "Mestrado", "Doutorado", "Pós-doc"],
                key="nivel_academico_select",
                index=["Licenciatura", "Mestrado", "Doutorado", "Pós-doc"].index(
                    st.session_state.get("nivel_academico", "Licenciatura")
                )
            )

        if st.button("✨ Sugerir Temas com IA"):
            if area_interesse and curso and palavras_chave:
                st.session_state.tema_suggestions = generate_tema_suggestions_gemini(
                    area_interesse, curso, palavras_chave, nivel_academico
                )
            else:
                st.warning("Preencha todos os campos para gerar sugestões")

        if 'tema_suggestions' in st.session_state and st.session_state.tema_suggestions:
            st.markdown("---")
            selected_suggestion = st.radio("Escolha uma sugestão ou edite abaixo:", st.session_state.tema_suggestions)
            st.session_state.tema = selected_suggestion.replace(
                f"{st.session_state.tema_suggestions.index(selected_suggestion)+1}. ", "")

        st.markdown("---")
        st.session_state.tema = st.text_input(
            "Tema final:",
            value=st.session_state.get("tema", "")
        )

        st.session_state.curso = curso
        st.session_state.area_interesse = area_interesse
        st.session_state.palavras_chave = palavras_chave
        st.session_state.nivel_academico = nivel_academico
        st.session_state.area_interesse = area_interesse
        st.session_state.palavras_chave = palavras_chave
        st.session_state.curso = curso
        local_foco = st.text_input(
            "Local/Contexto de Foco:",
            value=st.session_state.get("local_foco", "")
        )
        st.session_state.local_foco = local_foco

        col1_nav, col2_nav = st.columns(2)
        with col2_nav:
            if st.button("Próximo ➡️") and st.session_state.tema and st.session_state.curso and st.session_state.local_foco:
                st.session_state.passo = 2
                st.rerun()
            else:
                st.warning("Por favor, preencha o Tema, Curso e Local de Foco antes de prosseguir.")

    elif st.session_state.passo == 2:
        st.subheader("2. ❓ Problema de Pesquisa")
        st.write(f"Qual é a questão central que sua pesquisa busca responder? (Tema: **{st.session_state.tema}**)")

        if st.session_state.objetivo_geral:
            st.info(f"Objetivo Geral atual (para referência): {st.session_state.objetivo_geral}")

        if st.button("✨ Sugerir Problemas com IA", key="sugerir_problemas_button"):
            if st.session_state.tema and st.session_state.curso and st.session_state.local_foco:
                temp_objetivo_geral = st.session_state.objetivo_geral if st.session_state.objetivo_geral else f"Analisar {st.session_state.tema}"
                st.session_state.problema_pesquisa_sugestoes = generate_problema_suggestions_gemini(
                    st.session_state.tema,
                    st.session_state.curso,
                    st.session_state.local_foco,
                    temp_objetivo_geral
                )
            else:
                st.warning("Por favor, preencha Tema, Curso e Local de Foco primeiro.")

        if st.session_state.problema_pesquisa_sugestoes:
            st.markdown("---")
            st.subheader("Sugestões de Problema de Pesquisa:")
            selected_problema_suggestion = st.radio(
                "Escolha uma sugestão ou edite abaixo:",
                st.session_state.problema_pesquisa_sugestoes,
                key="problema_suggestion_radio"
            )
            st.session_state.problema_pesquisa = selected_problema_suggestion

        st.markdown("---")
        st.subheader("Seu Problema de Pesquisa Escolhido:")
        st.session_state.problema_pesquisa = st.text_area(
            "Formule seu problema de pesquisa aqui (idealmente em formato de pergunta):",
            value=st.session_state.problema_pesquisa,
            height=100,
            help="Ex: 'Quais são os impactos da inteligência artificial na educação primária em Moçambique?'"
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo2_prev"):
                st.session_state.passo = 1
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo2_next"):
                if st.session_state.problema_pesquisa:
                    st.session_state.passo = 3
                    st.rerun()
                else:
                    st.warning("Por favor, defina o Problema de Pesquisa antes de prosseguir.")

    elif st.session_state.passo == 3:
        st.subheader("3. ✍️ Problematização")
        st.write("Desenvolva a relevância e o contexto do seu problema de pesquisa, explicando por que ele precisa ser investigado.")

        st.info(f"Problema de Pesquisa: **{st.session_state.problema_pesquisa}**")

        if st.button("✨ Gerar Problematização com IA", key="gerar_problematizacao_button"):
            if st.session_state.problema_pesquisa:
                with st.spinner("Gerando problematização..."):
                    st.session_state.problematizacao = generate_problematization_gemini(
                        st.session_state.problema_pesquisa
                    )
                st.success("Problematização gerada com sucesso!")
            else:
                st.error("Por favor, defina o Problema de Pesquisa primeiro.")

        st.markdown("---")
        st.subheader("Edite a Problematização:")
        st.session_state.problematizacao = st.text_area(
            "Problematização da Pesquisa:",
            value=st.session_state.problematizacao,
            height=300,
            help="Detalhe o contexto, a relevância e as implicações do seu problema de pesquisa. Esta seção deve convencer o leitor da importância do seu estudo."
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.passo = 2
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.passo = 4
                st.rerun()

    elif st.session_state.passo == 5:
        st.subheader("5. 🔍 Perguntas de Pesquisa ou Hipóteses")
        st.write("Defina as perguntas mais específicas que sua pesquisa responderá, ou as hipóteses que testará.")
        ''
        st.info(f"Tipo de inquérito atual: **{st.session_state.tipo_inquerito_pesquisa}**")

        st.info(f"Problema de Pesquisa: **{st.session_state.problema_pesquisa}**")
        st.info(f"Objetivo Geral: **{st.session_state.objetivo_geral if st.session_state.objetivo_geral else 'Ainda não definido, pode ser gerado no próximo passo.'}**")
        
        if st.session_state.objetivos_especificos:
            st.markdown("---")
            st.subheader("Objectivos Específicos para alinhamento:")
            for i, obj in enumerate(st.session_state.objetivos_especificos):
                st.write(f"- {obj}")
            st.markdown("---")

        # st.session_state.tipo_inquerito_pesquisa = st.radio(
        #     "O que você deseja gerar?",
        #     ("Perguntas de Pesquisa", "Hipóteses"),
        #     key="tipo_inquerito_radio"
        # )
      
        # st.radio(
        #     "O que você deseja gerar?",
        #     ("Perguntas de Pesquisa", "Hipóteses"),
        #     key="tipo_inquerito_pesquisa"
        # )


        # Define a lista de opções
        opcoes_inquerito = ("Perguntas de Pesquisa", "Hipóteses")

        # Verifica se já temos um valor carregado
        valor_salvo = st.session_state.get("tipo_inquerito_pesquisa", "Perguntas de Pesquisa")

        # Define qual índice estará selecionado (0 ou 1)
        index_selecionado = opcoes_inquerito.index(valor_salvo) if valor_salvo in opcoes_inquerito else 0

        # Exibe o radio corretamente com a opção carregada marcada
        tipo_escolhido = st.radio(
            "O que você deseja gerar?",
            opcoes_inquerito,
            index=index_selecionado,
            key="tipo_inquerito_pesquisa_radio"
        )

        # Atualiza a session_state manualmente, se quiser manter compatibilidade com outras partes do sistema
        st.session_state["tipo_inquerito_pesquisa"] = tipo_escolhido





        # Depois você pode acessar com:
        tipo_escolhido = st.session_state.tipo_inquerito_pesquisa
        st.write(f"Você escolheu: {tipo_escolhido}")



        if st.button(f"✨ Gerar {st.session_state.tipo_inquerito_pesquisa} com IA", key="gerar_inquerito_button"):
            if not st.session_state.problema_pesquisa:
                st.error("Por favor, defina o Problema de Pesquisa (Passo 2) antes de gerar.")
            elif not st.session_state.objetivo_geral or not st.session_state.objetivos_especificos:
                st.warning("O Objetivo Geral e os Objetivos Específicos (Passo 5) ainda não foram definidos ou não foram passados. A geração será baseada apenas no problema. Considere definir os objetivos primeiro para um melhor alinhamento.")
                og_temp = st.session_state.objetivo_geral if st.session_state.objetivo_geral else "Um objetivo geral provisório."
                oe_temp = st.session_state.objetivos_especificos if st.session_state.objetivos_especificos else [f"Um objetivo específico provisório para {st.session_state.problema_pesquisa}"]

                with st.spinner(f"Gerando {st.session_state.tipo_inquerito_pesquisa.lower()}..."):
                    if st.session_state.tipo_inquerito_pesquisa == "Perguntas de Pesquisa":
                        st.session_state.perguntas_pesquisa = generate_inquiry_gemini(
                            st.session_state.problema_pesquisa, og_temp, oe_temp, "Perguntas de Pesquisa"
                        )
                        st.session_state.hipoteses = []
                    else:
                        st.session_state.hipoteses = generate_inquiry_gemini(
                            st.session_state.problema_pesquisa, og_temp, oe_temp, "Hipóteses"
                        )
                        st.session_state.perguntas_pesquisa = []
                st.success(f"{st.session_state.tipo_inquerito_pesquisa} geradas com sucesso!")
            else:
                with st.spinner(f"Gerando {st.session_state.tipo_inquerito_pesquisa.lower()}..."):
                    if st.session_state.tipo_inquerito_pesquisa == "Perguntas de Pesquisa":
                        st.session_state.perguntas_pesquisa = generate_inquiry_gemini(
                            st.session_state.problema_pesquisa, st.session_state.objetivo_geral, st.session_state.objetivos_especificos, "Perguntas de Pesquisa"
                        )
                        st.session_state.hipoteses = []
                    else:
                        st.session_state.hipoteses = generate_inquiry_gemini(
                            st.session_state.problema_pesquisa, st.session_state.objetivo_geral, st.session_state.objetivos_especificos, "Hipóteses"
                        )
                        st.session_state.perguntas_pesquisa = []
                st.success(f"{st.session_state.tipo_inquerito_pesquisa} geradas com sucesso!")

        st.markdown("---")
        st.subheader(f"Edite suas {st.session_state.tipo_inquerito_pesquisa}:")
        
        if st.session_state.tipo_inquerito_pesquisa == "Perguntas de Pesquisa":
            perguntas_str = "\n".join(st.session_state.perguntas_pesquisa)
            edited_perguntas_str = st.text_area(
                "Lista de Perguntas de Pesquisa (uma por linha):",
                value=perguntas_str,
                height=200,
                help="Cada pergunta deve ser uma linha separada. Elas devem ser específicas e alinhadas aos seus objectivos."
            )
            st.session_state.perguntas_pesquisa = [line.strip() for line in edited_perguntas_str.split('\n') if line.strip()]
        else:
            hipoteses_str = "\n".join(st.session_state.hipoteses)
            edited_hipoteses_str = st.text_area(
                "Lista de Hipóteses (uma por linha):",
                value=hipoteses_str,
                height=200,
                help="Cada hipótese deve ser uma linha separada. Devem ser declarações testáveis sobre a relação entre variáveis."
            )
            st.session_state.hipoteses = [line.strip() for line in edited_hipoteses_str.split('\n') if line.strip()]

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo4_prev"):
                st.session_state.passo = 4
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo4_next"):
                if (st.session_state.tipo_inquerito_pesquisa == "Perguntas de Pesquisa" and st.session_state.perguntas_pesquisa) or \
                (st.session_state.tipo_inquerito_pesquisa == "Hipóteses" and st.session_state.hipoteses):
                    st.session_state.passo = 6
                    st.rerun()
                else:
                    st.warning(f"Por favor, gere ou insira suas {st.session_state.tipo_inquerito_pesquisa} antes de prosseguir.")

    elif st.session_state.passo == 4:
        st.subheader("4. 🎯 Objectivos")
        st.write(f"Defina o que sua pesquisa pretende alcançar. (Tema: **{st.session_state.tema}**)")
        st.info(f"Problema de Pesquisa: **{st.session_state.problema_pesquisa}**")

        if st.session_state.tipo_inquerito_pesquisa == "Perguntas de Pesquisa" and st.session_state.perguntas_pesquisa:
            st.markdown("---")
            st.subheader("Perguntas de Pesquisa para alinhamento:")
            for i, q in enumerate(st.session_state.perguntas_pesquisa):
                st.write(f"- {q}")
            st.markdown("---")
        elif st.session_state.tipo_inquerito_pesquisa == "Hipóteses" and st.session_state.hipoteses:
            st.markdown("---")
            st.subheader("Hipóteses para alinhamento:")
            for i, h in enumerate(st.session_state.hipoteses):
                st.write(f"- {h}")
            st.markdown("---")

        if st.button("✨ Sugerir Objectivos com IA", key="sugerir_objetivos_button"):
            if st.session_state.tema and st.session_state.problema_pesquisa and st.session_state.curso and st.session_state.local_foco:
                objetivo_geral_sug, objetivos_especificos_sug = generate_objetivos_suggestions_gemini(
                    st.session_state.tema,
                    st.session_state.problema_pesquisa,
                    st.session_state.curso,
                    st.session_state.local_foco
                )
                st.session_state.objetivo_geral = objetivo_geral_sug
                st.session_state.objetivos_especificos = objetivos_especificos_sug
                st.success("Objetivos sugeridos com sucesso!")
            else:
                st.warning("Por favor, preencha Tema, Problema de Pesquisa, Curso e Local de Foco.")

        st.markdown("---")
        st.subheader("Seus Objectivos:")
        st.session_state.objetivo_geral = st.text_input(
            "Objectivo Geral (Ex: 'Analisar o impacto de X em Y'):",
            value=st.session_state.objetivo_geral,
            help="O que sua pesquisa pretende alcançar de forma ampla."
        )

        objetivos_especificos_str = "\n".join(st.session_state.objetivos_especificos)
        edited_objetivos_especificos_str = st.text_area(
            "Objetivos Específicos (um por linha):",
            value=objetivos_especificos_str,
            height=150,
            help="Passos concretos e mensuráveis para atingir o objectivo geral. Cada objectivo deve começar com um verbo no infinitivo."
        )
        st.session_state.objetivos_especificos = [line.strip() for line in edited_objetivos_especificos_str.split('\n') if line.strip()]

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo5_prev"):
                st.session_state.passo = 3
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo5_next"):
                if st.session_state.objetivo_geral and st.session_state.objetivos_especificos:
                    st.session_state.passo = 5
                    st.rerun()
                else:
                    st.warning("Por favor, preencha o Objectivo Geral e pelo menos um Objectivo Específico.")

    elif st.session_state.passo == 6:
        st.subheader("6. 💡 Justificativa")
        st.write("Explique a importância e a relevância do seu pré-projeto nos âmbitos pessoal, acadêmico e social.")

        st.info(f"Tema: **{st.session_state.tema}**")
        st.info(f"Problema: **{st.session_state.problema_pesquisa}**")
        st.info(f"Objectivo Geral: **{st.session_state.objetivo_geral}**")
        st.info(f"Objectivos Específicos: **{', '.join(st.session_state.objetivos_especificos)}**")

        if st.button("✨ Gerar Justificativa Completa com IA", key="gerar_justificativa_button"):
            if (st.session_state.tema and st.session_state.problema_pesquisa and
                st.session_state.objetivo_geral and st.session_state.objetivos_especificos and
                st.session_state.curso and st.session_state.local_foco):
                
                with st.spinner("Gerando justificativa..."):
                    p, a, s = generate_justificativa_suggestions_gemini(
                        st.session_state.tema,
                        st.session_state.problema_pesquisa,
                        st.session_state.objetivo_geral,
                        st.session_state.objetivos_especificos,
                        st.session_state.curso,
                        st.session_state.local_foco
                    )
                    st.session_state.justificativa_pessoal = p
                    st.session_state.justificativa_academica = a
                    st.session_state.justificativa_social = s
                st.success("Justificativa gerada com sucesso!")
            else:
                st.warning("Por favor, preencha todos os campos anteriores (Tema, Problema, Objetivos, Curso, Local de Foco) para gerar a justificativa.")

        st.markdown("---")
        st.subheader("Edite as Justificativas:")
        st.session_state.justificativa_pessoal = st.text_area(
            "Justificativa Pessoal:",
            value=st.session_state.justificativa_pessoal,
            height=150,
            help="Descreva sua motivação pessoal para realizar esta pesquisa."
        )
        st.session_state.justificativa_academica = st.text_area(
            "Justificativa Acadêmica:",
            value=st.session_state.justificativa_academica,
            height=150,
            help="Explique a contribuição do seu trabalho para o avanço do conhecimento na área."
        )
        st.session_state.justificativa_social = st.text_area(
            "Justificativa Social:",
            value=st.session_state.justificativa_social,
            height=150,
            help="Discorra sobre o impacto e a relevância social dos resultados da sua pesquisa."
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo6_prev"):
                st.session_state.passo = 5
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo6_next"):
                st.session_state.passo = 7
                st.rerun()

    # elif st.session_state.passo == 7:
    #     st.subheader("7. 📚 Referencial Teórico")
    #     st.write("Defina os principais autores, teorias e conceitos que fundamentarão sua pesquisa.")

    #     st.info(f"Tema: **{st.session_state.tema}**")
    #     st.info(f"Problema: **{st.session_state.problema_pesquisa}**")

    #     if st.button("✨ Sugerir Tópicos e Autores com IA", key="gerar_referencial_button"):
    #         if st.session_state.tema and st.session_state.problema_pesquisa and st.session_state.curso and st.session_state.local_foco:
    #             st.session_state.referencial_teorico_sugestoes = generate_referencial_teorico_suggestions_gemini(
    #                 st.session_state.tema,
    #                 st.session_state.problema_pesquisa,
    #                 st.session_state.curso,
    #                 st.session_state.local_foco,
    #                 st.session_state.objetivo_geral if st.session_state.objetivo_geral else "Um objetivo geral provisório.",
    #                 st.session_state.objetivos_especificos if st.session_state.objetivos_especificos else ["Um objetivo específico provisório"]
    #             )
    #             st.success("Sugestões de referencial teórico geradas!")
    #         else:
    #             st.warning("Por favor, preencha Tema, Problema de Pesquisa, Curso e Local de Foco.")
    #     st.markdown("---")
    # st.subheader("🔍 Buscar Referências Automaticamente")
    elif st.session_state.passo == 7:  # Passo do Referencial Teórico
        st.subheader("7. 📚 Referencial Teórico")
    
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✨ Gerar Referencial Completo", key="gerar_referencial_completo"):
                with st.spinner("Buscando referências e gerando sugestões..."):
                    st.session_state.referencial_teorico_sugestoes = generate_referencial_teorico_suggestions_gemini(
                        tema=st.session_state.tema,
                        problema=st.session_state.problema_pesquisa,
                        curso=st.session_state.curso,
                        local_foco=st.session_state.local_foco,
                        objetivo_geral=st.session_state.objetivo_geral,
                        objetivos_especificos=st.session_state.objetivos_especificos
                    )
        
        with col2:
            if st.button("🔍 Buscar apenas no Google Scholar", disabled=not SERPAPI_KEY):
                st.session_state.referencial_teorico_sugestoes = buscar_referencias_serpapi(st.session_state.tema)
        
        # Edição do referencial
        if st.session_state.referencial_teorico_sugestoes:
            edited_ref = st.text_area(
                "Edite seu referencial teórico:",
                value="\n\n".join(st.session_state.referencial_teorico_sugestoes_gemini),
                height=400
            )
            st.session_state.referencial_teorico_sugestoes = [r.strip() for r in edited_ref.split("\n\n") if r.strip()]
        
        if SERPAPI_KEY:
            if st.button("🌐 Buscar no Google Scholar", key="btn_buscar_scholar"):
                with st.spinner("Consultando Google Scholar..."):
                    referencias = buscar_referencias_serpapi(st.session_state.tema)
                    
                    if referencias:
                        st.success(f"Encontradas {len(referencias)} referências!")
                        for i, ref in enumerate(referencias):
                            with st.expander(f"{ref['titulo']} ({ref['ano']})"):
                                st.write(f"**Autores:** {ref['autores']}")
                                st.markdown(f"**Link:** [Acessar artigo]({ref['link']})")
                                
                                if st.button(f"Adicionar ao Referencial", key=f"add_ref_{i}"):
                                    citacao = f"{ref['autores']} ({ref['ano']}). {ref['titulo']}."
                                    st.session_state.referencial_teorico_sugestoes.append(citacao)
                                    st.rerun()
                    else:
                        st.warning("Nenhuma referência encontrada.")
            else:
                st.warning("Funcionalidade desativada (chave SerpAPI não configurada).")

        st.markdown("---")
        st.subheader("Edite seu Referencial Teórico:")

        if st.session_state.referencial_teorico_sugestoes:
            referencial_str = "\n".join(st.session_state.referencial_teorico_sugestoes)
            edited_referencial_str = st.text_area(
                "Tópicos do Referencial Teórico (um por linha):",
                value=referencial_str,
                height=300,
                help="Liste os principais tópicos, autores ou teorias que você irá abordar em sua revisão bibliográfica. "
            )
            st.session_state.referencial_teorico_sugestoes = [line.strip() for line in edited_referencial_str.split('\n') if line.strip()]

            col1_nav, col2_nav = st.columns(2)
            with col1_nav:
                if st.button("⬅️ Anterior", key="passo7_prev"):
                    st.session_state.passo = 6
                    st.rerun()
            with col2_nav:
                if st.button("Próximo ➡️", key="passo7_next"):
                    st.session_state.autores_citados = extract_authors_from_referencial(st.session_state.referencial_teorico_sugestoes)
                    st.session_state.passo = 8
                    st.rerun()

    elif st.session_state.passo == 8:
        st.subheader("8. 🔬 Metodologia")
        st.write("Descreva como sua pesquisa será realizada, incluindo tipo, abordagem, procedimentos e instrumentos.")

        st.info(f"Objetivo Geral: **{st.session_state.objetivo_geral}**")
        st.info(f"Objetivos Específicos: **{', '.join(st.session_state.objetivos_especificos)}**")
        st.info(f"Problema: **{st.session_state.problema_pesquisa}**")

        if st.button("✨ Gerar Metodologia Completa com IA", key="gerar_metodologia_button"):
            if (st.session_state.objetivo_geral and st.session_state.objetivos_especificos and
                st.session_state.tema and st.session_state.problema_pesquisa and
                st.session_state.curso and st.session_state.local_foco):
                
                with st.spinner("Gerando metodologia..."):
                    st.session_state.metodologia_sugestoes = generate_metodologia_suggestions_gemini(
                        st.session_state.objetivo_geral,
                        st.session_state.objetivos_especificos,
                        st.session_state.tema,
                        st.session_state.problema_pesquisa,
                        st.session_state.curso,
                        st.session_state.local_foco
                    )
                st.success("Metodologia gerada com sucesso!")
            else:
                st.warning("Por favor, preencha todos os campos anteriores (Objetivos, Tema, Problema, Curso, Local de Foco) para gerar a metodologia.")

        st.markdown("---")
        st.subheader("Edite sua Metodologia:")

        met_sug = st.session_state.metodologia_sugestoes

        st.session_state.metodologia_sugestoes['natureza'] = st.text_area("Natureza da Pesquisa:", value=met_sug['natureza'], height=70)
        st.session_state.metodologia_sugestoes['natureza_justificativa'] = st.text_area("Justificativa da Natureza:", value=met_sug['natureza_justificativa'], height=100)
        st.session_state.metodologia_sugestoes['abordagem'] = st.text_area("Abordagem da Pesquisa:", value=met_sug['abordagem'], height=70)
        st.session_state.metodologia_sugestoes['abordagem_justificativa'] = st.text_area("Justificativa da Abordagem:", value=met_sug['abordagem_justificativa'], height=100)
        st.session_state.metodologia_sugestoes['objetivos_pesquisa'] = st.text_area("Objetivos da Pesquisa (Tipo):", value=met_sug['objetivos_pesquisa'], height=70)
        st.session_state.metodologia_sugestoes['fundamentacao_teorica'] = st.text_area("Fundamentação Teórica da Metodologia:", value=met_sug['fundamentacao_teorica'], height=150)
        st.session_state.metodologia_sugestoes['visao_pesquisador'] = st.text_area("Visão do Pesquisador (Sua Perspectiva Metodológica):", value=met_sug['visao_pesquisador'], height=100)
        
        procedimentos_str = "\n".join(met_sug.get('procedimentos_tecnicos', []))
        edited_procedimentos_str = st.text_area("Procedimentos Técnicos (um por linha):", value=procedimentos_str, height=150)
        st.session_state.metodologia_sugestoes['procedimentos_tecnicos'] = [line.strip() for line in edited_procedimentos_str.split('\n') if line.strip()]

        st.session_state.metodologia_sugestoes['universo_amostra'] = st.text_area("Universo e Amostra:", value=met_sug['universo_amostra'], height=150, help="Descreva a população, amostra e tipo de amostragem.")
        
        instrumentos_str = "\n".join(met_sug.get('instrumentos_coleta', []))
        edited_instrumentos_str = st.text_area("Instrumentos de Coleta de Dados (um por linha):", value=instrumentos_str, height=150)
        st.session_state.metodologia_sugestoes['instrumentos_coleta'] = [line.strip() for line in edited_instrumentos_str.split('\n') if line.strip()]

        analise_str = "\n".join(met_sug.get('analise_de_dados', []))
        edited_analise_str = st.text_area("Análise de Dados (um por linha):", value=analise_str, height=150)
        st.session_state.metodologia_sugestoes['analise_de_dados'] = [line.strip() for line in edited_analise_str.split('\n') if line.strip()]

        st.session_state.metodologia_sugestoes['consideracoes_eticas'] = st.text_area("Considerações Éticas:", value=met_sug['consideracoes_eticas'], height=100)

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo8_prev"):
                st.session_state.passo = 7
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo8_next"):
                st.session_state.passo = 9
                st.rerun()

    elif st.session_state.passo == 9:
        st.subheader("9. 🗓️ Cronograma e Resultados Esperados")
        st.write("Organize as atividades da sua pesquisa ao longo do tempo e defina o que você espera alcançar.")

        st.info(f"Tema: **{st.session_state.tema}**")
        st.info(f"Objetivos Específicos: **{', '.join(st.session_state.objetivos_especificos)}**")

        if st.button("✨ Gerar Cronograma e Resultados com IA", key="gerar_cronograma_button"):
            if st.session_state.tema and st.session_state.objetivos_especificos and st.session_state.curso:
                with st.spinner("Gerando cronograma e resultados esperados..."):
                    cronograma_sug, resultados_sug = generate_cronograma_results_gemini(
                        st.session_state.tema,
                        st.session_state.objetivos_especificos,
                        st.session_state.curso
                    )
                    st.session_state.cronograma_results_sugestoes['cronograma'] = cronograma_sug
                    st.session_state.cronograma_results_sugestoes['resultados_esperados'] = resultados_sug

                    st.session_state.meses_cronograma = sorted(cronograma_sug.keys(), key=lambda x: int(x.split(' ')[1]))
                st.success("Cronograma e resultados gerados com sucesso!")
            else:
                st.warning("Por favor, preencha Tema, Objetivos Específicos e Curso para gerar o cronograma.")

        st.markdown("---")
        st.subheader("Edite seu Cronograma (6 Meses):")

        meses_cronograma = st.session_state.meses_cronograma if st.session_state.meses_cronograma else [f"Mês {i+1}" for i in range(6)]
        
        num_cols_cronograma = len(meses_cronograma)
        if num_cols_cronograma > 0:
            cols = st.columns(num_cols_cronograma)
            for i, mes in enumerate(meses_cronograma):
                with cols[i]:
                    current_value = st.session_state.cronograma_results_sugestoes['cronograma'].get(mes, "")
                    st.session_state.cronograma_results_sugestoes['cronograma'][mes] = st.text_area(
                        f"Atividades do {mes}:",
                        value=current_value,
                        height=150,
                        key=f"cronograma_mes_{mes}"
                    )
        else:
            st.info("Gere o cronograma para visualizar e editar as atividades por mês.")

        st.markdown("---")
        st.subheader("Edite os Resultados Esperados:")
        resultados_esperados_str = "\n".join(st.session_state.cronograma_results_sugestoes.get('resultados_esperados', []))
        edited_resultados_esperados_str = st.text_area(
            "Resultados Esperados (um por linha):",
            value=resultados_esperados_str,
            height=150,
            help="O que você espera alcançar ou produzir com a sua pesquisa (Ex: Artigo científico, protótipo, relatório)."
        )
        st.session_state.cronograma_results_sugestoes['resultados_esperados'] = [line.strip() for line in edited_resultados_esperados_str.split('\n') if line.strip()]

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo9_prev"):
                st.session_state.passo = 8
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo9_next"):
                st.session_state.passo = 10
                st.rerun()

    elif st.session_state.passo == 10:
        st.subheader("10. 📝 Introdução")
        st.write("Revise e edite a introdução do seu pré-projeto, que apresenta uma visão geral de todo o trabalho.")

        st.info(f"Tema: **{st.session_state.tema}**")
        st.info(f"Problema: **{st.session_state.problema_pesquisa}**")
        st.info(f"Objetivo Geral: **{st.session_state.objetivo_geral}**")
        
        metodologia_resumo = f"""
        - Natureza: {st.session_state.metodologia_sugestoes.get('natureza', '')}
        - Abordagem: {st.session_state.metodologia_sugestoes.get('abordagem', '')}
        - Procedimentos: {', '.join(st.session_state.metodologia_sugestoes.get('procedimentos_tecnicos', []))}
        """
        
        if st.button("✨ Gerar Introdução com IA", key="gerar_introducao_button"):
            if (st.session_state.tema and st.session_state.problema_pesquisa and 
                st.session_state.objetivo_geral and st.session_state.objetivos_especificos):
                
                with st.spinner("Gerando introdução..."):
                    st.session_state.introducao = generate_introducao_gemini(
                        st.session_state.tema,
                        st.session_state.problema_pesquisa,
                        st.session_state.problematizacao,
                        st.session_state.objetivo_geral,
                        st.session_state.objetivos_especificos,
                        metodologia_resumo
                    )
                st.success("Introdução gerada com sucesso!")
            else:
                st.warning("Por favor, preencha pelo menos Tema, Problema, Objetivo Geral e Objetivos Específicos para gerar a introdução.")

        st.markdown("---")
        st.subheader("Edite sua Introdução:")
        st.session_state.introducao = st.text_area(
            "Introdução do Pré-Projeto:",
            value=st.session_state.introducao,
            height=400,
            help="Esta seção deve apresentar uma visão geral do seu trabalho, incluindo tema, problema, objetivos e metodologia."
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo10_prev"):
                st.session_state.passo = 9
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo10_next"):
                st.session_state.passo = 11
                st.rerun()

    elif st.session_state.passo == 11:
        st.subheader("11. 🏁 Conclusão")
        st.write("Revise e edite a conclusão do seu pré-projeto antes de gerar o documento final.")

        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**Tema:** {st.session_state.tema}")
            st.info(f"**Problema:** {st.session_state.problema_pesquisa}")
        with col2:
            st.info(f"**Objetivo Geral:** {st.session_state.objetivo_geral}")
            st.info(f"**Resultados Esperados:** {', '.join(st.session_state.cronograma_results_sugestoes.get('resultados_esperados', []))}")

        with st.expander("🔍 Informações Adicionais para Conclusão"):
            contribuicoes = st.text_area(
                "Descreva as contribuições esperadas do seu trabalho:",
                value="Este estudo pretende contribuir para [...] tanto no âmbito acadêmico quanto [...]",
                height=100,
                help="Descreva como seu trabalho pode avançar o conhecimento na área ou resolver problemas práticos."
            )
            
            limitacoes = st.text_area(
                "Descreva limitações ou desafios previstos:",
                value="Algumas limitações podem incluir [...]",
                height=100,
                help="Mencione quaisquer restrições metodológicas ou de escopo."
            )
            
            pesquisas_futuras = st.text_area(
                "Sugestões para pesquisas futuras:",
                value="Estudos futuros poderiam investigar [...]",
                height=100,
                help="Indique direções para pesquisas subsequentes."
            )

        if st.button("✨ Gerar Conclusão com IA", key="gerar_conclusao_button"):
            if (st.session_state.tema and st.session_state.problema_pesquisa and 
                st.session_state.objetivo_geral and st.session_state.objetivos_especificos):
                
                with st.spinner("Gerando conclusão acadêmica..."):
                    contexto_conclusao = {
                        "contribuicoes": contribuicoes,
                        "limitacoes": limitacoes,
                        "pesquisas_futuras": pesquisas_futuras
                    }
                    
                    st.session_state.conclusao = generate_conclusao_gemini(
                        st.session_state.tema,
                        st.session_state.problema_pesquisa,
                        st.session_state.objetivo_geral,
                        st.session_state.objetivos_especificos,
                        st.session_state.cronograma_results_sugestoes.get('resultados_esperados', []),
                        contexto_conclusao
                    )
                st.success("Conclusão gerada com sucesso!")
            else:
                st.warning("Por favor, complete pelo menos os passos de Tema, Problema e Objetivos para gerar a conclusão.")

        st.markdown("---")
        st.subheader("Edite sua Conclusão:")
        st.session_state.conclusao = st.text_area(
            "Texto da Conclusão:",
            value=st.session_state.conclusao,
            height=400,
            help="Revise cuidadosamente a conclusão gerada. Ela deve sintetizar o trabalho e destacar suas contribuições."
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo11_prev"):
                st.session_state.passo = 10
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo11_next"):
                if st.session_state.conclusao:
                    st.session_state.passo = 12
                    st.rerun()
                else:
                    st.warning("Por favor, gere ou insira uma conclusão antes de prosseguir.")

    elif st.session_state.passo == 12:
        st.subheader("12. 📖 Referências Bibliográficas")
        st.write("Gere as referências completas com base em todas as citações do seu pré-projeto.")

        # Consolida o texto completo do pré-projeto para análise de referências
        full_text = f"""
        PROBLEMA: {st.session_state.problema_pesquisa}
        PROBLEMATIZAÇÃO: {st.session_state.problematizacao}
        TIPO DE INQUÉRITO: {st.session_state.tipo_inquerito_pesquisa}
        PERGUNTAS DE PESQUISA: {'; '.join(st.session_state.perguntas_pesquisa)}
        HIPÓTESES: {'; '.join(st.session_state.hipoteses)}
        OBJETIVO GERAL: {st.session_state.objetivo_geral}
        OBJETIVOS ESPECÍFICOS: {'; '.join(st.session_state.objetivos_especificos)}
        JUSTIFICATIVA PESSOAL: {st.session_state.justificativa_pessoal}
        JUSTIFICATIVA ACADÊMICA: {st.session_state.justificativa_academica}
        JUSTIFICATIVA SOCIAL: {st.session_state.justificativa_social}
        REFERENCIAL TEÓRICO: {'; '.join(st.session_state.referencial_teorico_sugestoes)}
        METODOLOGIA: {str(st.session_state.metodologia_sugestoes)}
        RESULTADOS ESPERADOS: {'; '.join(st.session_state.cronograma_results_sugestoes.get('resultados_esperados', []))}
        INTRODUÇÃO: {st.session_state.introducao}
        CONCLUSÃO: {st.session_state.conclusao}
        INSTRUMENTO GERADO: {st.session_state.instrumento_gerado}
        TIPO DE INSTRUMENTO: {st.session_state.tipo_instrumento_display}
        NORMA BIBLIOGRÁFICA: {st.session_state.norma_bibliografica}
        """
        
        st.session_state.norma_bibliografica = st.selectbox(
            "Selecione a Norma Bibliográfica:",
            [ "APA", "ABNT", "Vancouver", "IEEE"],
            index=[ "APA", "ABNT", "Vancouver", "IEEE"].index(st.session_state.norma_bibliografica),
            key="norma_bibliografica_select"
        )

        if st.button("✨ Gerar Referências com IA", key="gerar_referencias_button"):
            with st.spinner("Analisando o documento e gerando referências..."):
                st.session_state.referencias_bibliograficas = generate_references_gemini(
                    full_text,
                    st.session_state.norma_bibliografica,
                    st.session_state.tema,
                    st.session_state.curso,
                )
            st.success("Referências geradas com sucesso!")

        st.markdown("---")
        st.subheader("Referências Geradas:")
        st.session_state.referencias_bibliograficas = st.text_area(
            "Edite as referências conforme necessário:",
            value=st.session_state.referencias_bibliograficas,
            height=400,
            help="Revise e complete as referências conforme necessário. As referências foram geradas com base em todas as citações detectadas no seu pré-projeto."
        )

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo12_prev"):
                st.session_state.passo = 11
                st.rerun()
        with col2_nav:
            if st.button("Próximo ➡️", key="passo12_next"):
                st.session_state.passo = 13
                st.rerun()

    
#--- Step 13: Data Collection Instrument Generation ---
    elif st.session_state.passo == 13:
        st.subheader("13. 📊 Geração de Instrumento de Coleta de Dados")
        st.markdown("Crie questionários, roteiros de entrevista ou outros instrumentos para sua pesquisa.")

        # Verificação de campos obrigatórios
        required_fields = {
            'tema': "Tema da pesquisa",
            'problema_pesquisa': "Problema de pesquisa",
            'objetivo_geral': "Objetivo geral",
            'objetivos_especificos': "Objetivos específicos"
        }

        missing_fields = [name for field, name in required_fields.items() 
                        if not st.session_state.get(field)]
        
        if missing_fields:
            st.error(f"⚠️ Complete primeiro: {', '.join(missing_fields)}")
            st.button("↩️ Voltar para corrigir", on_click=lambda: setattr(st.session_state, 'passo', 4))
            st.stop()

        # Seleção do tipo de instrumento
        instrumentos_disponiveis = st.session_state.metodologia_sugestoes.get('instrumentos_coleta', [])
        if not instrumentos_disponiveis:
            instrumentos_disponiveis = ["Questionário", "Roteiro de Entrevista", "Formulário de Observação"]
            st.info("📌 Sugestão automática de instrumentos baseada em metodologias comuns")

        tipo_instrumento = st.selectbox(
            "Selecione o tipo de instrumento:",
            options=instrumentos_disponiveis,
            index=0,
            key="select_tipo_instrumento"
        )

        # Seção de geração
        col1, col2 = st.columns([3, 1])
        with col1:
            if st.button("🔄 Gerar Instrumento", type="primary", use_container_width=True):
                with st.spinner(f"Gerando {tipo_instrumento.lower()}..."):
                    try:
                        generated_content = generate_data_collection_instrument_gemini(
                            tema=st.session_state.tema,
                            problema=st.session_state.problema_pesquisa,
                            objetivo_geral=st.session_state.objetivo_geral,
                            objetivos_especificos=st.session_state.objetivos_especificos,
                            tipo_instrumento=tipo_instrumento,
                            curso=st.session_state.get('curso', ''),
                            local_foco=st.session_state.get('local_foco', '')
                        )
                        
                        # Atualiza o estado
                        if 'todos_instrumentos_gerados' not in st.session_state:
                            st.session_state.todos_instrumentos_gerados = {}
                        
                        st.session_state.todos_instrumentos_gerados[tipo_instrumento] = generated_content
                        st.session_state.instrumento_counter = len(st.session_state.todos_instrumentos_gerados)
                        st.toast(f"{tipo_instrumento} gerado com sucesso!", icon="✅")
                        
                    except Exception as e:
                        st.error(f"Falha na geração: {str(e)}")
                        st.stop()

        with col2:
            if st.button("🧹 Limpar", use_container_width=True, 
                    disabled=tipo_instrumento not in st.session_state.get('todos_instrumentos_gerados', {})):
                del st.session_state.todos_instrumentos_gerados[tipo_instrumento]
                st.rerun()

        # Exibição do instrumento gerado
        if tipo_instrumento in st.session_state.get('todos_instrumentos_gerados', {}):
            st.markdown("---")
            st.subheader(f"📝 {tipo_instrumento} Gerado")
            
            with st.expander("🔍 Visualizar Instrumento", expanded=True):
                st.markdown(st.session_state.todos_instrumentos_gerados[tipo_instrumento])

            # Download
            doc_buffer = create_instrument_word_document(
                title=f"Instrumento_{tipo_instrumento.replace(' ', '_')}",
                instrument_text=st.session_state.todos_instrumentos_gerados[tipo_instrumento]
            )
            
            st.download_button(
                label="📥 Baixar Documento (.docx)",
                data=doc_buffer,
                file_name=f"{tipo_instrumento.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.info("ℹ️ Gere o instrumento clicando no botão acima")

        # Navegação
        st.markdown("---")
        nav_col1, nav_col2, nav_col3 = st.columns([1, 1, 2])
        with nav_col1:
            if st.button("⬅️ Voltar", key="btn_prev_step13"):
                st.session_state.passo = 12
                st.rerun()
        
        with nav_col2:
            if st.button("➡️ Próximo", key="btn_next_step13", 
                        disabled=not st.session_state.get('todos_instrumentos_gerados')):
                st.session_state.passo = 14
                st.rerun()
        
        with nav_col3:
            if st.button("💾 Salvar Rascunho", key="btn_save_draft"):
                save_current_project()
    # elif st.session_state.passo == 13:
    #     st.subheader("13. 📊 Geração de Instrumento de Colecta de Dados")
    #     st.markdown("Crie um Instrumento de Colecta de Dados com base nos detalhes do seu pré-projecto.")

    #     # Inicializa a variável
    #     tipo_instrumento_selecionado = None
        
    #     # Verification of required fields
    #     required_fields = ['tema', 'problema_pesquisa', 'objetivo_geral', 'objetivos_especificos']
    #     if not all(st.session_state.get(field) for field in required_fields):
    #         st.warning("Por favor, preencha as secções de Tema, Problema e Objectivos antes de continuar.")
    #     else:
    #         # Get suggested instruments or use default
    #         instrumentos_sugeridos = st.session_state.metodologia_sugestoes.get('instrumentos_coleta', [])
    #         if not instrumentos_sugeridos:
    #             instrumentos_sugeridos = ["Roteiro de Entrevista", "Questionário"]
    #             st.info("Usando opções padrão de instrumentos de coleta.")

    #         # Selection of instrument type
    #         tipo_instrumento_selecionado = st.selectbox(
    #             "Selecione o tipo de instrumento para gerar ou visualizar:",
    #             instrumentos_sugeridos,
    #             key="select_tipo_instrumento"
    #         )

    #     # --- Botão para gerar instrumento ---
    #     if tipo_instrumento_selecionado and st.button("Gerar Instrumento", key="gerar_instrumento_button"):
    #         if (st.session_state.tema and st.session_state.problema_pesquisa and 
    #             st.session_state.objetivo_geral and st.session_state.objetivos_especificos):

    #             with st.spinner(f"Gerando {tipo_instrumento_selecionado.lower()}..."):
    #                 # Garante que o dicionário existe
    #                 if 'todos_instrumentos_gerados' not in st.session_state:
    #                     st.session_state.todos_instrumentos_gerados = {}
    #                 # Gera o conteúdo
    #                 generated_content = generate_data_collection_instrument_gemini(
    #                     tema=st.session_state.tema,
    #                     problema=st.session_state.problema_pesquisa,
    #                     objetivo_geral=st.session_state.objetivo_geral,
    #                     objetivos_especificos=st.session_state.objetivos_especificos,
    #                     tipo_instrumento=tipo_instrumento_selecionado,
    #                     curso=st.session_state.get('curso', 'não especificado'),
    #                     local_foco=st.session_state.get('local_foco', 'não especificado')
    #                 )

    #                 # Armazena no session state
    #                 st.session_state.todos_instrumentos_gerados[tipo_instrumento_selecionado] = generated_content
    #                 st.session_state.current_instrument_type = tipo_instrumento_selecionado
    #                 st.session_state.instrumento_gerado = generated_content

    #             st.success(f"{tipo_instrumento_selecionado} gerado com sucesso!")
    #         else:
    #             st.warning("Por favor, complete pelo menos os passos de Tema, Problema e Objetivos para gerar o instrumento.")

    #     # --- Visualização do instrumento gerado ---
    #     if tipo_instrumento_selecionado and tipo_instrumento_selecionado in st.session_state.get('todos_instrumentos_gerados', {}):
    #         st.markdown("---")
    #         st.subheader(f"Visualização do {tipo_instrumento_selecionado} Gerado:")
    #         st.markdown(st.session_state.todos_instrumentos_gerados[tipo_instrumento_selecionado])

    #         # Botão para resetar/remover o instrumento gerado
    #         if st.button("🔄 Gerar Novo Instrumento", key="reset_instrumento_button"):
    #             # Remove o instrumento atual para permitir nova geração
    #             if tipo_instrumento_selecionado in st.session_state.todos_instrumentos_gerados:
    #                 del st.session_state.todos_instrumentos_gerados[tipo_instrumento_selecionado]
    #             st.session_state.instrumento_gerado = ""
    #             st.session_state.current_instrument_type = ""
    #             st.success("Instrumento removido. Pronto para gerar um novo.")
    #             st.rerun()

    #         # Download do instrumento gerado
    #         doc_buffer = create_instrument_word_document(
    #             title=tipo_instrumento_selecionado,
    #             instrument_text=st.session_state.todos_instrumentos_gerados[tipo_instrumento_selecionado]
    #         )
    #         file_name = f"{tipo_instrumento_selecionado.replace(' ', '_')}.docx"
    #         file_name = re.sub(r'[\\/*?:"<>|]', "", file_name)
    #         st.download_button(
    #             label=f"⬇️ Baixar {tipo_instrumento_selecionado}",
    #             data=doc_buffer,
    #             file_name=file_name,
    #             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    #             key=f"download_{tipo_instrumento_selecionado}"
    #         )
    #     elif tipo_instrumento_selecionado:
    #         st.info("Nenhum instrumento gerado ainda. Clique em 'Gerar Instrumento' para criar um.")

    #     col1, col2 = st.columns(2)
    #     with col1:
    #             if st.button("⬅️ Voltar", key="btn_voltar_passo12"):
    #                 st.session_state.passo = 12
    #                 st.rerun()
    #     with col2:
    #             if st.button("➡️ Próximo", key="btn_avancar_passo12"):
    #                 st.session_state.passo = 14
    #                 st.rerun()



    elif st.session_state.passo == 14:
        st.subheader("14. 🔍 Revisão Final")
        st.write("Revise todas as informações do seu pré-projeto antes de gerar o documento final.")

        with st.expander("📋 Resumo do Pré-Projeto"):
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Informações Básicas")
                st.write(f"**Tema:** {st.session_state.tema}")
                st.write(f"**Curso:** {st.session_state.curso}")
                st.write(f"**Local de Foco:** {st.session_state.local_foco}")
                
            with col2:
                st.subheader("Elementos Principais")
                st.write(f"**Problema:** {st.session_state.problema_pesquisa[:100]}...")
                st.write(f"**Objetivo Geral:** {st.session_state.objetivo_geral[:100]}...")
                st.write(f"**Metodologia:** {st.session_state.metodologia_sugestoes.get('natureza', '')}")

        col1_nav, col2_nav = st.columns(2)
        with col1_nav:
            if st.button("⬅️ Anterior", key="passo14_prev"):
                st.session_state.passo = 13
                st.rerun()
        with col2_nav:
            if st.button("Gerar Documento Final ➡️", key="passo14_next"):
                st.session_state.passo = 15
                st.rerun()

    elif st.session_state.passo == 15:
        st.subheader("15. ✅ Download do Documento Final")
        st.write("Seu pré-projeto está completo! Revise os detalhes abaixo e gere o documento final.")

        with st.expander("🔍 Visualizar Resumo do Pré-Projeto"):
            st.json({
                "Tema": st.session_state.tema,
                "Problema": st.session_state.problema_pesquisa,
                "Objetivos": {
                    "Geral": st.session_state.objetivo_geral,
                    "Específicos": st.session_state.objetivos_especificos
                },
                "Metodologia": list(st.session_state.metodologia_sugestoes.keys()),
                "Principais Referências": st.session_state.autores_citados[:5] if st.session_state.autores_citados else "Nenhuma"
            })

        data_for_docx = get_project_data()
        
        try:
            word_file = create_word_document(data_for_docx)
            file_name = f"Pre_Projeto_{st.session_state.tema[:50]}.docx"
            file_name = re.sub(r'[\\/*?:"<>|]', "", file_name)
            
            if word_file is None:
                st.error("Falha ao gerar o documento. Por favor, verifique os dados do projeto.")
                st.stop()

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button(" Voltar ", key="back_to_review"):
                    st.session_state.passo = 14
                    st.rerun()
          
            with col2:
                st.download_button(
                    label="⬇️ Baixar Pré-Projeto (.docx)",
                    data=word_file,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_docx"
                )
            
            with col3:
                if st.button("💾 Salvar Projeto", key="save_project"):
                    save_current_project()
                    st.rerun()
            with col4:
                if st.button("🔄 Começar Novo Projeto", key="passo15_new"):
                    for key in list(st.session_state.keys()):
                        if key not in ['user_projects']:
                            del st.session_state[key]
                            initialize_session_state()
                    st.rerun()
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {str(e)}")
            st.stop()

        # Substitua a seção de projetos salvos (dentro do passo 15) por este código corrigido:

    st.markdown("---")
    st.subheader("📂 Meus Projetos Salvos")

    if 'user_projects' not in st.session_state:
        st.session_state['user_projects'] = db_manager.get_user_projects(user_id)

    if st.session_state['user_projects']:
        for project in st.session_state['user_projects']:
            # Verifica se o projeto tem a estrutura esperada
            project_data = project.get('data', {}) if isinstance(project, dict) else {}
            project_name = project.get('name', 'Projeto sem nome')
            created_at = project.get('created_at', '')[:10] if project.get('created_at') else 'Data desconhecida'
            updated_at = project.get('updated_at', '')[:16] if project.get('updated_at') else 'Data desconhecida'
            
            with st.expander(f"📄 {project_name} - {created_at}"):
                col1, col2, col3 = st.columns([3,1,1])
                with col1:
                    st.write(f"**Tema:** {project_data.get('tema', 'Não definido')}")
                    st.write(f"**Última atualização:** {updated_at}")
                
                with col2:
                    if st.button("Carregar", key=f"load_{project.get('id', '')}"):
                        loaded_project = db_manager.load_project(project.get('id'))
                        if loaded_project:
                            # Verifica se o projeto carregado tem a estrutura esperada
                            if isinstance(loaded_project, dict):
                                if 'data' in loaded_project:
                                    load_project_data(loaded_project['data'])
                                    st.session_state['current_project_id'] = project.get('id')
                                    st.success(f"Projeto '{project_name}' carregado com sucesso!")
                                    st.rerun()
                                else:
                                    st.error("Estrutura do projeto inválida - chave 'data' não encontrada")
                            else:
                                st.error("Tipo de projeto inválido ao carregar")
                
                with col3:
                    if st.button("Excluir", key=f"del_{project.get('id', '')}"):
                        if db_manager.delete_project(project.get('id')):
                            st.session_state['user_projects'] = db_manager.get_user_projects(user_id)
                            if 'current_project_id' in st.session_state and st.session_state['current_project_id'] == project.get('id'):
                                del st.session_state['current_project_id']
                            st.success("Projeto excluído com sucesso!")
                            st.rerun()
                    else:
                        st.info("Você ainda não tem projetos salvos.")
    
    if st.button("🧪 Avaliar Qualidade do Documentos"):
      avaliar_qualidade_documento(get_project_data())
# 
    if 'secao_a_regenerar' in st.session_state:
        nova = regenerar_secao(st.session_state['secao_a_regenerar'], get_project_data())
        st.success(f"{st.session_state['secao_a_regenerar']} regenerada com sucesso!")
    if 'secao_a_regenerar' in st.session_state:
        secao = st.session_state['secao_a_regenerar']
        projeto = get_project_data()
        nova_versao = regenerar_secao(secao, projeto)

        # Aplica o novo conteúdo à seção correta
        if secao == "Introdução":
            st.session_state['introducao'] = nova_versao
        elif secao == "Conclusão":
            st.session_state['conclusao'] = nova_versao
        elif secao == "Problema de Pesquisa":
            st.session_state['problema_pesquisa'] = nova_versao
        elif secao == "Problematização":
            st.session_state['problematizacao'] = nova_versao
        elif secao == "Hipóteses":
            st.session_state['hipoteses'] = nova_versao
        elif secao == "Perguntas de Pesquisa":
            st.session_state['perguntas_pesquisa'] = nova_versao
        elif secao == "Objetivo Geral":
            st.session_state['objetivo_geral'] = nova_versao[0]
            st.session_state['objetivos_especificos'] = nova_versao[1]
        elif secao.startswith("Justificativa"):
            pessoal, academica, social = nova_versao
            st.session_state['justificativa_pessoal'] = pessoal
            st.session_state['justificativa_academica'] = academica
            st.session_state['justificativa_social'] = social
        elif secao == "Referencial Teórico":
            st.session_state['referencial_teorico_sugestoes'] = nova_versao
        elif secao == "Metodologia":
            st.session_state['metodologia_sugestoes'] = nova_versao
        elif secao == "Resultados Esperados":
            cronograma, resultados = nova_versao
            st.session_state['cronograma_results_sugestoes']['cronograma'] = cronograma
            st.session_state['cronograma_results_sugestoes']['resultados_esperados'] = resultados
        elif secao == "Referências Bibliográficas":
            st.session_state['referencias_bibliograficas'] = nova_versao

        st.success(f"✅ Seção '{secao}' regenerada e atualizada com sucesso!")
        del st.session_state['secao_a_regenerar']


   # Rodapé
    st.markdown("""
    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p style="font-size: 0.8em;">© 2025 AcadêmicoPro - Todos os direitos reservados</p>
    </div>
    """, unsafe_allow_html=True)
if __name__ == "__main__":
    main()