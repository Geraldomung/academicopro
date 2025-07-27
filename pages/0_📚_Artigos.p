import subprocess
import streamlit as st
import sqlite3
from datetime import datetime
import requests
from docx import Document
import re
from collections import Counter
from utils.auth import get_auth_status
import os
from dotenv import load_dotenv
import google.generativeai as genai
import language_tool_python

load_dotenv()

# Configuração inicial da página
st.set_page_config(page_title="Editor de Artigo Científico", layout="wide")

# Configuração do Gemini
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

# Autenticação
auth_status, username, user_id, role, is_active = get_auth_status()
if not auth_status:
    st.warning("Por favor, faça login para acessar esta página.")
    st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Login")
    st.stop()




## 1. Funções de Banco de Dados SQLite
def init_db():
    """Inicializa o banco de dados SQLite"""
    conn = sqlite3.connect('artigos.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS artigos
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  titulo TEXT,
                  palavras_chave TEXT,
                  resumo TEXT,
                  introducao TEXT,
                  metodologia TEXT,
                  resultados TEXT,
                  conclusao TEXT,
                  referencias TEXT,
                  user_id INTEGER,
                  data_criacao TIMESTAMP,
                  data_atualizacao TIMESTAMP)''')
    conn.commit()
    conn.close()

def salvar_artigo():
    """Salva ou atualiza o artigo no banco de dados"""
    conn = sqlite3.connect('artigos.db')
    c = conn.cursor()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if st.session_state.artigo['id']:
        c.execute('''UPDATE artigos SET 
                    titulo=?, palavras_chave=?, resumo=?, introducao=?,
                    metodologia=?, resultados=?, conclusao=?, referencias=?,
                    data_atualizacao=?
                    WHERE id=? AND user_id=?''',
                 (st.session_state.artigo['titulo'],
                  st.session_state.artigo['palavras_chave'],
                  st.session_state.artigo['resumo'],
                  st.session_state.artigo['introducao'],
                  st.session_state.artigo['metodologia'],
                  st.session_state.artigo['resultados'],
                  st.session_state.artigo['conclusao'],
                  st.session_state.artigo['referencias'],
                  agora,
                  st.session_state.artigo['id'],
                  user_id))
    else:
        c.execute('''INSERT INTO artigos 
                    (titulo, palavras_chave, resumo, introducao, metodologia,
                     resultados, conclusao, referencias, user_id, data_criacao, data_atualizacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                 (st.session_state.artigo['titulo'],
                  st.session_state.artigo['palavras_chave'],
                  st.session_state.artigo['resumo'],
                  st.session_state.artigo['introducao'],
                  st.session_state.artigo['metodologia'],
                  st.session_state.artigo['resultados'],
                  st.session_state.artigo['conclusao'],
                  st.session_state.artigo['referencias'],
                  user_id,
                  agora, agora))
        st.session_state.artigo['id'] = c.lastrowid
    
    conn.commit()
    conn.close()
    st.success("Artigo salvo no banco de dados com sucesso!")

def carregar_artigos():
    """Retorna lista de artigos do banco de dados"""
    conn = sqlite3.connect('artigos.db')
    c = conn.cursor()
    c.execute("SELECT id, titulo, data_atualizacao FROM artigos WHERE user_id=? ORDER BY data_atualizacao DESC", (user_id,))
    artigos = c.fetchall()
    conn.close()
    return artigos

def carregar_artigo_por_id(artigo_id):
    """Carrega um artigo específico do banco de dados"""
    conn = sqlite3.connect('artigos.db')
    c = conn.cursor()
    c.execute("SELECT * FROM artigos WHERE id=? AND user_id=?", (artigo_id, user_id))
    artigo_db = c.fetchone()
    conn.close()
    return artigo_db

def excluir_artigo(artigo_id):
    """Remove um artigo do banco de dados"""
    conn = sqlite3.connect('artigos.db')
    c = conn.cursor()
    c.execute("DELETE FROM artigos WHERE id=? AND user_id=?", (artigo_id, user_id))
    conn.commit()
    conn.close()
    st.success("Artigo excluído com sucesso!")

## 2. Funções de Integração com APIs
def buscar_referencias_crossref(titulo):
    """Busca referências acadêmicas usando a API do Crossref"""
    url = "https://api.crossref.org/works"
    params = {"query.title": titulo, "rows": 5}
    try:
        response = requests.get(url, params=params)
        if response.status_code == 200:
            dados = response.json()
            referencias = []
            for item in dados["message"]["items"]:
                autores = ", ".join([autor.get("given", "") + " " + autor.get("family", "") 
                                   for autor in item.get("author", []) if autor.get("given") and autor.get("family")])
                ano = item.get("issued", {}).get("date-parts", [[None]])[0][0]
                titulo_ref = item.get("title", ["Sem título"])[0]
                revista = item.get("container-title", ["Sem revista"])[0]
                volume = item.get("volume", "")
                issue = item.get("issue", "")
                paginas = item.get("page", "")
                doi = item.get("DOI", "")
                
                referencia = f"{autores}. ({ano}). {titulo_ref}. {revista}, {volume}({issue}), {paginas}. DOI: {doi}"
                referencias.append(referencia)
            return "\n".join(referencias)
        return "Nenhum resultado encontrado."
    except Exception as e:
        return f"Erro na busca: {str(e)}"

## 3. Funções de Exportação
def exportar_para_word():
    """Exporta o artigo para um documento Word formatado"""
    try:
        doc = Document()
        
        # Título
        doc.add_heading(st.session_state.artigo['titulo'], level=1)
        
        # Palavras-chave
        doc.add_paragraph(f"Palavras-chave: {st.session_state.artigo['palavras_chave']}")
        
        # Seções
        secoes = [
            ("Resumo", st.session_state.artigo['resumo']),
            ("Introdução", st.session_state.artigo['introducao']),
            ("Metodologia", st.session_state.artigo['metodologia']),
            ("Resultados e Discussão", st.session_state.artigo['resultados']),
            ("Conclusão", st.session_state.artigo['conclusao']),
            ("Referências", st.session_state.artigo['referencias'])
        ]
        
        for titulo, conteudo in secoes:
            doc.add_heading(titulo, level=2)
            doc.add_paragraph(conteudo)
        
        doc.save("artigo_cientifico.docx")
        
        with open("artigo_cientifico.docx", "rb") as f:
            st.download_button(
                label="⬇️ Baixar Documento Word",
                data=f,
                file_name="artigo_cientifico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"Erro ao exportar para Word: {str(e)}")

# def verificar_coerencia_avancada(secao, novo_conteudo):
#     """Verificação avançada de coerência entre seções"""
#     context = get_global_context()
#     problemas = []
    
#     # 1. Verificar consistência terminológica
#     termos_chave = set()
#     if context['palavras_chave']:
#         termos_chave.update([kw.strip().lower() for kw in context['palavras_chave'].split(',')])
    
#     # Adicionar termos do título
#     if context['titulo']:
#         termos_chave.update([word.lower() for word in context['titulo'].split() if len(word) > 4])
    
#     # Verificar se os termos aparecem no novo conteúdo
#     for termo in termos_chave:
#         if termo not in novo_conteudo.lower():
#             problemas.append(f"Termo-chave '{termo}' não mencionado")

#     # 2. Verificar consistência metodológica
#     if secao == "resultados":
#         if context['metodologia']:
#             metodos = ["qualitativo", "quantitativo", "experimento", "entrevista", "survey"]
#             usado = [m for m in metodos if m in context['metodologia'].lower()]
#             for m in usado:
#                 if m not in novo_conteudo.lower():
#                     problemas.append(f"Resultados não mencionam método '{m}' usado na metodologia")

#     # 3. Verificar alinhamento de objetivos
#     if secao in ["metodologia", "resultados", "conclusao"]:
#         if context['resumo']:
#             objetivos = re.findall(r'objetivo.*?(?:é|são)\s*(.*?)(?:\.|,)', context['resumo'].lower())
#             for obj in objetivos:
#                 if obj not in novo_conteudo.lower():
#                     problemas.append(f"Objetivo '{obj}' do resumo não abordado")

#     return problemas if problemas else None
def analisar_secao(secao, conteudo):
    """Analisa uma seção específica retornando problemas e sugestões"""
    problemas = []
    
    # Verificação de elementos essenciais
    if secao == 'resumo':
        elementos = ['objetivo', 'método', 'resultado', 'conclusão']
        faltantes = [e for e in elementos if e not in conteudo.lower()]
        if faltantes:
            problemas.append(f"Elementos faltantes: {', '.join(faltantes)}")
    
    # Verificação de tamanho
    palavras = len(conteudo.split())
    if secao == 'resumo' and palavras < 100:
        problemas.append(f"Resumo muito curto ({palavras}/100 palavras mínimas)")
    
    return problemas, []
def verificar_coerencia_avancada(secao, conteudo):
    """Retorna erros de coerência em formato estruturado para correção pelo Gemini"""
    context = get_global_context()
    erros = []
    
    # Verificação de termos-chave
    termos_chave = []
    if context['palavras_chave']:
        termos_chave.extend([kw.strip().lower() for kw in context['palavras_chave'].split(',') if kw.strip()])
    if context['titulo']:
        termos_chave.extend([word.lower() for word in context['titulo'].split() if len(word) > 4])
    
    for termo in set(termos_chave):
        if termo and termo not in conteudo.lower():
            erros.append(f"Incluir o termo-chave: '{termo}'")

    # Verificações específicas por seção
    if secao == "metodologia":
        if context['resumo']:
            if "qualitativo" in context['resumo'].lower() and "qualitativo" not in conteudo.lower():
                erros.append("Mencionar abordagem qualitativa")
            if "quantitativo" in context['resumo'].lower() and "quantitativo" not in conteudo.lower():
                erros.append("Mencionar abordagem quantitativa")

    elif secao == "resultados":
        if context['metodologia']:
            if "entrevista" in context['metodologia'].lower() and "entrevista" not in conteudo.lower():
                erros.append("Relacionar aos dados de entrevista")
            if "experimento" in context['metodologia'].lower() and "experimento" not in conteudo.lower():
                erros.append("Relacionar aos experimentos realizados")

    return erros
def gerar_conteudo_com_correcao(secao):
    """Gera conteúdo automaticamente corrigindo problemas de coerência"""
    context = get_global_context()
    conteudo_atual = st.session_state.artigo[secao]
    
    # 1. Obter erros atuais
    erros = verificar_coerencia_avancada(secao, conteudo_atual)
    
    # 2. Construir prompt de correção
    prompt = f"""
    Você é um editor acadêmico especializado. Corrija esta seção de {secao} com base nos requisitos abaixo:

    CONTEXTO DO ARTIGO:
    - Título: {context['titulo']}
    - Palavras-chave: {context['palavras_chave']}
    - Objetivos do resumo: {context['resumo'][:200] if context['resumo'] else 'N/A'}

    CORREÇÕES NECESSÁRIAS:
    {chr(10).join(erros) if erros else 'Nenhum erro crítico detectado'}

    CONTEÚDO ATUAL:
    {conteudo_atual}

    INSTRUÇÕES:
    1. Mantenha o estilo acadêmico formal
    2. Integre todas as correções necessárias
    3. Preserve o conteúdo válido existente
    4. Retorne APENAS o texto corrigido, sem comentários
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Erro na geração: {str(e)}")
        return None
def gerar_conteudo_com_coerencia(secao):
    context = get_global_context()
    prompt = generate_section_prompt(secao, context)
    
    try:
        response = model.generate_content(prompt)
        conteudo_gerado = response.text
        
        # Verificação de coerência
        problemas = verificar_coerencia_avancada(secao, conteudo_gerado)
        
        if problemas:
            st.warning("Problemas de coerência detectados:")
            for p in problemas:
                st.error(p)
            
            st.info("Gerando versão alternativa...")
            prompt_corrigido = prompt + "\n\nCORREÇÕES NECESSÁRIAS:\n" + "\n".join(problemas)
            response = model.generate_content(prompt_corrigido)
            return response.text
        return conteudo_gerado
        
    except Exception as e:
        st.error(f"Erro na geração: {str(e)}")
        return None
def java_versao_compatível():
    try:
        resultado = subprocess.run(["java", "-version"], stderr=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
        saida = resultado.stderr if resultado.stderr else resultado.stdout

        # Extrair a versão do Java usando regex
        versao_match = re.search(r'version "(?P<versao>\d+)(\.(\d+))?', saida)
        if versao_match:
            major_version = int(versao_match.group('versao'))
            if major_version >= 17:
                return True, f"✅ Java {major_version} detectado: compatível."
            else:
                return False, f"❌ Java {major_version} detectado. LanguageTool requer Java 17 ou superior."
        else:
            return False, "❌ Não foi possível detectar a versão do Java."
    except FileNotFoundError:
        return False, "❌ Java não encontrado. Certifique-se de que está instalado e no PATH."

def armazenar_sugestoes_cache(problemas, sugestoes):
    """Armazena as sugestões de qualidade no session_state para exibição nos campos"""
    st.session_state.cache_sugestoes = {
        'titulo': [],
        'palavras_chave': [],
        'resumo': [],
        'introducao': [],
        'metodologia': [],
        'resultados': [],
        'conclusao': [],
        'referencias': []
    }
    
    # Mapeamento de problemas para seções
    mapeamento_secoes = {
        'título': 'titulo',
        'palavras-chave': 'palavras_chave',
        'resumo': 'resumo',
        'introdução': 'introducao',
        'metodologia': 'metodologia',
        'resultados': 'resultados',
        'conclusão': 'conclusao',
        'referências': 'referencias'
    }
    
    for problema, sugestao in zip(problemas, sugestoes):
        # Encontrar a seção relacionada ao problema
        secao_encontrada = None
        for termo, secao in mapeamento_secoes.items():
            if termo in problema.lower():
                secao_encontrada = secao
                break
        
        # Se não encontrou específico, verifica se é geral
        if not secao_encontrada:
            if 'gramaticais' in problema.lower():
                # Erros gramaticais podem estar em qualquer seção
                continue
            else:
                # Problema geral, adiciona a todas as seções
                for secao in st.session_state.cache_sugestoes.keys():
                    st.session_state.cache_sugestoes[secao].append((problema, sugestao))
        else:
            st.session_state.cache_sugestoes[secao_encontrada].append((problema, sugestao))
# def exibir_warnings_campos():
#     """Exibe os warnings nos campos com base no cache de sugestões"""
#     if 'cache_sugestoes' not in st.session_state:
#         return
    
#     # Para cada seção do artigo, verifica se há sugestões
#     for secao, sugestoes in st.session_state.cache_sugestoes.items():
#         if sugestoes:
#             # Cria um container de warning flutuante
#             warning_container = st.container()
#             with warning_container:
#                 st.warning(f"⚠️ Problemas na seção {secao.replace('_', ' ').title()}")
#                 for problema, sugestao in sugestoes:
#                     st.write(f"• **Problema:** {problema}")
#                     st.write(f"◦ **Sugestão:** {sugestao}")
#                 st.markdown("---")


def obter_warnings_por_secao(secao):
    """Retorna warnings específicos para uma seção"""
    if 'cache_sugestoes' not in st.session_state:
        return []
    
    # Mapeamento de seções para padrões de busca
    mapeamento = {
        'titulo': ['título', 'titulo'],
        'palavras_chave': ['palavras-chave', 'palavras chave'],
        'resumo': ['resumo'],
        'introducao': ['introdução', 'introducao'],
        'metodologia': ['metodologia'],
        'resultados': ['resultados'],
        'conclusao': ['conclusão', 'conclusao'],
        'referencias': ['referências', 'referencias']
    }
    
    # Filtra warnings específicos da seção
    warnings_secao = []
    for problema, sugestao in st.session_state.cache_sugestoes.get(secao, []):
        warnings_secao.append((problema, sugestao))
    
    # Adiciona problemas gerais que mencionam a seção
    for secao_key in st.session_state.cache_sugestoes.keys():
        for problema, sugestao in st.session_state.cache_sugestoes[secao_key]:
            if any(padrao in problema.lower() for padrao in mapeamento[secao]):
                if (problema, sugestao) not in warnings_secao:
                    warnings_secao.append((problema, sugestao))
    
    return warnings_secao

def exibir_warnings_secao(secao):
    """Exibe warnings apenas para a seção especificada"""
    warnings = obter_warnings_por_secao(secao)
    
    if warnings:
        nome_secao = secao.replace('_', ' ').title()
        with st.expander(f"⚠️ Problemas na {nome_secao}", expanded=True):
            for problema, sugestao in warnings:
                col1, col2 = st.columns([1, 4])
                with col1:
                    st.error("Problema")
                with col2:
                    st.error(problema)
                
                col1, col2 = st.columns([1, 4])
                with col1:
                    st.success("Sugestão")
                with col2:
                    st.success(sugestao)
                
                st.markdown("---")

def analisar_qualidade_artigo():
    """Analisa o artigo completo e fornece sugestões de melhoria"""
    problemas = []
    sugestoes = []
    expansoes = {}
    
    try:
        # 1. Verificação básica de preenchimento
        secoes_vazias = []
        for secao, conteudo in st.session_state.artigo.items():
            if secao not in ['id', 'data_criacao', 'data_atualizacao'] and not conteudo.strip():
                secoes_vazias.append(secao.replace('_', ' ').title())
        
        if secoes_vazias:
            problemas.append(f"Seções vazias: {', '.join(secoes_vazias)}")
        
        # 2. Verificação gramatical com tratamento de erros
        erros_gramaticais = 0
        try:
            tool = language_tool_python.LanguageTool('pt-BR')
            for secao, conteudo in st.session_state.artigo.items():
                if secao not in ['id', 'data_criacao', 'data_atualizacao'] and conteudo:
                    try:
                        matches = tool.check(conteudo)
                        erros_gramaticais += len(matches)
                    except Exception as e:
                        problemas.append(f"Erro ao verificar gramática na seção {secao}: {str(e)}")
        except Exception as e:
            problemas.append(f"Erro ao inicializar verificador gramatical: {str(e)}")
        
        if erros_gramaticais > 0:
            problemas.append(f"Erros gramaticais encontrados: {erros_gramaticais}")
            sugestoes.append("Revise o texto com atenção aos erros gramaticais destacados.")
        
        # 3. Análise do título
        titulo = st.session_state.artigo['titulo']
        if titulo:
            palavras_titulo = len(titulo.split())
            if palavras_titulo > 15:
                problemas.append("Título muito longo (mais de 15 palavras)")
                sugestoes.append("Considere reduzir o título para torná-lo mais conciso.")
            elif palavras_titulo < 5:
                problemas.append("Título muito curto (menos de 5 palavras)")
                sugestoes.append("Considere expandir o título para melhor descrever o conteúdo.")
            
            if not titulo[-1] in ['.', '?'] and not titulo[-1].isdigit():
                problemas.append("Título não termina com pontuação adequada")
                sugestoes.append("Considere adicionar um ponto final no título se for declarativo.")
        
        # 4. Análise do resumo com cálculo de expansão
        resumo = st.session_state.artigo['resumo']
        if resumo:
            palavras_resumo = len(resumo.split())
            paragrafos_resumo = len([p for p in resumo.split('\n') if p.strip()])
            
            if palavras_resumo < 100:
                problemas.append(f"Resumo muito curto ({palavras_resumo} palavras)")
                sugestoes.append("O resumo deve conter entre 100-250 palavras.")
                expansoes['resumo'] = {
                    'palavras_faltantes': 100 - palavras_resumo,
                    'paragrafos_faltantes': max(1 - paragrafos_resumo, 1)
                }
            elif palavras_resumo > 250:
                problemas.append(f"Resumo muito longo ({palavras_resumo} palavras)")
                sugestoes.append("O resumo deve conter entre 100-250 palavras.")
            
            # Verificar estrutura do resumo
            elementos_essenciais = ['objetivo', 'método', 'resultado', 'conclusão']
            faltantes = [elem for elem in elementos_essenciais if elem not in resumo.lower()]
            if faltantes:
                problemas.append(f"Elementos faltantes no resumo: {', '.join(faltantes)}")
                sugestoes.append("Certifique-se que o resumo contém: objetivo, método, resultados e conclusão.")
        
        # 5. Análise de palavras-chave
        palavras_chave = st.session_state.artigo['palavras_chave']
        if palavras_chave:
            palavras_lista = [p.strip() for p in palavras_chave.split(',') if p.strip()]
            num_palavras = len(palavras_lista)
            
            if num_palavras < 3:
                problemas.append(f"Poucas palavras-chave ({num_palavras})")
                sugestoes.append("Adicione pelo menos 3-5 palavras-chave relevantes.")
            elif num_palavras > 5:
                problemas.append(f"Muitas palavras-chave ({num_palavras})")
                sugestoes.append("Reduza para 3-5 palavras-chave mais relevantes.")
            
            # Verificar repetição no título
            termos_repetidos = [kw for kw in palavras_lista if kw.lower() in titulo.lower()]
            if not termos_repetidos:
                problemas.append("Nenhuma palavra-chave aparece no título")
                sugestoes.append("Considere incluir pelo menos uma palavra-chave no título.")
        
        # 6. Análise de citações nas referências
        referencias = st.session_state.artigo['referencias']
        if referencias:
            refs = [r.strip() for r in referencias.split('\n') if r.strip()]
            num_refs = len(refs)
            
            if num_refs < 5:
                problemas.append(f"Poucas referências ({num_refs})")
                sugestoes.append("Adicione pelo menos 5 referências relevantes.")
                expansoes['referencias'] = {
                    'palavras_faltantes': 0,
                    'paragrafos_faltantes': max(5 - num_refs, 1)
                }
            
            # Verificar formatação básica
            padroes_validos = 0
            padroes = [
                r'[A-ZÀ-Ú][a-zà-ú]+, [A-ZÀ-Ú]\. .* \(\d{4}\)\.',  # ABNT básico
                r'\([A-Za-z]+, \d{4}\)',  # APA básico
                r'http[s]?://',  # URLs
                r'doi:',  # DOI
            ]
            
            for ref in refs:
                if any(re.search(padrao, ref) for padrao in padroes):
                    padroes_validos += 1
            
            if padroes_validos / num_refs < 0.7:
                problemas.append(f"Muitas referências mal formatadas ({padroes_validos}/{num_refs} válidas)")
                sugestoes.append("Verifique a formatação das referências conforme normas ABNT ou APA.")
        
        # 7. Verificação de plágio básica e redundância
        texto_completo = ' '.join([v for k, v in st.session_state.artigo.items() 
                                 if k not in ['id', 'data_criacao', 'data_atualizacao']])
        
        # a. Palavras repetidas
        palavras = re.findall(r'\b\w{5,}\b', texto_completo.lower())
        contagem = Counter(palavras)
        palavras_repetidas = [(p, c) for p, c in contagem.items() if c > 10]
        if palavras_repetidas:
            problemas.append("Uso excessivo de algumas palavras: " + 
                           ", ".join([f"{p} ({c}x)" for p, c in palavras_repetidas[:3]]))
            sugestoes.append("Considere usar sinônimos para variar o vocabulário.")
        
        # b. Frases repetidas
        frases = re.findall(r'[^.!?]+[.!?]', texto_completo)
        frases_repetidas = [f for f in set(frases) if frases.count(f) > 1]
        if frases_repetidas:
            problemas.append(f"Frases repetidas encontradas: {len(frases_repetidas)}")
            sugestoes.append("Revise o texto para evitar repetição de frases idênticas.")
        
        # 8. Coerência entre seções
        if st.session_state.artigo['resumo'] and st.session_state.artigo['conclusao']:
            objetivos_resumo = set(re.findall(r'objetivo.*?(?:é|são)\s*(.*?)(?:\.|,)', 
                                          st.session_state.artigo['resumo'].lower()))
            conclusoes = set(re.findall(r'(?:conclu|demonstra|mostra).*?(que|os|as)\s*(.*?)(?:\.|,)', 
                                      st.session_state.artigo['conclusao'].lower()))
            
            objetivos_nao_abordados = []
            for obj in objetivos_resumo:
                if not any(obj in conc for conc in conclusoes):
                    objetivos_nao_abordados.append(obj)
            
            if objetivos_nao_abordados:
                problemas.append(f"Objetivos do resumo não abordados na conclusão: {', '.join(objetivos_nao_abordados[:3])}")
                sugestoes.append("Certifique-se que todos os objetivos declarados no resumo são abordados na conclusão.")
        
        # 9. Cálculo de expansões para outras seções
        secoes_principais = {
            'introducao': {'min_palavras': 200, 'min_paragrafos': 3},
            'metodologia': {'min_palavras': 300, 'min_paragrafos': 2},
            'resultados': {'min_palavras': 400, 'min_paragrafos': 3},
            'conclusao': {'min_palavras': 200, 'min_paragrafos': 2}
        }
        
        for secao, params in secoes_principais.items():
            conteudo = st.session_state.artigo.get(secao, '')
            if conteudo:
                palavras = len(conteudo.split())
                paragrafos = len([p for p in conteudo.split('\n') if p.strip()])
                
                if palavras < params['min_palavras'] or paragrafos < params['min_paragrafos']:
                    expansoes[secao] = {
                        'palavras_faltantes': max(params['min_palavras'] - palavras, 0),
                        'paragrafos_faltantes': max(params['min_paragrafos'] - paragrafos, 0)
                    }
        armazenar_sugestoes_cache(problemas, sugestoes)
        
        return problemas, sugestoes, expansoes
        
    except Exception as e:
        st.error(f"Erro inesperado na análise de qualidade: {str(e)}")
        return ["Erro na análise"], ["Recarregue a página e tente novamente"], {}
        return problemas, sugestoes, expansoes
        
    except Exception as e:
        st.error(f"Erro inesperado na análise de qualidade: {str(e)}")
        return ["Erro na análise"], ["Recarregue a página e tente novamente"], {}
def gerar_expansao_gemini(secao, conteudo_atual, expansao_necessaria):
    """Gera conteúdo adicional para atingir os requisitos mínimos"""
    try:
        with st.spinner(f"Gerando conteúdo adicional para {secao}..."):
            prompt = f"""
            Você é um assistente acadêmico. Expanda esta seção de {secao} para atender aos requisitos mínimos:
            
            CONTEÚDO ATUAL:
            {conteudo_atual}
            
            REQUISITOS:
            - Adicionar aproximadamente {expansao_necessaria['palavras_faltantes']} palavras
            - Adicionar {expansao_necessaria['paragrafos_faltantes']} parágrafos
            - Manter estilo acadêmico formal
            - Integrar perfeitamente com o conteúdo existente
            
            FORNECA:
            1. A versão COMPLETA da seção com as expansões necessárias
            2. Justificativa das alterações
            """
            
            response = model.generate_content(prompt)
            return response.text
            
    except Exception as e:
        return f"Erro ao gerar expansão: {str(e)}"

# def obter_sugestoes_gemini(secao, conteudo):
#     """Obtém sugestões do Gemini para melhorar uma seção específica"""
#     try:
#         with st.spinner("O Gemini está analisando sua seção..."):
#             prompt = f"""
#             Você é um assistente acadêmico especializado em redação científica. 
#             Forneça sugestões concisas para melhorar a seção de {secao} abaixo:
            
#             {conteudo}
            
#             Sugira melhorias específicas na estrutura, conteúdo e linguagem acadêmica.
#             """
            
#             response = model.generate_content(prompt)
#             return response.text
            
#     except Exception as e:
#         return f"Erro ao consultar o Gemini: {str(e)}"

# Função que aplica as melhorias ao conteúdo existente
def gerar_resumo_melhorado(resumo_original, sugestoes):
    """Usa o Gemini para gerar uma versão melhorada incorporando as sugestões"""
    prompt = f"""
    Você é um editor acadêmico. Reescreva este resumo incorporando as sugestões de melhoria:

    RESUMO ORIGINAL:
    {resumo_original}

    SUGESTÕES DE MELHORIA:
    {sugestoes}

    INSTRUÇÕES:
    1. Mantenha todas as informações válidas do original
    2. Incorpore as melhorias sugeridas
    3. Preserve o estilo acadêmico
    4. Não invente novos conteúdos não mencionados no original
    5. Retorne APENAS o texto revisado, sem comentários
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar versão melhorada: {str(e)}")
        return resumo_original  # Fallback para o original em caso de erro
def obter_sugestoes_gemini(secao, conteudo, problemas=None, erros_coerencia=None):
    """Gera sugestões considerando problemas específicos e contexto global"""
    context = get_global_context()
    
    prompt = f"""
    Como editor acadêmico, forneça sugestões para melhorar esta seção de {secao}.
    
    CONTEXTO DO ARTIGO:
    - Título: {context['titulo']}
    - Objetivos: {context['resumo'][:150] if context['resumo'] else 'N/A'}
    
    PROBLEMAS IDENTIFICADOS:
    {chr(10).join(problemas) if problemas else 'Nenhum problema grave detectado'}
    
    PROBLEMAS DE COERÊNCIA:
    {chr(10).join(erros_coerencia) if erros_coerencia else 'Nenhum problema de coerência'}
    
    CONTEÚDO ATUAL:
    {conteudo}
    
    SUA TAREFA:
    1. Liste 3-5 melhorias prioritárias
    2. Para cada uma, explique brevemente o motivo
    3. Sugira exemplos concretos de implementação
    4. Mantenha um tom construtivo e acadêmico
    
    FORMATE COMO:
    ### [Prioridade] Título da Sugestão
    **Motivo:** Explicação concisa
    **Exemplo:** "Você poderia... [exemplo concreto]"
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Erro ao gerar sugestões: {str(e)}"
def get_global_context():
    """Retorna o contexto completo do artigo para garantir coerência na geração."""
    return {
        'titulo': st.session_state.artigo['titulo'],
        'palavras_chave': st.session_state.artigo['palavras_chave'],
        'resumo': st.session_state.artigo['resumo'],
        'introducao': st.session_state.artigo['introducao'],
        'metodologia': st.session_state.artigo['metodologia'],
        'resultados': st.session_state.artigo['resultados'],
        'conclusao': st.session_state.artigo['conclusao']
    }

# def generate_section_prompt(section, context):
#     """Gera um prompt contextualizado para o Gemini baseado em todas as seções."""
#     prompts = {
#         'introducao': (
#             f"Com base no título '{context['titulo']}' e nas palavras-chave '{context['palavras_chave']}', "
#             "gere uma introdução acadêmica que: "
#             "1. Contextualize o tema; 2. Apresente o problema de pesquisa; 3. Defina os objetivos. "
#             "Mantenha coerência com o resumo: " + context['resumo'][:200] + "..."
#         ),
#         'metodologia': (
#             f"Considerando o objetivo declarado no resumo '{context['resumo'][:100]}'...', "
#             "descreva uma metodologia detalhada incluindo: "
#             "1. Tipo de estudo; 2. População/amostra; 3. Instrumentos; 4. Procedimentos. "
#             "Garanta que os métodos possam gerar os resultados sugeridos na introdução: " + context['introducao'][:100] + "..."
#         ),
#         'resultados': (
#             f"A partir da metodologia '{context['metodologia'][:100]}'...', "
#             "gere resultados plausíveis incluindo: "
#             "1. Dados quantitativos/qualitativos; 2. Análise estatística (se aplicável); "
#             "3. Relação com a literatura mencionada na introdução: " + context['introducao'][:100] + "..."
#         ),
#         'conclusao': (
#             f"Com base nestes resultados '{context['resultados'][:100]}'...', "
#             "elabore uma conclusão que: "
#             "1. Responda aos objetivos do resumo; 2. Destaque contribuições; "
#             "3. Sugira pesquisas futuras. Mantenha coerência com o título: " + context['titulo']
#         )
#     }
#     return prompts.get(section, f"Gere conteúdo acadêmico para a seção '{section}' mantendo coerência com o artigo completo.")

def generate_section_prompt(section, context):
    """Gera um prompt contextualizado para o Gemini baseado em todas as seções."""
    prompts = {
        'introducao': (
            f"Com base no título '{context['titulo']}' e nas palavras-chave '{context['palavras_chave']}', "
            "gere uma introdução acadêmica que: "
            "1. Contextualize o tema; 2. Apresente o problema de pesquisa; 3. Defina os objetivos. "
            "Mantenha coerência com o resumo: " + (context['resumo'][:200] + "..." if context['resumo'] else "")
        ),
        'metodologia': (
            f"Considerando o objetivo declarado no resumo '{context['resumo'][:100]}'...', "
            "descreva uma metodologia detalhada incluindo: "
            "1. Tipo de estudo; 2. População/amostra; 3. Instrumentos; 4. Procedimentos. "
            "Garanta que os métodos possam gerar os resultados sugeridos na introdução: " + (context['introducao'][:100] + "..." if context['introducao'] else "")
        ),
        'resultados': (
            f"A partir da metodologia '{context['metodologia'][:100]}'...', "
            "gere resultados plausíveis incluindo: "
            "1. Dados quantitativos/qualitativos; 2. Análise estatística (se aplicável); "
            "3. Relação com a literatura mencionada na introdução: " + (context['introducao'][:100] + "..." if context['introducao'] else "")
        ),
        'conclusao': (
            f"Com base nestes resultados '{context['resultados'][:100]}'...', "
            "elabore uma conclusão que: "
            "1. Responda aos objetivos do resumo; 2. Destaque contribuições; "
            "3. Sugira pesquisas futuras. Mantenha coerência com o título: " + context['titulo']
        )
    }
    return prompts.get(section, f"Gere conteúdo acadêmico para a seção '{section}' mantendo coerência com o artigo completo.")
## 5. Inicialização do Sistema


init_db()

# Estado da sessão
if 'artigo' not in st.session_state:
    st.session_state.artigo = {
        'id': None,
        'titulo': '',
        'palavras_chave': '',
        'resumo': '',
        'introducao': '',
        'metodologia': '',
        'resultados': '',
        'conclusao': '',
        'referencias': ''
    }

if 'secao_atual' not in st.session_state:
    st.session_state.secao_atual = "Título"

## 6. Interface - Barra Lateral
# passos_labels={
#     1:"Título", 2:"Palavras-chave", 3:"Resumo", 4:"Introdução", 
#          5:"Metodologia", 6:"Resultados", 7:"Conclusão", 8:"Referências"
# }
# with st.sidebar:
#     st.header("📚 Editor de Artigo Científico")
    
#     for passo_num, label in passos_labels.items():
#             if st.button(label, key=f"nav_passo_{passo_num}"):
#                 st.session_state.secao_atual = passo_num
#                 st.rerun()    


passos_labels = {
    1: "Título",
    2: "Palavras-chave",
    3: "Resumo",
    4: "Introdução",
    5: "Metodologia",
    6: "Resultados",
    7: "Conclusão",
    8: "Referências"
}

with st.sidebar:
    st.header("📚 Editor de Artigo Científico")

    # Obtém os rótulos em uma lista para usar com st.radio
    opcoes_radio = list(passos_labels.values())

    # Define o valor padrão para o radio button, se st.session_state.secao_atual já existir
    # Caso contrário, define como a primeira opção (Título)
    indice_selecionado = 0
    if "secao_atual" in st.session_state:
        # Encontra o índice da seção atual para que o radio button esteja pré-selecionado
        for num, label in passos_labels.items():
            if num == st.session_state.secao_atual:
                indice_selecionado = list(passos_labels.values()).index(label)
                break
    
    secao_selecionada_label = st.radio(
        "Navegar para a seção:",
        options=opcoes_radio,
        index=indice_selecionado, # Define a opção pré-selecionada
        key="radio_navegacao"
    )

    # Encontra o número do passo correspondente ao rótulo selecionado
    for passo_num, label in passos_labels.items():
        if label == secao_selecionada_label:
            if "secao_atual" not in st.session_state or st.session_state.secao_atual != passo_num:
                st.session_state.secao_atual = passo_num
                st.rerun()

# Exemplo de como você usaria st.session_state.secao_atual em outras partes do seu aplicativo
st.write(f"Você está atualmente na seção: {passos_labels.get(st.session_state.get('secao_atual', 1))}")

st.divider()
    # st.subheader("🔍 Análise de Qualidade")
    # if st.button("Analisar Artigo Completo"):
    #     problemas, sugestoes, expansoes = analisar_qualidade_artigo()
    #     st.session_state.ultima_analise = datetime.now()


    #     if problemas:
    #             st.error("📊 Relatório de Qualidade Gerado")
    #             st.write(f"Última análise: {st.session_state.ultima_analise.strftime('%H:%M:%S')}")
    #             st.rerun()
    #     else:
    #             st.success("✔ Artigo atende aos critérios básicos!")
            
    
    # st.divider()
    # st.subheader("🗃️ Gerenciamento de Artigos")
    
    # artigos = carregar_artigos()
    # opcoes_artigos = ["Novo Artigo"] + [f"{id}: {titulo[:30]}... ({data.split()[0]})" for id, titulo, data in artigos]
    
    # artigo_selecionado = st.selectbox(
    #     "Selecione um artigo:",
    #     opcoes_artigos,
    #     index=0
    # )
    
    # if artigo_selecionado != "Novo Artigo" and st.button("Carregar Artigo"):
    #     artigo_id = int(artigo_selecionado.split(":")[0])
    #     artigo_db = carregar_artigo_por_id(artigo_id)
        
    #     if artigo_db:
    #         st.session_state.artigo = {
    #             'id': artigo_db[0],
    #             'titulo': artigo_db[1],
    #             'palavras_chave': artigo_db[2],
    #             'resumo': artigo_db[3],
    #             'introducao': artigo_db[4],
    #             'metodologia': artigo_db[5],
    #             'resultados': artigo_db[6],
    #             'conclusao': artigo_db[7],
    #             'referencias': artigo_db[8]
    #         }
    #         st.rerun()
    
    # col1, col2 = st.columns(2)
    # with col1:
    #     if st.button("💾 Salvar Artigo"):
    #         salvar_artigo()
    # with col2:
    #     if st.session_state.artigo['id'] and st.button("🗑️ Excluir Artigo"):
    #         excluir_artigo(st.session_state.artigo['id'])
    #         st.session_state.artigo = {
    #             'id': None,
    #             'titulo': '',
    #             'palavras_chave': '',
    #             'resumo': '',
    #             'introducao': '',
    #             'metodologia': '',
    #             'resultados': '',
    #             'conclusao': '',
    #             'referencias': ''
    #         }
    #         st.rerun()
    
    # st.divider()
    # st.subheader("📤 Exportar Artigo")
    # if st.button("📄 Exportar para Word"):
    #     exportar_para_word()

## 7. Interface - Conteúdo Principal
st.title("✍️ Editor de Artigos")

# Renderização condicional das seções
if st.session_state.secao_atual == 1:
    st.header("📌 Título do Artigo")
    st.markdown("""**Instruções:**  
    - Use a criatividade para instigar a leitura  
    - Seja claro e objetivo  
    - Reflita o conteúdo principal da pesquisa  
    - Máximo recomendado: 2 linhas""")
        # Exibir warnings específicos para o título
    exibir_warnings_secao('titulo')
    st.session_state.artigo['titulo'] = st.text_area(
        "Digite o título do seu artigo:",
        value=st.session_state.artigo['titulo'],
        height=100,
        key="titulo_input"
    )
    
    palavras = len(st.session_state.artigo['titulo'].split())
    st.metric("Palavras", palavras, "Recomendado: 10-15", delta_color="off")
    
    if palavras < 10:
        if st.button("✨ Sugerir título completo"):
            sugestao = gerar_expansao_gemini(
                "Título",
                st.session_state.artigo['titulo'],
                {'tipo': 'completa', 'palavras_faltantes': 10-palavras, 'paragrafos_faltantes': 1}
            )
            st.session_state.sugestao_titulo = sugestao
    
    if 'sugestao_titulo' in st.session_state:
        st.subheader("Sugestão de Título")
        st.write(st.session_state.sugestao_titulo)
        if st.button("✅ Aplicar sugestão de título"):
            st.session_state.artigo['titulo'] = st.session_state.sugestao_titulo
            del st.session_state.sugestao_titulo
            st.rerun()

    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            ''
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =2
                st.rerun()           

elif st.session_state.secao_atual == 2:
    st.header("🔑 Palavras-chave")
    st.markdown("""**Instruções:**  
    - Termos que representam seu trabalho  
    - Utilize conceitos centrais da pesquisa  
    - Separe por vírgulas  
    - Recomendado: 3-5 palavras""")
    
    st.session_state.artigo['palavras_chave'] = st.text_area(
        "Digite as palavras-chave:",
        value=st.session_state.artigo['palavras_chave'],
        height=100,
        key="palavras_chave_input"
    )

    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 1
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =3
                st.rerun() 

elif st.session_state.secao_atual == 3:
    st.header("📝 Resumo")
    st.markdown("""**Instruções:**  
    - Apresente objetivos, métodos e conclusões principais  
    - Seja conciso e informativo  
    - Evite citações e abreviações  
    - Tamanho recomendado: 100-250 palavras""")
      # Exibir warnings específicos para o resumo
    exibir_warnings_secao('resumo')
    st.session_state.artigo['resumo'] = st.text_area(
        "Digite o resumo do artigo:",
        value=st.session_state.artigo['resumo'],
        height=300,
        key="resumo_input"
    )
    
    palavras = len(st.session_state.artigo['resumo'].split())
    paragrafos = len([p for p in st.session_state.artigo['resumo'].split('\n') if p.strip()])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Palavras", palavras, "Recomendado: 100-250", delta_color="off")
    with col2:
        st.metric("Parágrafos", paragrafos, "Recomendado: 1", delta_color="off")
    
    if palavras < 100 or paragrafos < 1:
        if st.button(f"✨ Gerar resumo completo (faltam {max(0, 100-palavras)} palavras)"):
            sugestao = gerar_expansao_gemini(
                "Resumo",
                st.session_state.artigo['resumo'],
                {'tipo': 'completa', 'palavras_faltantes': max(100-palavras, 50), 'paragrafos_faltantes': 1}
            )
            st.session_state.sugestao_resumo = sugestao
    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 2
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =4
                st.rerun() 
    if 'sugestao_resumo' in st.session_state:
        st.subheader("Sugestão de Resumo Completo")
        st.write(st.session_state.sugestao_resumo)
        if st.button("✅ Aplicar sugestão de resumo"):
            st.session_state.artigo['resumo'] = st.session_state.sugestao_resumo
            del st.session_state.sugestao_resumo
            st.rerun()
    
    if st.button("🔍 Obter Sugestões de Melhoria para o Resumo"):
        if not st.session_state.artigo['resumo']:
            st.warning("Por favor, escreva um rascunho inicial do resumo para receber sugestões.")
        else:
            with st.spinner("Analisando resumo e gerando sugestões..."):
                try:
                    # 1. Gerar sugestões detalhadas
                    sugestao = obter_sugestoes_gemini("Resumo", st.session_state.artigo['resumo'])
                    
                    # 2. Armazenar no estado da sessão
                    st.session_state.sugestao_melhoria_resumo = {
                        'texto': sugestao,
                        'resumo_original': st.session_state.artigo['resumo']  # Guarda cópia do original
                    }
                    
                except Exception as e:
                    st.error(f"Erro ao gerar sugestões: {str(e)}")

# Se existem sugestões geradas, mostrar o painel de aplicação
# if 'sugestao_melhoria_resumo' in st.session_state:
#     st.subheader("💡 Sugestões de Melhoria")
#     st.markdown(st.session_state.sugestao_melhoria_resumo['texto'])
    
#     # Container para os botões de ação
#     col1, col2, col3 = st.columns([1,1,2])
    
#     with col1:
#         if st.button("✅ Aplicar Melhorias Automaticamente", 
#                     help="Substitui o resumo atual pela versão melhorada"):
#             with st.spinner("Aplicando melhorias..."):
#                 try:
#                     # Gera a versão melhorada incorporando todas as sugestões
#                     novo_resumo = gerar_resumo_melhorado(
#                         st.session_state.sugestao_melhoria_resumo['resumo_original'],
#                         st.session_state.sugestao_melhoria_resumo['texto']
#                     )
                    
#                     st.session_state.artigo['resumo'] = novo_resumo
#                     st.success("Melhorias aplicadas com sucesso!")
#                     st.rerun()
                    
#                 except Exception as e:
#                     st.error(f"Falha ao aplicar melhorias: {str(e)}")
    
#     with col2:
#         if st.button("❌ Descartar Sugestões"):
#             del st.session_state.sugestao_melhoria_resumo
#             st.rerun()
    
#     with col3:
#         if st.button("📋 Visualizar Comparação"):
#             st.subheader("Comparação: Original vs. Melhorado")
            
#             col_orig, col_novo = st.columns(2)
#             with col_orig:
#                 st.markdown("**Versão Original**")
#                 st.text_area("Original", 
#                             value=st.session_state.sugestao_melhoria_resumo['resumo_original'],
#                             height=300,
#                             key="original_resumo",
#                             label_visibility="collapsed")
            
#             with col_novo:
#                 st.markdown("**Versão Melhorada**")
#                 novo_resumo = gerar_resumo_melhorado(
#                     st.session_state.sugestao_melhoria_resumo['resumo_original'],
#                     st.session_state.sugestao_melhoria_resumo['texto']
#                 )
#                 st.text_area("Melhorado", 
#                             value=novo_resumo,
#                             height=300,
#                             key="melhorado_resumo",
#                             label_visibility="collapsed")
#     if 'sugestao_melhoria_resumo' in st.session_state:
#         st.subheader("Sugestões de Melhoria")
#         st.write(st.session_state.sugestao_melhoria_resumo)

# [...] (Implementar padrão similar para as outras seções: Introdução, Metodologia, etc.)
elif st.session_state.secao_atual == 4:
    st.header("📖 Introdução")
    st.markdown("""**Instruções:**  
    - Contextualize o tema da pesquisa  
    - Apresente o problema investigado  
    - Defina os objetivos do trabalho  
    - Explique a importância do estudo  
    - Tamanho sugerido: 1-3 páginas  
    """)
    exibir_warnings_secao('introducao')
    st.session_state.artigo['introducao'] = st.text_area(
        "Digite a introdução:",
        value=st.session_state.artigo['introducao'],
        height=500,
        key="introducao_input"
    )
    
    palavras = len(st.session_state.artigo['introducao'].split())
    paragrafos = len([p for p in st.session_state.artigo['introducao'].split('\n') if p.strip()])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Palavras", palavras, "Recomendado: 200-800", delta_color="off")
    with col2:
        st.metric("Parágrafos", paragrafos, "Recomendado: 3", delta_color="off")
    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 3
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =5
                st.rerun() 
    if palavras < 200 or paragrafos < 3:
        if st.button(f"✨ Gerar introdução completa (faltam {max(0, 200-palavras)} palavras)"):
            sugestao = gerar_expansao_gemini(
                "Introdução",
                st.session_state.artigo['introducao'],
                {'tipo': 'completa', 'palavras_faltantes': max(200-palavras, 150), 'paragrafos_faltantes': max(3-paragrafos, 2)}
            )
            st.session_state.sugestao_introducao = sugestao
    
    if 'sugestao_introducao' in st.session_state:
        st.subheader("Sugestão de Introdução Completa")
        st.write(st.session_state.sugestao_introducao)
        if st.button("✅ Aplicar sugestão à Introdução"):
            st.session_state.artigo['introducao'] = st.session_state.sugestao_introducao
            del st.session_state.sugestao_introducao
            st.rerun()

if 'secao_atual' not in st.session_state:
    st.session_state.secao_atual = "Título"
elif st.session_state.secao_atual == 5:
    st.header("🔬 Metodologia")
    
    # Botão de geração com contexto
    if st.button("✨ Gerar Metodologia com Gemini"):
        context = get_global_context()
    
    # Verificação dos pré-requisitos
        if not context['resumo'] or not context['introducao']:
                st.warning("""
                **Pré-requisitos incompletos!**  
                Para garantir uma metodologia coerente, preencha primeiro:
                - ✍️ Resumo (especialmente os objetivos)
                - 📖 Introdução (contexto do estudo)
                """)
        else:
            with st.spinner("🧠 Gerando metodologia alinhada com seu artigo..."):
                try:
                        # Geração com verificação em tempo real
                        conteudo_gerado = gerar_conteudo_com_coerencia('metodologia')
                        
                        if conteudo_gerado:
                            # Análise pós-geração
                            problemas_coerencia = verificar_coerencia_avancada('metodologia', conteudo_gerado)
                            
                            if problemas_coerencia:
                                # Se houver problemas, tentar corrigir automaticamente
                                st.warning("""
                                **Ajustes necessários detectados:**  
                                O Gemini identificou oportunidades de melhoria na coerência
                                """)
                                
                                with st.expander("🔍 Ver detalhes dos ajustes"):
                                    for problema in problemas_coerencia:
                                        st.write(f"- {problema}")
                                
                                if st.button("🔄 Aplicar correções automaticamente", key="corrigir_metodologia"):
                                    conteudo_corrigido = gerar_conteudo_com_correcao('metodologia')
                                    if conteudo_corrigido:
                                        st.session_state.artigo['metodologia'] = conteudo_corrigido
                                        st.success("Metodologia ajustada com sucesso!")
                                        st.rerun()
                            else:
                                # Se estiver tudo ok, aplicar diretamente
                                st.session_state.artigo['metodologia'] = conteudo_gerado
                                st.success("Metodologia gerada com sucesso!")
                                st.rerun()
                        
                except Exception as e:
                        st.error(f"""
                        **Erro na geração:**  
                        {str(e)}  
                        *Sugestão:* Tente novamente ou ajuste manualmente os pré-requisitos
                        """)
                    
            # Mostrar preview mesmo com erros (se houver conteúdo)
            if 'conteudo_gerado' in locals() and conteudo_gerado:
                with st.expander("📝 Pré-visualização da metodologia gerada"):
                    st.write(conteudo_gerado)
                    st.info("""
                    **Antes de aplicar:**  
                    Verifique se o conteúdo atende às expectativas.  
                    Você pode editar manualmente ou solicitar ajustes automáticos acima.
                    """)
        
    # Campo de texto principal (que estava faltando)
    st.session_state.artigo['metodologia'] = st.text_area(
        "Digite a metodologia:",
        value=st.session_state.artigo['metodologia'],
        height=500,
        key="metodologia_input"
    )
    
    palavras = len(st.session_state.artigo['metodologia'].split())
    paragrafos = len([p for p in st.session_state.artigo['metodologia'].split('\n') if p.strip()])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Palavras", palavras, "Recomendado: 300-1000", delta_color="off")
    with col2:
        st.metric("Parágrafos", paragrafos, "Recomendado: 2", delta_color="off")
    
        col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 4
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =6
                st.rerun() 

    if palavras < 300 or paragrafos < 2:
        if st.button(f"✨ Gerar metodologia completa (faltam {max(0, 300-palavras)} palavras)"):
            sugestao = gerar_expansao_gemini(
                "Metodologia",
                st.session_state.artigo['metodologia'],
                {'tipo': 'completa', 'palavras_faltantes': max(300-palavras, 200), 'paragrafos_faltantes': max(2-paragrafos, 1)}
            )
            st.session_state.sugestao_metodologia = sugestao
    
    if 'sugestao_metodologia' in st.session_state:
        st.subheader("Sugestão de Metodologia Completa")
        st.write(st.session_state.sugestao_metodologia)
        if st.button("✅ Aplicar sugestão à Metodologia"):
            st.session_state.artigo['metodologia'] = st.session_state.sugestao_metodologia
            del st.session_state.sugestao_metodologia
            st.rerun()

elif st.session_state.secao_atual == 6:
    st.header("📊 Resultados e Discussão")
    st.markdown("""**Instruções:**  
    - Apresente e interprete os resultados  
    - Relacione com a literatura existente  
    - Destaque descobertas importantes  
    - Discuta limitações do estudo  
    - Tamanho sugerido: 6-8 páginas  
    """)
    exibir_warnings_secao('resultados')
    st.session_state.artigo['resultados'] = st.text_area(
        "Digite os resultados e discussão:",
        value=st.session_state.artigo['resultados'],
        height=500,
        key="resultados_input"
    )
    
    palavras = len(st.session_state.artigo['resultados'].split())
    paragrafos = len([p for p in st.session_state.artigo['resultados'].split('\n') if p.strip()])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Palavras", palavras, "Recomendado: 400-1200", delta_color="off")
    with col2:
        st.metric("Parágrafos", paragrafos, "Recomendado: 3", delta_color="off")

        col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 5
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =7
                st.rerun()     
    
    if palavras < 400 or paragrafos < 3:
        if st.button(f"✨ Gerar resultados completos (faltam {max(0, 400-palavras)} palavras)"):
            sugestao = gerar_expansao_gemini(
                "Resultados",
                st.session_state.artigo['resultados'],
                {'tipo': 'completa', 'palavras_faltantes': max(400-palavras, 300), 'paragrafos_faltantes': max(3-paragrafos, 2)}
            )
            st.session_state.sugestao_resultados = sugestao
    
    if 'sugestao_resultados' in st.session_state:
        st.subheader("Sugestão de Resultados Completos")
        st.write(st.session_state.sugestao_resultados)
        if st.button("✅ Aplicar sugestão aos Resultados"):
            st.session_state.artigo['resultados'] = st.session_state.sugestao_resultados
            del st.session_state.sugestao_resultados
            st.rerun()

elif st.session_state.secao_atual == 7:
    st.header("🎯 Conclusão")
    st.markdown("""**Instruções:**  
    - Resuma os principais achados  
    - Destaque as contribuições do estudo  
    - Sugira pesquisas futuras  
    - Evite repetir resultados detalhadamente  
    - Tamanho sugerido: 3-5 páginas  
    """)
    
    st.session_state.artigo['conclusao'] = st.text_area(
        "Digite a conclusão:",
        value=st.session_state.artigo['conclusao'],
        height=500,
        key="conclusao_input"
    )
    
    palavras = len(st.session_state.artigo['conclusao'].split())
    paragrafos = len([p for p in st.session_state.artigo['conclusao'].split('\n') if p.strip()])
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Palavras", palavras, "Recomendado: 200-800", delta_color="off")
    with col2:
        st.metric("Parágrafos", paragrafos, "Recomendado: 2", delta_color="off")
    
    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 6
                st.rerun()
    with col2_nav:
            if st.button("Próximo ➡️", key="passo3_next"):
                st.session_state.secao_atual =8
                st.rerun() 

    if palavras < 200 or paragrafos < 2:
        if st.button(f"✨ Gerar conclusão completa (faltam {max(0, 200-palavras)} palavras)"):
            sugestao = gerar_expansao_gemini(
                "Conclusão",
                st.session_state.artigo['conclusao'],
                {'tipo': 'completa', 'palavras_faltantes': max(200-palavras, 150), 'paragrafos_faltantes': max(2-paragrafos, 1)}
            )
            st.session_state.sugestao_conclusao = sugestao
    
    if 'sugestao_conclusao' in st.session_state:
        st.subheader("Sugestão de Conclusão Completa")
        st.write(st.session_state.sugestao_conclusao)
        if st.button("✅ Aplicar sugestão à Conclusão"):
            st.session_state.artigo['conclusao'] = st.session_state.sugestao_conclusao
            del st.session_state.sugestao_conclusao
            st.rerun()
elif st.session_state.secao_atual == 8:
    st.header("📚 Referências Bibliográficas")
    st.markdown("""**Instruções:**  
    - Liste todas as fontes citadas  
    - Formate conforme normas APA ou ABNT  
    - Inclua livros, artigos, sites e outros materiais relevantes 
    - Organize em ordem alfabética  
    - Verifique a completude das informações""")
    
    with st.expander("🔍 Buscar Referências Automáticas (Crossref)"):
        termo_busca = st.text_input("Digite um termo para buscar referências:")
        if st.button("Buscar no Crossref"):
            resultados = buscar_referencias_crossref(termo_busca)
            if resultados:
                st.session_state.artigo['referencias'] = resultados
                st.rerun()
    
    st.session_state.artigo['referencias'] = st.text_area(
        "Digite as referências bibliográficas:",
        value=st.session_state.artigo['referencias'],
        height=500,
        key="referencias_input"
    )
    
    refs = [r.strip() for r in st.session_state.artigo['referencias'].split('\n') if r.strip()]
    st.metric("Referências", len(refs), "Mínimo recomendado: 5", delta_color="off")
    
    if len(refs) < 5:
        if st.button("✨ Sugerir referências adicionais"):
            sugestao = gerar_expansao_gemini(
                "Referências",
                st.session_state.artigo['referencias'],
                {'tipo': 'completa', 'palavras_faltantes': 0, 'paragrafos_faltantes': max(5-len(refs), 3)}
            )
            st.session_state.sugestao_referencias = sugestao
    
    if 'sugestao_referencias' in st.session_state:
        st.subheader("Sugestão de Referências Adicionais")
        st.write(st.session_state.sugestao_referencias)
        if st.button("✅ Aplicar sugestão de referências"):
            st.session_state.artigo['referencias'] = st.session_state.sugestao_referencias
            del st.session_state.sugestao_referencias
            st.rerun()
    col1_nav, col2_nav = st.columns(2)
    with col1_nav:
            if st.button("⬅️ Anterior", key="passo3_prev"):
                st.session_state.secao_atual = 7
                st.rerun()
    with col2_nav:
             if st.button("🔄 Começar Novo Artigo", key="passo15_new"):
                for key in list(st.session_state.keys()):
                    if key not in ['user_projects']:
                        del st.session_state[key]
                #initialize_session_state()
                st.rerun()       

# Navegação entre seções



# Sidebar with buttons to navigate through sections
# col1, col2, col3 = st.columns([2, 2, 8]) 

# with col1:  
#     if st.button("⬅️ Voltar"):
#         navegar(-1)   # Navigate back

# with col2: 
#     if st.button("➡️ Próximo"):
#         navegar(+1)  # Navigate forward




# Visualização do artigo completo
st.divider()
with st.expander("👁️ Visualizar Artigo Completo"):
    if st.session_state.artigo['titulo']:
        st.subheader(st.session_state.artigo['titulo'])
        st.caption(f"Palavras-chave: {st.session_state.artigo['palavras_chave']}")
        
        st.subheader("Resumo")
        st.write(st.session_state.artigo['resumo'])
        
        st.subheader("Introdução")
        st.write(st.session_state.artigo['introducao'])
        
        st.subheader("Metodologia")
        st.write(st.session_state.artigo['metodologia'])
        
        st.subheader("Resultados e Discussão")
        st.write(st.session_state.artigo['resultados'])
        
        st.subheader("Conclusão")
        st.write(st.session_state.artigo['conclusao'])
        
        st.subheader("Referências")
        st.write(st.session_state.artigo['referencias'])
    else:
        st.warning("Nenhum conteúdo disponível para visualização.")
st.subheader("🗃️ Gerenciamento de Artigos")
    
artigos = carregar_artigos()
opcoes_artigos = ["Novo Artigo"] + [f"{id}: {titulo[:30]}... ({data.split()[0]})" for id, titulo, data in artigos]
    
artigo_selecionado = st.selectbox(
        "Selecione um artigo:",
        opcoes_artigos,
        index=0
    )
    
if artigo_selecionado != "Novo Artigo" and st.button("Carregar Artigo"):
        artigo_id = int(artigo_selecionado.split(":")[0])
        artigo_db = carregar_artigo_por_id(artigo_id)
        
        if artigo_db:
            st.session_state.artigo = {
                'id': artigo_db[0],
                'titulo': artigo_db[1],
                'palavras_chave': artigo_db[2],
                'resumo': artigo_db[3],
                'introducao': artigo_db[4],
                'metodologia': artigo_db[5],
                'resultados': artigo_db[6],
                'conclusao': artigo_db[7],
                'referencias': artigo_db[8]
            }
            st.rerun()
    
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    if st.button("🔍Analisar Artigo"):
        
            problemas, sugestoes, expansoes = analisar_qualidade_artigo()
            st.session_state.ultima_analise = datetime.now()


            if problemas:
                    st.error("📊 Relatório de Qualidade Gerado")
                    st.write(f"Última análise: {st.session_state.ultima_analise.strftime('%H:%M:%S')}")
                    st.rerun()
            else:
                    st.success("✔ Artigo atende aos critérios básicos!")
with col5:
    # Botão para reiniciar o artigo
    if st.button("🔄 Começar Novo Artigo", key="passo15_new"):
        for key in list(st.session_state.keys()):
            if key not in ['user_projects']:
                del st.session_state[key]
        #initialize_session_state()
        st.rerun()
with col2:
        if st.button("💾 Salvar Artigo"):
            salvar_artigo()
with col4:
        if st.session_state.artigo['id'] and st.button("🗑️ Excluir Artigo"):
            excluir_artigo(st.session_state.artigo['id'])
            st.session_state.artigo = {
                'id': None,
                'titulo': '',
                'palavras_chave': '',
                'resumo': '',
                'introducao': '',
                'metodologia': '',
                'resultados': '',
                'conclusao': '',
                'referencias': ''
            }
            st.rerun()
with col3:
    if st.button("📄 Exportar para Word"):
        exportar_para_word()
    




