import streamlit as st

st.set_page_config(
    page_title="AcadêmicoPro",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Título Principal e Boas-Vindas ---
st.title("🎓 Bem-vindo ao AcadêmicoPro!")

st.markdown("""
    ## Seu Assistente Inteligente para Projectos de Pesquisa

    O **AcadêmicoPro** é a sua ferramenta definitiva para **simplificar e otimizar** o desenvolvimento de projectos acadêmicos e científicos. Desde a formulação da ideia até a geração do documento final, nós guiamos você em cada etapa.

    **Com o AcadêmicoPro, você pode:**
    * Definir seu tema, problema e objectivos de pesquisa.
    * Estruturar seu enquadramento teórico e metodologia.
    * Organizar cronogramas, e referências.
    * **Gerar seu projecto completo em um documento Word formatado!**
""")

# --- CHAMADA PARA AÇÃO PRINCIPAL: BOTÃO DE LOGIN GRANDE ---
st.markdown("---") # Separador visual

st.subheader("Pronto para começar seu projecto?")

# Criar um botão de login grande e centralizado
# Podemos usar colunas para centralizar um pouco
col_empty1, col_button, col_empty2 = st.columns([1, 2, 1])
with col_button:
    #st.markdown("##") # Espaço para o botão
    if st.button("🔐 **ENTRAR OU CRIAR MINHA CONTA AGORA!**", type="primary", use_container_width=True):
        st.switch_page("pages/2_🔐_Login.py")
    st.markdown("##") # Espaço após o botão
    st.markdown("<p style='text-align: center; font-size: 0.9em; color: #666;'>Acesse todas as funcionalidades da plataforma.</p>", unsafe_allow_html=True)

st.markdown("---") # Separador visual

# --- Acesso Rápido às Páginas (opcional, já que o botão acima é o foco) ---
# Você pode manter isso como uma alternativa ou remover se o botão principal for suficiente.
# st.subheader("Ou navegue diretamente:")
# col_nav1, col_nav2 = st.columns(2)
# with col_nav1:
#     st.page_link("pages/1_🏠_Trabalhos.py", label="🏠 Ir para Trabalhos (após Login)")
# with col_nav2:
#     st.page_link("pages/2_🔐_Login.py", label="🔐 Ir para Página de Login/Registro")

# --- Rodapé ---
st.markdown(
    """
    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p style="font-size: 0.8em;">© 2025 AcadêmicoPro - Todos os direitos reservados</p>
        <p style="font-size: 0.8em;">Desenvolvido com 💖 por geraldomung@outlook.com</p>
    </div>
    """,
    unsafe_allow_html=True
)