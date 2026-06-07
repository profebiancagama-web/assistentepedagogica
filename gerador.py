# Atualizacao do sistema
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from google import genai
from google.genai import types
from pypdf import PdfReader
from supabase import create_client, Client
import io, os, re

st.set_page_config(page_title="Gerador Olegário Pro", layout="wide")

SUPABASE_URL = "https://hgdpahtyvoridomkktkt.supabase.co"
SUPABASE_KEY = "sb_publishable_yxUHijFIrK7h2yznh8ySKw_4eBFxAlI"
MINHA_CHAVE = "AQ.Ab8RN6LW-l2a1Qvu-8ZCD75G0OlT5RSVhxJpS3Pa6VcNQ88AFg"

if "supabase" not in st.session_state: st.session_state.supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
if "autenticado" not in st.session_state: st.session_state.autenticado = False
if "client" not in st.session_state: st.session_state.client = genai.Client(api_key=MINHA_CHAVE)
if "chat" not in st.session_state: st.session_state.chat = []
if "resultado" not in st.session_state: st.session_state.resultado = ""

# --- TELA DE LOGIN CORRIGIDA SEM LOOPING ---
if not st.session_state.autenticado:
    st.title("🔐 Acesso Restrito: E.E.B. Prefeito Olegário Bernardes")
    aba_login, aba_cadastro = st.tabs(["Entrar no Sistema", "Cadastrar Novo Professor"])
    with aba_login:
        email = st.text_input("E-mail funcional:", key="login_email")
        senha = st.text_input("Senha de acesso:", type="password", key="login_senha")
        if st.button("🔓 VALIDAR CREDENCIAIS", key="btn_login_exec"):
            try:
                res = st.session_state.supabase.auth.sign_in_with_password({"email": email, "password": senha})
                if res.session:
                    st.session_state.autenticado = True
                    st.success("✅ Usuário validado com sucesso!")
            except:
                st.error("E-mail ou senha incorretos no banco de dados.")
        
        if st.session_state.autenticado:
            if st.button("🚀 ENTRAR NO GERADOR AGORA"):
                st.skip()
                
    with aba_cadastro:
        novo_email = st.text_input("E-mail do Professor:", key="cad_email")
        nova_senha = st.text_input("Crie uma Senha (mínimo 6 dígitos):", type="password", key="cad_senha")
        if st.button("📝 CADASTRAR PROFESSOR", key="btn_cad_exec"):
            try:
                st.session_state.supabase.auth.sign_up({"email": novo_email, "password": nova_senha})
                st.success("Professor cadastrado com sucesso! Volte à aba anterior.")
            except Exception as e: st.error(f"Erro ao cadastrar: {e}")
    st.stop()

# --- INTERFACE PRINCIPAL ---
st.title("Sistema Pedagógico Pro: E.E.B. Olegário Bernardes")
with st.sidebar:
    st.header("⚙️ Painel do Professor")
    net = st.toggle("🔓 Liberar Internet (Google Search)", value=False)
    st.write("---")
    if st.button("🚪 Sair do Sistema"):
        st.session_state.autenticado = False
        st.session_state.supabase.auth.sign_out()
        st.skip()

prof = st.text_input("Nome do(a) Professor(a):", value="")
comp = st.text_input("Componente Curricular / Disciplina:")
turma = st.text_input("Ano / Turma:")
st.write("---")
arqs = st.file_uploader("Arquivos de Referência (Até 3):", type=["pdf", "docx", "txt"], accept_multiple_files=True)

txt_ref = ""
if arqs and len(arqs) <= 3:
    for arq in arqs:
        try:
            if arq.type == "text/plain": txt_ref += arq.read().decode("utf-8") + "\n\n"
            elif arq.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                txt_ref += "\n".join([p.text for p in Document(arq).paragraphs]) + "\n\n"
            elif arq.type == "application/pdf":
                for pg in PdfReader(arq).pages: txt_ref += (pg.extract_text() or "") + "\n"
            st.success(f"✅ {arq.name} ok!")
        except: st.error("Erro ao ler arquivo.")
# 🧠 TRADUTOR DE ASTERISCOS PARA NEGRITO REAL NO WORD
def aplicar_formatacao_inteligente(paragrafo, texto_com_tags):
    paragrafo.text = ""
    partes = re.split(r'(\*\*.*?\*\*)', texto_com_tags)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**'):
            texto_limpo = parte[2:-2]
            if texto_limpo:
                run = paragrafo.add_run(texto_limpo)
                run.bold = True
        elif parte:
            run = paragrafo.add_run(parte)
        try:
            for r in paragrafo.runs: r.font.name, r.font.size, r.font.color.rgb = 'Arial', Pt(12), RGBColor(0,0,0)
        except: pass

def preencher_word(nome_modelo, dados_tags):
    if not os.path.exists(nome_modelo): return None
    doc = Document(nome_modelo)
    def processar_paragrafo(p):
        for tg, tx in dados_tags.items():
            if tg in p.text:
                if tg in ["{{CORPO_PROVA}}", "{{TEXTO_RELATORIO}}"] or tg.startswith("{{"):
                    texto_final = p.text.replace(tg, tx)
                    aplicar_formatacao_inteligente(p, texto_final)
                else:
                    p.text = p.text.replace(tg, tx)
                    for r in p.runs: r.font.name, r.font.size, r.font.color.rgb = 'Arial', Pt(12), RGBColor(0,0,0)
    for p in doc.paragraphs: processar_paragrafo(p)
    for t in doc.tables:
        for l in t.rows:
            for c in l.cells:
                for p in c.paragraphs: processar_paragrafo(p)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

st.write("---")
aba1, aba2, aba3, aba4 = st.tabs(["📅 Plano Anual", "📝 Plano Mensal/Quinzenal", "✍️ Avaliações/Atividades", "📊 Relatórios"])
cfg = types.GenerateContentConfig()

with aba1:
    if st.button("✨ GERAR PLANO ANUAL", key="b1"):
        with st.spinner("Processando..."):
            pt = f"Crie um plano anual de {comp} ({turma}). Use as tags: [COMPETENCIAS_GERAIS], [COMPETENCIAS_ESPECIFICAS], [CONCEITOS1], [OBJETO1], [HABILIDADES1], [CONCEITOS2], [OBJETO2], [HABILIDADES2], [CONCEITOS3], [OBJETO3], [HABILIDADES3], [INSTRUMENTOS], [REFERENCIAS]. Use asteriscos duplos para marcar negritos.\n\nREF:\n{txt_ref}"
            st.session_state.resultado = st.session_state.client.models.generate_content(model='gemini-2.5-flash', contents=pt, config=cfg).text
            st.session_state.modelo_atual = "modelo_anual.docx"

with aba2:
    mes = st.selectbox("Selecione o Mês:", ["Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    duracao = st.text_input("Tempo de duração (ex: '4 aulas', '15 dias'):", value="15 dias")
    cmd_mensal = st.text_input("Foco temático do plano:")
    if st.button("✨ GERAR ESTRUTURA MENSAL", key="b2"):
        with st.spinner("Processando..."):
            pt = f"Crie um plano de aula para {mes} de {comp} ({turma}). Foco: {cmd_mensal}. Separe o texto estritamente por: [AREA], [HABILIDADES], [OBJETO], [CRITERIOS], [METODOLOGIA], [INSTRUMENTOS], [REFERENCIAS]. Use asteriscos duplos para marcar negritos.\n\nREF:\n{txt_ref}"
            st.session_state.resultado = st.session_state.client.models.generate_content(model='gemini-2.5-flash', contents=pt, config=cfg).text
            st.session_state.modelo_atual = "modelo_mensal.docx"
            st.session_state.duracao_input = duracao

with aba3:
    t_av = st.selectbox("Tipo de Atividade:", ["Prova Objetiva/Discursiva", "Recuperação Paralela", "Trabalho Dirigido", "Lista de Exercícios"])
    qst = st.slider("Questões:", 1, 10, 5)
    if st.button("✨ GERAR ATIVIDADE/AVALIAÇÃO", key="b3"):
        with st.spinner("Processando..."):
            pt = f"Crie uma atividade do tipo {t_av} com {qst} questões de {comp} ({turma}) com GABARITO. Use asteriscos duplos para marcar negritos nos enunciados."
            st.session_state.resultado = st.session_state.client.models.generate_content(model='gemini-2.5-flash', contents=pt, config=cfg).text
            st.session_state.modelo_atual = "modelo_avaliacao.docx"

with aba4:
    t_re = st.selectbox("Tipo Relatório:", ["Desempenho da Turma", "Aluno PDI/AEE"])
    ctx = st.text_area("Contexto do Aluno/Turma:")
    if st.button("✨ GERAR RELATÓRIO", key="b4"):
        with st.spinner("Processando..."):
            pt = f"Escreva um relatório do tipo {t_re} para {comp} ({turma}). Contexto: {ctx}. Use asteriscos duplos para marcar negritos."
            st.session_state.resultado = st.session_state.client.models.generate_content(model='gemini-2.5-flash', contents=pt, config=cfg).text
            st.session_state.modelo_atual = "modelo_avaliacao.docx"

if st.session_state.resultado:
    st.write("---")
    tags_map = {"{{PROFESSOR}}": prof, "{{COMPONENTE}}": comp, "{{TURMA}}": turma, "{{CORPO_PROVA}}": st.session_state.resultado, "{{TEXTO_RELATORIO}}": st.session_state.resultado}
    if st.session_state.modelo_atual == "modelo_anual.docx":
        def ext(tg, tx):
            m = re.search(rf"\[{tg}\](.*?)(?=\[\w+\]|$)", tx, re.DOTALL)
            return m.group(1).strip() if m else "Em branco."
        tags_map.update({"{{COMPETENCIAS_GERAIS}}": ext("COMPETENCIAS_GERAIS", st.session_state.resultado), "{{COMPETENCIAS_ESPECIFICAS}}": ext("COMPETENCIAS_ESPECIFICAS", st.session_state.resultado), "{{CONCEITOS1}}": ext("CONCEITOS1", st.session_state.resultado), "{{OBJETO_CONHECIMENTO1}}": ext("OBJETO1", st.session_state.resultado), "{{HABILIDADES1}}": ext("HABILIDADES1", st.session_state.resultado), "{{CONCEITOS2}}": ext("CONCEITOS2", st.session_state.resultado), "{{OBJETO_CONHECIMENTO2}}": ext("OBJETO2", st.session_state.resultado), "{{HABILIDADES2}}": ext("HABILIDADES2", st.session_state.resultado), "{{CONCEITOS3}}": ext("CONCEITOS3", st.session_state.resultado), "{{OBJETO_CONHECIMENTO3}}": ext("OBJETO3", st.session_state.resultado), "{{HABILIDADES3}}": ext("HABILIDADES3", st.session_state.resultado), "{{INSTRUMENTOS}}": ext("INSTRUMENTOS", st.session_state.resultado), "{{REFERENCIAS}}": ext("REFERENCIAS", st.session_state.resultado)})
    elif st.session_state.modelo_atual == "modelo_mensal.docx":
        def ext_m(tg, tx):
            m = re.search(rf"\[{tg}\](.*?)(?=\[\w+\]|$)", tx, re.DOTALL)
            return m.group(1).strip() if m else "Em branco."
        tags_map.update({"{{AREA_CONHECIMENTO}}": ext_m("AREA", st.session_state.resultado), "{{HABILIDADES_MENSAL}}": ext_m("HABILIDADES", st.session_state.resultado), "{{OBJETO_MENSAL}}": ext_m("OBJETO", st.session_state.resultado), "{{CRITERIOS_MENSAL}}": ext_m("CRITERIOS", st.session_state.resultado), "{{METODOLOGIA_MENSAL}}": ext_m("METODOLOGIA", st.session_state.resultado), "{{INSTRUMENTOS_MENSAL}}": ext_m("INSTRUMENTOS", st.session_state.resultado), "{{DURACAO_MENSAL}}": st.session_state.get("duracao_input", "15 dias"), "{{REFERENCIAS_MENSAL}}": ext_m("REFERENCIAS", st.session_state.resultado)})
    w_bytes = preencher_word(st.session_state.modelo_atual, tags_map)
    if w_bytes: st.download_button("📥 BAIXAR DOCUMENTO NO MODELO OFICIAL (.DOCX)", data=w_bytes, file_name=f"Documento_{comp}.docx", key="dl_f")
    else: st.error("⚠️ Verifique os arquivos de modelo na pasta.")
    st.subheader("📄 Conteúdo Gerado:")
    st.markdown(st.session_state.resultado)
