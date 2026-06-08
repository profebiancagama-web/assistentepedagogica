import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from google import genai
from google.genai import types
from pypdf import PdfReader
from supabase import create_client, Client
import io, os, re
import pandas as pd

st.set_page_config(page_title="Gerador Olegário Pro", layout="wide")

SUPABASE_URL = "https://hgdpahtyvoridomkktkt.supabase.co"
SUPABASE_KEY = "sb_publishable_yxUHijFIrK7h2yznh8ySKw_4eBFxAlI"
MINHA_CHAVE = "AQ.Ab8RN6LW-l2a1Qvu-8ZCD75G0OlT5RSVhxJpS3Pa6VcNQ88AFg"

if "supabase" not in st.session_state: st.session_state.supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
if "autenticado" not in st.session_state: st.session_state.autenticado = False
if "client" not in st.session_state: st.session_state.client = genai.Client(api_key=MINHA_CHAVE)
if "resultado" not in st.session_state: st.session_state.resultado = ""
if "historico_chat" not in st.session_state:
    st.session_state.historico_chat = [
        {"role": "model", "text": "Olá! Eu sou a B.IA, sua assistente pedagógica e companheira de trabalho. Sobre o que você gostaria de conversar ou planejar hoje? Estou pronta para ajudar em qualquer assunto!"}
    ]

# --- TELA DE LOGIN ---
if not st.session_state.autenticado:
    st.title("🔐 B.IA: Assistente Pedagógica")
    aba_login, aba_cadastro = st.tabs(["Entrar no Sistema", "Cadastrar Novo Professor"])
    with aba_login:
        email = st.text_input("E-mail funcional:", key="login_email")
        senha = st.text_input("Senha de acesso:", type="password", key="login_senha")
        if st.button("🔓 VALIDAR CREDENCIAIS", key="btn_login_exec"):
            try:
                res = st.session_state.supabase.auth.sign_in_with_password({"email": email, "password": senha})
                if res.session:
                    st.session_state.autenticado = True
                    st.success("✅ Credenciais validadas com sucesso!")
            except:
                st.error("E-mail ou senha incorretos no banco de dados.")
        if st.session_state.autenticado:
            if st.button("🚀 CLIQUE AQUI PARA ENTRAR NO PAINEL"):
                st.rerun()
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
st.title("🧠 B.IA: Assistente Pedagógica")
with st.sidebar:
    st.header("⚙️ Painel do Professor")
    net = st.toggle("🔓 Liberar Internet (Google Search)", value=False)
    st.write("---")
    if st.button("🚪 Sair do Sistema"):
        st.session_state.autenticado = False
        st.session_state.supabase.auth.sign_out()
        st.rerun()

prof = st.text_input("Nome do(a) Professor(a):", value="")
comp = st.text_input("Componente Curricular / Disciplina:")
turma = st.text_input("Ano / Turma:")
st.write("---")
arqs = st.file_uploader("Arquivos de Referência Pedagógica (Até 3):", type=["pdf", "docx", "txt"], accept_multiple_files=True)

txt_ref = ""
if arqs and len(arqs) <= 3:
    for arq in arqs:
        try:
            if arq.type == "text/plain": txt_ref += arq.read().decode("utf-8") + "\n\n"
            elif arq.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                txt_ref += "\n".join([p.text for p in Document(arq).paragraphs]) + "\n\n"
            elif arq.type == "application/pdf":
                for pg in PdfReader(arq).pages: txt_ref += (pg.extract_text() or "") + "\n"
            st.success(f"✅ {arq.name} carregado para planejamento!")
        except: st.error("Erro ao ler arquivo.")

# 🧠 FORMATADOR DO WORD
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
                if tg in
