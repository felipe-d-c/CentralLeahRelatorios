"""Central de Relatórios Léah — senha + menu dos relatórios."""
import hmac
import sys
import time
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parent
MOD = RAIZ / "modulos"
for p in (RAIZ, MOD / "comercial", MOD / "vendas", MOD / "fichas", MOD / "relatorios", MOD / "explorador"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

st.set_page_config(page_title="Central de Relatórios · Léah", page_icon="💜", layout="wide")

SENHA_PADRAO = "12qwaszx"   # use o secret SENHA_CENTRAL para trocar sem mexer no código


def _senha():
    try:
        return str(st.secrets.get("SENHA_CENTRAL", SENHA_PADRAO))
    except Exception:
        return SENHA_PADRAO


if not st.session_state.get("autenticado"):
    st.markdown("## 💜 Central de Relatórios Léah")
    with st.form("login"):
        digitada = st.text_input("Senha", type="password")
        if st.form_submit_button("Entrar", type="primary"):
            if hmac.compare_digest(digitada.encode(), _senha().encode()):
                st.session_state["autenticado"] = True
                st.rerun()
            time.sleep(1.5)   # freia tentativas em sequência
            st.error("Senha incorreta.")
    st.stop()


def _rodar(pasta, arquivo="app.py"):
    # o 1º st.* de cada página deve vir depois do set_page_config da Central (já chamado acima)
    """Executa o app de um módulo como página. O set_page_config dele é ignorado (a Central já definiu)."""
    import runpy
    original = st.set_page_config
    st.set_page_config = lambda *a, **k: None
    try:
        runpy.run_path(str(MOD / pasta / arquivo), run_name=f"central_{pasta}")
    finally:
        st.set_page_config = original


def pagina_inicio():
    st.markdown("## 💜 Central de Relatórios Léah")
    st.write("Escolha o relatório no menu à esquerda.")
    st.page_link(pg_comercial, label="Relatório Comercial — metas, comissão, vendedoras", icon="📈")
    st.page_link(pg_vendas, label="Relatório de Vendas por clínica", icon="🧾")
    st.page_link(pg_fichas, label="Fichas de clientes — sessões dos planos ativos", icon="📋")
    st.page_link(pg_rel, label="Relatórios — inativos, aniversariantes, planos, boletos, NPS, agenda, estoque…", icon="🗂️")
    st.page_link(pg_pacote, label="Pacote completo — vários relatórios num ZIP", icon="📦")
    st.page_link(pg_exp, label="Explorador da API — testar endpoints", icon="🔎")


pg_comercial = st.Page(lambda: _rodar("comercial"), title="Relatório Comercial", icon="📈", url_path="comercial")
pg_vendas = st.Page(lambda: _rodar("vendas"), title="Relatório de Vendas", icon="🧾", url_path="vendas")
pg_fichas = st.Page(lambda: _rodar("fichas"), title="Fichas de clientes", icon="📋", url_path="fichas")
pg_rel = st.Page(lambda: _rodar("relatorios"), title="Relatórios", icon="🗂️", url_path="relatorios")
pg_pacote = st.Page(lambda: _rodar("relatorios", "pacote.py"), title="Pacote completo", icon="📦", url_path="pacote")
pg_exp = st.Page(lambda: _rodar("explorador"), title="Explorador da API", icon="🔎", url_path="explorador")
pg_inicio = st.Page(pagina_inicio, title="Início", icon="🏠", default=True)

with st.sidebar:
    if st.button("Sair"):
        st.session_state.clear()
        st.rerun()

st.navigation([pg_inicio, pg_rel, pg_pacote, pg_comercial, pg_vendas, pg_fichas, pg_exp]).run()
