"""Explorador da API — só leitura (GET). Para testar endpoints e ver os nomes reais dos campos."""
import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from belle_core import BelleClient, slug, tokens_dos_secrets

UNIDADES = ["Videira", "Joaçaba", "Treze Tílias", "Campos Novos", "Rota"]
EXEMPLOS = {
    "Clientes (página 0)": ("clientes", "pagina=0\ncodEstab=1"),
    "Sessões de planos": ("relatorios/sessoes_planos", "codEstab=1\ndtVndIni=01/10/2026\ndtVndFim=19/11/2026"),
    "Atendimentos": ("relatorios/relatorio_atendimentos",
                     "codEstab=1\ndtInicio=01/11/2026\ndtFim=19/11/2026\nstatusAgendamento=Atendido"),
    "Demonstrativo de vendas": ("relatorios/demonstrativo_vendas_detalhado",
                                "codEstab=1\ndtInicio=01/11/2026\ndtFim=19/11/2026"),
    "Agendamentos em aberto": ("agendamentos", "codEstab=1\ndtInicio=20/11/2026\ndtFim=30/11/2026"),
    "Serviços": ("servico/listar", ""),
    "Outro (digite)": ("", ""),
}
st.markdown('<div class="hero"><h1>Explorador da API</h1><p>Consulta de leitura (GET) para ver como o Belle responde.</p></div>',
            unsafe_allow_html=True)
tokens = tokens_dos_secrets()
c1, c2 = st.columns([1, 2])
unidade = c1.selectbox("Unidade", UNIDADES)
ex = c2.selectbox("Exemplo", list(EXEMPLOS))
cam0, prm0 = EXEMPLOS[ex]
caminho = st.text_input("Caminho (depois de .../v1.0/)", cam0, key=f"cam_{ex}")
prm_txt = st.text_area("Parâmetros (um por linha: nome=valor)", prm0, key=f"prm_{ex}", height=120)
if st.button("Consultar", type="primary"):
    prm = dict(l.split("=", 1) for l in prm_txt.splitlines() if "=" in l)
    prm = {k.strip(): v.strip() for k, v in prm.items()}
    tok = tokens.get(slug(unidade))
    if not tok:
        st.error("Sem token dessa unidade nos secrets.")
    else:
        try:
            status, corpo = BelleClient(tok).get_raw(caminho, prm)
        except Exception as e:
            st.error(f"Falhou: {e}")
        else:
            (st.success if status == 200 else st.error)(f"HTTP {status}")
            lista = corpo if isinstance(corpo, list) else next(
                (v for v in corpo.values() if isinstance(v, list)), None) if isinstance(corpo, dict) else None
            if lista:
                df = pd.json_normalize(lista, sep="_")
                st.write(f"**{len(lista)} registros** · colunas: " + ", ".join(map(str, df.columns)))
                st.dataframe(df.astype(str).head(200), use_container_width=True)
            with st.expander("Resposta bruta (primeiros 5000 caracteres)"):
                st.code(json.dumps(corpo, ensure_ascii=False, indent=1)[:5000] if not isinstance(corpo, str) else corpo)
            st.download_button("⬇️ Baixar JSON", json.dumps(corpo, ensure_ascii=False, indent=1).encode("utf-8"),
                               "resposta.json", "application/json")
