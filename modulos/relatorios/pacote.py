"""Pacote completo — vários relatórios de uma vez, num ZIP de planilhas."""
import datetime as dt
import io
import sys
import zipfile
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from belle_core import tokens_dos_secrets
from reports import base
from reports.base import CampoAusente, Ctx
from reports.catalogo import REGISTRO

UNIDADES = ["Videira", "Joaçaba", "Treze Tílias", "Campos Novos", "Rota"]
tokens = tokens_dos_secrets()

st.markdown('<div class="hero"><h1>Pacote completo</h1><p>Vários relatórios de uma vez, num ZIP só.</p></div>',
            unsafe_allow_html=True)
c1, c2 = st.columns(2)
unidades = c1.multiselect("Unidades", UNIDADES, default=UNIDADES)
hoje = dt.date.today()
per = c2.date_input("Período (para os relatórios que usam período)", (hoje - dt.timedelta(days=30), hoje),
                    format="DD/MM/YYYY")
padrao = [k for k, r in REGISTRO.items() if r.pacote]
sel = st.multiselect("Relatórios do pacote", list(REGISTRO), default=padrao,
                     format_func=lambda k: f"{REGISTRO[k].grupo} · {REGISTRO[k].titulo}")
st.caption("Cada relatório usa os parâmetros padrão dele (ex.: inativos há 30+ dias). "
           "Os que falharem entram no ZIP em `ERROS.txt`; os demais seguem normalmente.")

if st.button("Gerar pacote", type="primary", disabled=not (unidades and sel and isinstance(per, tuple) and len(per) == 2)):
    ini, fim = per
    memo = st.session_state.setdefault("_memo", {})
    ctx = Ctx(tokens, unidades, memo, st.session_state.get("_overrides", {}))
    buf, erros = io.BytesIO(), []
    barra = st.progress(0.0, text="Começando…")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, k in enumerate(sel):
            r = REGISTRO[k]
            barra.progress(i / len(sel), text=f"{r.titulo}…")
            prm = {p.key: (int(p.default) if p.key == "mes" else p.default) for p in r.params}
            if k == "aniversarios":
                prm["mes"] = hoje.month
            try:
                ctx.erros = {}
                res = r.run(ctx, ini, fim, prm)
                z.writestr(f"{i + 1:02d}_{k}.xlsx", base.xlsx_bytes(res.sheets, r.titulo))
                for u, e in ctx.erros.items():
                    erros.append(f"{r.titulo} · {u}: {e}")
            except CampoAusente as e:
                erros.append(f"{r.titulo}: coluna '{e.campo}' não encontrada. Colunas: {', '.join(map(str, e.colunas))}")
            except Exception as e:
                erros.append(f"{r.titulo}: {e}")
        if erros:
            z.writestr("ERROS.txt", "\n".join(erros))
    barra.empty()
    st.session_state["pacote"] = (buf.getvalue(), erros)

if "pacote" in st.session_state:
    z, erros = st.session_state["pacote"]
    for e in erros:
        st.warning(e)
    st.download_button(f"⬇️ Baixar pacote ({len(z) / 1e6:.1f} MB)", z, f"pacote_leah_{dt.date.today():%Y%m%d}.zip",
                       "application/zip")
