"""Mais relatórios — escolhe o tipo, as unidades e o período; gera online direto da API."""
import datetime as dt
import sys
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from belle_core import tokens_dos_secrets
from reports import base
from reports.base import CampoAusente, Ctx
from reports.catalogo import GRUPOS, REGISTRO

UNIDADES = ["Videira", "Joaçaba", "Treze Tílias", "Campos Novos", "Rota"]
tokens = tokens_dos_secrets()
S = st.session_state

st.markdown('<div class="hero"><h1>Relatórios</h1><p>Tudo direto da API do Belle, sem subir planilha.</p></div>',
            unsafe_allow_html=True)


def widget_param(r, p):
    k = f"p_{r.key}_{p.key}"
    if p.kind == "bool":
        return st.checkbox(p.label, value=bool(p.default), key=k, help=p.help or None)
    if p.kind == "select":
        opts = p.options or [p.default]
        idx = opts.index(str(p.default)) if str(p.default) in opts else 0
        if p.key == "mes" and k not in S:
            idx = dt.date.today().month - 1
        return st.selectbox(p.label, opts, index=idx, key=k, help=p.help or None)
    if p.kind == "text":
        return st.text_area(p.label, value=str(p.default), key=k, help=p.help or None)
    return st.number_input(p.label, value=int(p.default), step=1, key=k, help=p.help or None)


opcoes = {f"{r.grupo} · {r.titulo}": r for g in GRUPOS for r in REGISTRO.values() if r.grupo == g}
escolha = st.selectbox("Tipo de relatório", list(opcoes))
r = opcoes[escolha]
st.caption(r.desc)

c1, c2 = st.columns([2, 2])
unidades = c1.multiselect("Unidades", UNIDADES, default=UNIDADES, key=f"u_{r.key}")
ini = fim = None
if r.periodo is not None:
    hoje = dt.date.today()
    padrao = (hoje - dt.timedelta(days=r.periodo[0]), hoje + dt.timedelta(days=r.periodo[1]))
    per = c2.date_input("Período", padrao, format="DD/MM/YYYY", key=f"d_{r.key}")
    if isinstance(per, tuple) and len(per) == 2:
        ini, fim = per
    else:
        st.info("Escolha a data inicial e a final.")
        st.stop()
params = {p.key: widget_param(r, p) for p in r.params}
if r.key == "aniversarios":
    params["mes"] = int(params["mes"])

salvar_gh = st.checkbox("Guardar o resultado no GitHub ao gerar", value=True, key=f"sg_{r.key}")
memo = S.setdefault("_memo", {})
chave_res = f"res_{r.key}"


def gerar():
    ctx = Ctx(tokens, unidades, memo, S.get("_overrides", {}))
    with st.spinner("Consultando a API do Belle… (pode levar alguns minutos na primeira vez)"):
        try:
            res = r.run(ctx, ini, fim, params)
        except CampoAusente as e:
            S["_faltando"] = (e.campo, e.colunas)
            return
    S.pop("_faltando", None)
    S[chave_res] = (res, ctx.erros, dt.datetime.now().strftime("%d/%m/%Y %H:%M"))
    if salvar_gh and any(not d.empty for d in res.sheets.values()):
        try:
            base.salvar_resultado(f"{r.key}", res)
        except Exception as e:
            ctx.erros["salvar no GitHub"] = str(e)


b1, b2 = st.columns([1, 1])
if b1.button("Gerar relatório", type="primary", disabled=not unidades):
    S.pop("_faltando", None)
    gerar()
if b2.button("Abrir último resultado salvo"):
    try:
        got = base.carregar_resultado(r.key)
        if got:
            S[chave_res] = (got[1], {}, f"salvo em {got[0]}")
        else:
            st.warning("Ainda não há resultado salvo desse relatório.")
    except Exception as e:
        st.error(f"Não consegui ler o GitHub: {e}")

if "_faltando" in S:
    campo, cols = S["_faltando"]
    st.warning(f"A API devolveu o dado com nomes diferentes do esperado. Qual coluna é **{campo}**?")
    esc = st.selectbox("Coluna", cols, key="falt_sel")
    if st.button("Usar essa coluna e gerar de novo"):
        S.setdefault("_overrides", {})[campo] = esc
        S.pop("_faltando", None)
        gerar()
        st.rerun()

if chave_res in S:
    res, erros, quando = S[chave_res]
    st.markdown("---")
    st.caption(f"Resultado de {quando}")
    for u, e in erros.items():
        st.warning(f"{u}: {e}")
    for n in res.notas:
        st.info(n)
    if res.kpis:
        cols = st.columns(len(res.kpis))
        for c, (lab, val) in zip(cols, res.kpis):
            c.metric(lab, val)
    if res.sheets:
        abas = st.tabs([f"{k} ({len(v)})" for k, v in res.sheets.items()])
        for aba, (nome, df) in zip(abas, res.sheets.items()):
            with aba:
                cfg = {"WhatsApp": st.column_config.LinkColumn("WhatsApp", display_text="abrir")} \
                    if "WhatsApp" in df.columns else None
                st.dataframe(df, use_container_width=True, hide_index=True, column_config=cfg)
        d1, d2, d3 = st.columns(3)
        nome_arq = f"{r.key}_{dt.date.today():%Y%m%d}"
        d1.download_button("⬇️ Excel", base.xlsx_bytes(res.sheets, r.titulo), f"{nome_arq}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("⬇️ CSV (aba principal)", res.principal.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           f"{nome_arq}.csv", "text/csv")
        secoes = res.pdf or [(next(iter(res.sheets)), res.principal)]
        d3.download_button("⬇️ PDF", base.pdf_bytes(r.titulo, secoes), f"{nome_arq}.pdf", "application/pdf")
