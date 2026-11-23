import datetime as dt
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import streamlit as st

import relatorio as rel
from belle_api import (UNIDADES, atualizar_cache, cache_cobre, cache_do_dia, hoje_br, caminho_cache, carregar_cache, carregar_tokens,
                       gerar_relatorio_de_dados, gerar_relatorio_unidade, slug, DIAGNOSTICO,
                       vendedoras_do_periodo)

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Relatório Comercial · Léah Estética",
    page_icon="💜",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# CSS CUSTOMIZADO — identidade visual Léah
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Importa fonte elegante */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&display=swap');

    /* Reset e fundo */
    .stApp {
        background: linear-gradient(160deg, #fdf2f7 0%, #f8fafc 40%, #f0fdfa 100%);
        font-family: 'Inter', sans-serif;
    }

    /* Remove padding excessivo do container principal */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 960px;
    }

    /* Header / Hero */
    .hero {
        background: linear-gradient(135deg, #ed2a78 0%, #c2185b 50%, #04c6c6 100%);
        border-radius: 20px;
        padding: 2.2rem 2.5rem;
        margin-bottom: 2rem;
        color: white;
        box-shadow: 0 12px 40px rgba(237, 42, 120, 0.25);
        position: relative;
        overflow: hidden;
    }
    .hero::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -20%;
        width: 60%;
        height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.12) 0%, transparent 70%);
        pointer-events: none;
    }
    .hero h1 {
        font-family: 'Playfair Display', serif;
        font-size: 2.1rem;
        font-weight: 700;
        margin: 0 0 0.35rem 0;
        letter-spacing: -0.02em;
    }
    .hero p {
        margin: 0;
        opacity: 0.92;
        font-size: 1.05rem;
        font-weight: 400;
    }

    /* Cards */
    .card {
        background: white;
        border-radius: 16px;
        padding: 1.5rem 1.75rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.06);
        border: 1px solid rgba(237, 42, 120, 0.08);
        margin-bottom: 1.25rem;
    }
    .card-title {
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #ed2a78;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Inputs */
    .stDateInput > div > div,
    .stMultiSelect > div > div {
        border-radius: 12px !important;
        border: 1.5px solid #fce7f3 !important;
        background: #fff !important;
    }
    .stDateInput > div > div:focus-within,
    .stMultiSelect > div > div:focus-within {
        border-color: #ed2a78 !important;
        box-shadow: 0 0 0 3px rgba(237, 42, 120, 0.12) !important;
    }

    /* Botões */
    .stButton > button {
        border-radius: 12px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        padding: 0.65rem 1.5rem !important;
        transition: all 0.2s ease !important;
        border: none !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #ed2a78, #c2185b) !important;
        color: white !important;
        box-shadow: 0 4px 14px rgba(237, 42, 120, 0.35) !important;
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(237, 42, 120, 0.45) !important;
    }
    .stButton > button[kind="secondary"],
    .stButton > button:not([kind="primary"]) {
        background: white !important;
        color: #64748b !important;
        border: 1.5px solid #e2e8f0 !important;
    }
    .stButton > button:not([kind="primary"]):hover {
        border-color: #ed2a78 !important;
        color: #ed2a78 !important;
    }

    /* Download buttons */
    .stDownloadButton > button {
        border-radius: 12px !important;
        background: linear-gradient(135deg, #04c6c6, #00a8a8) !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(4, 198, 198, 0.3) !important;
        width: 100%;
        padding: 0.7rem 1.2rem !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(4, 198, 198, 0.4) !important;
    }

    /* Status / progress */
    .stStatus {
        border-radius: 12px !important;
        border: 1px solid #fce7f3 !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        font-weight: 600 !important;
        color: #64748b !important;
        border-radius: 12px !important;
    }
    .streamlit-expanderContent {
        border-radius: 0 0 12px 12px !important;
        background: #fafafa !important;
    }

    /* Alertas */
    .stAlert {
        border-radius: 12px !important;
    }

    /* Labels */
    label {
        font-weight: 500 !important;
        color: #334155 !important;
        font-size: 0.9rem !important;
    }

    /* Footer sutil */
    .footer {
        text-align: center;
        margin-top: 2.5rem;
        padding-top: 1.5rem;
        border-top: 1px solid #fce7f3;
        color: #94a3b8;
        font-size: 0.8rem;
    }

    /* Badge de clínicas */
    .badge {
        display: inline-block;
        background: #fce7f3;
        color: #ed2a78;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.25rem 0.7rem;
        border-radius: 999px;
        margin-right: 0.35rem;
        margin-bottom: 0.25rem;
    }

    /* Seção de downloads */
    .download-section {
        background: linear-gradient(135deg, #f0fdfa 0%, #ecfeff 100%);
        border-radius: 16px;
        padding: 1.5rem;
        border: 1px solid #ccfbf1;
        margin-top: 0.5rem;
    }
    .download-title {
        font-weight: 600;
        color: #0f766e;
        margin-bottom: 1rem;
        font-size: 1.05rem;
    }

    /* Caixas vazias (o <div> do card fecha sozinho no Streamlit) */
    .card { display: none; }

    /* Texto sempre escuro, mesmo se o navegador estiver em modo escuro */
    .stApp, .stApp p, .stApp span, .stApp label, .stApp summary,
    [data-testid="stExpander"] summary, [data-testid="stExpander"] p,
    .stDateInput input, .stMultiSelect input, .stNumberInput input { color: #1f2937 !important; }
    [data-testid="stExpander"] { background: #ffffff !important; border: 1px solid #fce7f3 !important; border-radius: 12px !important; }
    .stButton > button:not([kind="primary"]) { background: #ffffff !important; color: #475569 !important; }
    .hero, .hero h1, .hero p { color: #ffffff !important; }

    /* Esconde menu hamburger e footer padrão do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HERO
# ─────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>Relatório Comercial</h1>
    <p>Léah Estética · Metas, categorias, vendedoras e comissão por unidade, com dados atualizados toda madrugada</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TOKENS / VALIDAÇÃO
# ─────────────────────────────────────────────
tokens = carregar_tokens()
disponiveis = [u for u in UNIDADES if slug(u) in tokens]
faltando = [u for u in UNIDADES if slug(u) not in tokens]

if not disponiveis:
    st.error("Nenhum token encontrado. Crie o **key.ini** com a seção `[tokens]`.")
    st.stop()

if faltando:
    st.warning(f"Sem token configurado para: **{', '.join(faltando)}**")

# ─────────────────────────────────────────────
# CONTROLES
# ─────────────────────────────────────────────
hoje = hoje_br()

st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">⚙️ Configuração do relatório</div>', unsafe_allow_html=True)

col1, col2 = st.columns([1.1, 1.4], gap="large")

with col1:
    periodo = st.date_input(
        "Período",
        value=(hoje.replace(day=1), hoje),
        format="DD/MM/YYYY",
        help="Selecione a data inicial e a data final do relatório",
    )

with col2:
    escolhidas = st.multiselect(
        "Clínicas",
        options=disponiveis,
        default=disponiveis,
        help="Selecione uma ou mais unidades para gerar o relatório",
    )

st.markdown("</div>", unsafe_allow_html=True)

if len(periodo) != 2:
    st.info("Selecione a data inicial e a data final do período.")
    st.stop()

ini, fim = periodo


# ─────────────────────────────────────────────
# HELPERS — divisão por mês
# ─────────────────────────────────────────────
def meses_no_periodo(ini: dt.date, fim: dt.date):
    """Retorna lista de (início, fim) para cada mês calendário dentro do período."""
    if ini > fim:
        return []
    fatias = []
    cur = ini
    while cur <= fim:
        # último dia do mês atual
        if cur.month == 12:
            fim_mes = cur.replace(day=31)
        else:
            fim_mes = (cur.replace(month=cur.month + 1, day=1) - dt.timedelta(days=1))
        fatia_fim = min(fim_mes, fim)
        fatias.append((cur, fatia_fim))
        # próximo mês
        cur = fatia_fim + dt.timedelta(days=1)
    return fatias


fatias = meses_no_periodo(ini, fim)
multi_mes = len(fatias) > 1

if multi_mes:
    nomes_meses = {
        1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
    }
    resumo = " · ".join(
        f"{nomes_meses[a.month]}/{a.year}" for a, _ in fatias
    )
    st.info(f"📅 Período atravessa **{len(fatias)} meses** → será gerado **1 PDF por mês** ({resumo})")


# ─────────────────────────────────────────────
# CACHE (dados baixados toda madrugada pelo precache.py) + GERAÇÃO
# ─────────────────────────────────────────────
def _versao(cidade):
    """Muda quando o cache noturno é atualizado -> invalida o cache do Streamlit."""
    p = caminho_cache(cidade)
    return p.stat().st_mtime if p.exists() else 0


@st.cache_data(ttl=86400, show_spinner=False, max_entries=8)
def dados_unidade(cidade, versao):
    return carregar_cache(cidade)


@st.cache_data(ttl=86400, show_spinner=False, max_entries=400)
def gerar(cidade, ini, fim, versao, metas, dias_rest, dia, metas_vend=()):
    """PDF da unidade. Usa o cache diário; só vai à API se o período não estiver nele.
    `metas`, `dias_rest` e `dia` (data de hoje) entram na chave: mudou a meta ou virou o dia, refaz."""
    dados = dados_unidade(cidade, versao)
    mv = {n: {"meta": m, "Corporal": pc, "Facial": pf, "Depilação": pd_}
          for n, m, pc, pf, pd_ in metas_vend}
    if cache_cobre(dados, ini, fim):
        pdf = gerar_relatorio_de_dados(cidade, dados, ini, fim, metas=metas, dias_restantes=dias_rest,
                                       metas_vend=mv)
        origem = "cache"
    else:
        pdf = gerar_relatorio_unidade(cidade, tokens[slug(cidade)], ini, fim,
                                      metas=metas, dias_restantes=dias_rest, metas_vend=mv)
        origem = "API (período fora do cache)"
    return pdf, origem, DIAGNOSTICO.get(cidade)


# ─────────────────────────────────────────────
# METAS DO MÊS (3 faixas por clínica)
# ─────────────────────────────────────────────
metas_ui, dias_ui = {}, {}
with st.expander("🎯 Metas do mês (Inicial · TOP · Turbo)", expanded=False):
    st.caption(
        f"O PDF compara o faturamento do mês até hoje ({hoje:%d/%m/%Y}) com cada meta e calcula quanto "
        "precisa vender por dia nos dias que restam. Em meses fechados mostra só o resultado. "
        "As metas valem para todos os meses gerados. Deixe **0** para não mostrar metas da clínica."
    )
    for cidade in escolhidas:
        padrao = rel.metas_padrao_da_unidade(cidade) or (0.0, 0.0, 0.0)
        st.markdown(f"**{cidade}**")
        cols = st.columns([1, 1, 1, 1])
        valores = []
        for col, nome_faixa, v0 in zip(cols[:3], rel.NOMES_FAIXAS, padrao):
            with col:
                valores.append(st.number_input(
                    nome_faixa, min_value=0.0, value=float(v0), step=1000.0,
                    format="%.2f", key=f"meta_{slug(cidade)}_{nome_faixa}"))
        with cols[3]:
            dias_ui[cidade] = st.number_input(
                "Dias restantes", min_value=0, value=0, step=1,
                key=f"dias_{slug(cidade)}",
                help="0 = automático (seg–sáb, sem feriados nacionais, contando hoje). "
                     "Preencha só se a clínica tiver folga/feriado local.")
        metas_ui[cidade] = tuple(valores)

# ─────────────────────────────────────────────
# METAS DAS VENDEDORAS (comissão)
# ─────────────────────────────────────────────
metas_vend_ui = {}
with st.expander("👩‍💼 Metas das vendedoras (regra da comissão)", expanded=False):
    st.caption(
        "Informe a **meta mensal total** de cada vendedora e como ela se divide entre as categorias "
        "(os 3 percentuais devem somar 100). A meta de cada categoria vale 100%: abaixo de 90% não "
        "comissiona; de 90% a 99,99% paga de 0 a 50% da comissão; a partir de 100% paga inteira. "
        "Meta **0** = sem meta: a vendedora recebe a comissão cheia, sem a regra. "
        "Valores digitados valem só nesta sessão; para deixar fixo, use o arquivo metas_vendedoras.json."
    )
    for cidade in escolhidas:
        salvas = rel.carregar_metas_vendedoras(cidade)
        nomes = vendedoras_do_periodo(dados_unidade(cidade, _versao(cidade)), ini, fim)
        if not nomes:
            continue
        linhas = []
        for n in nomes:
            cfg = rel.buscar_meta_vendedora(n, salvas) or {}
            pct = {g: cfg.get(g, rel.PCT_META_PADRAO[g]) for g in rel.GRUPOS}
            linhas.append({"Vendedora": n, "Meta total (R$)": float(cfg.get("meta", 0.0)),
                           "% Corporal": pct["Corporal"] * 100, "% Facial": pct["Facial"] * 100,
                           "% Depilação": pct["Depilação"] * 100})
        st.markdown(f"**{cidade}**")
        edit = st.data_editor(
            pd.DataFrame(linhas), hide_index=True, use_container_width=True, num_rows="fixed",
            disabled=["Vendedora"], key=f"mv_{slug(cidade)}",
            column_config={
                "Meta total (R$)": st.column_config.NumberColumn(min_value=0.0, step=1000.0, format="%.2f"),
                "% Corporal": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=5.0),
                "% Facial": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=5.0),
                "% Depilação": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=5.0),
            })
        for _, r in edit.iterrows():
            soma = r["% Corporal"] + r["% Facial"] + r["% Depilação"]
            if r["Meta total (R$)"] > 0 and abs(soma - 100) > 0.01:
                st.warning(f"{r['Vendedora']}: os percentuais somam {soma:.0f}%, não 100%.")
        metas_vend_ui[cidade] = tuple(
            (r["Vendedora"], float(r["Meta total (R$)"]), r["% Corporal"] / 100,
             r["% Facial"] / 100, r["% Depilação"] / 100) for _, r in edit.iterrows())

# Situação do cache de cada clínica
with st.expander("🗄️ Situação dos dados", expanded=False):
    for cidade in escolhidas:
        d = dados_unidade(cidade, _versao(cidade))
        if d:
            a, b = d["cobertura"]
            st.write(f"**{cidade}** — atualizado em {d['atualizado_em']} · cobre {a[8:]}/{a[5:7]}/{a[:4]} a {b[8:]}/{b[5:7]}/{b[:4]}")
        else:
            st.write(f"**{cidade}** — sem cache: será baixado ao gerar (a 1ª vez leva alguns minutos; depois fica rápido)")

# Botões de ação
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">🚀 Ações</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns([2, 1.2, 2])

with c1:
    gerar_btn = st.button(
        "Gerar relatórios",
        type="primary",
        disabled=not escolhidas,
        use_container_width=True,
    )

with c2:
    if st.button("Recarregar", use_container_width=True,
                 help="Descarta os PDFs já montados nesta sessão e lê o cache de novo"):
        gerar.clear()
        dados_unidade.clear()
        for k in ("pdfs", "erros", "diag"):
            st.session_state.pop(k, None)
        st.rerun()

with c3:
    if escolhidas:
        badges = " ".join(f'<span class="badge">{c}</span>' for c in escolhidas)
        st.markdown(
            f'<div style="padding-top:0.55rem;">{badges}</div>',
            unsafe_allow_html=True,
        )

st.markdown("</div>", unsafe_allow_html=True)

# Geração
if gerar_btn:
    pdfs, erros, diag = {}, {}, {}

    # 1) Garante dados de hoje: baixa da API só o que está vazio ou de antes da meia-noite
    pendentes = [c for c in escolhidas
                 if not cache_do_dia(dados_unidade(c, _versao(c))) and slug(c) in tokens]
    if pendentes:
        with st.status(f"Baixando dados de hoje: {', '.join(pendentes)}…", expanded=True) as stt:
            with ThreadPoolExecutor(max_workers=len(pendentes)) as ex:
                futuros = {c: ex.submit(atualizar_cache, c, tokens[slug(c)]) for c in pendentes}
                for c, fut in futuros.items():
                    try:
                        fut.result()
                        stt.write(f"✓ {c}")
                    except Exception as e:
                        erros[f"{c} · atualização dos dados"] = str(e)
                        stt.write(f"✗ {c}: {e}")
            stt.update(label="Dados atualizados", state="complete")
        dados_unidade.clear()
    total = len(escolhidas) * len(fatias)
    barra = st.progress(0.0, text="Preparando…")
    feito = 0
    for cidade in escolhidas:
        resultados = []
        for fatia_ini, fatia_fim in fatias:
            label = f"{fatia_ini:%m/%Y}"
            barra.progress(feito / total, text=f"{cidade} · {label}…")
            try:
                pdf_bytes, origem, dg = gerar(
                    cidade, fatia_ini, fatia_fim, _versao(cidade),
                    metas_ui.get(cidade), dias_ui.get(cidade) or None, hoje.isoformat(),
                    metas_vend_ui.get(cidade, ()))
                resultados.append((label, fatia_ini, fatia_fim, pdf_bytes))
                if dg:
                    diag[cidade] = dg
            except Exception as e:
                erros[f"{cidade} · {label}"] = str(e)
            feito += 1
        if resultados:
            pdfs[cidade] = resultados
    barra.empty()
    st.session_state["pdfs"] = (pdfs, ini, fim)
    st.session_state["erros"] = erros
    st.session_state["diag"] = diag
    st.rerun()

for onde, msg in st.session_state.get("erros", {}).items():
    st.error(f"**{onde}** falhou: {msg}")

# ─────────────────────────────────────────────
# DOWNLOADS
# ─────────────────────────────────────────────
if "pdfs" in st.session_state:
    pdfs, i, f = st.session_state["pdfs"]

    if pdfs:
        st.markdown('<div class="download-section">', unsafe_allow_html=True)
        st.markdown(
            f'<div class="download-title">📥 Relatórios prontos · {i:%d/%m/%Y} → {f:%d/%m/%Y}</div>',
            unsafe_allow_html=True,
        )

        for cidade, lista in pdfs.items():
            if len(lista) > 1:
                st.markdown(f"**{cidade}**")
                cols = st.columns(min(len(lista), 4))
                for idx, (label, fi, ff, pdf) in enumerate(lista):
                    with cols[idx % 4]:
                        st.download_button(
                            label=f"⬇️ {label}",
                            data=pdf,
                            file_name=f"{cidade} - {fi:%d-%m-%y} - {ff:%d-%m-%y}.pdf",
                            mime="application/pdf",
                            key=f"dl_{cidade}_{fi:%Y%m}",
                            use_container_width=True,
                        )
            else:
                label, fi, ff, pdf = lista[0]
                st.download_button(
                    label=f"⬇️  {cidade}",
                    data=pdf,
                    file_name=f"{cidade} - {fi:%d-%m-%y} - {ff:%d-%m-%y}.pdf",
                    mime="application/pdf",
                    key=f"dl_{cidade}",
                    use_container_width=True,
                )

        st.markdown("</div>", unsafe_allow_html=True)
    elif not st.session_state.get("erros"):
        st.warning("Nenhum relatório foi gerado.")

# ─────────────────────────────────────────────
# DIAGNÓSTICO (colapsado por padrão)
# ─────────────────────────────────────────────
with st.expander("🔧 Diagnóstico técnico (itens fora de Facial/Corporal/Depilação)", expanded=False):
    diag = st.session_state.get("diag")
    if not diag:
        st.caption("Gere um relatório para ver o diagnóstico.")
    else:
        for cid, d in diag.items():
            st.markdown(f"##### {cid}")
            st.caption(f"Cache atualizado em {d.get('cache_atualizado_em')} · {d['n_servicos']} serviços cadastrados")
            st.write("**Vendas com serviços lidos do plano:**", d.get("vendas_com_servicos_do_plano"))
            st.write("**Itens sem categoria:**", d["sem_categoria"] or "nenhum ✓")
            st.write("**Nomes dos itens vendidos:**", d["nomes_itens"])
            st.divider()

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
    Léah Estética · Relatório Comercial · Dados via Belle Software API
</div>
""", unsafe_allow_html=True)
