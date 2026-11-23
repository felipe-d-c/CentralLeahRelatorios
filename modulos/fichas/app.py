"""Fichas de clientes — gera online (API Belle), guarda o cache no GitHub só ao gerar."""
import datetime as dt
import sys
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parents[2]
for p in (str(RAIZ), str(Path(__file__).resolve().parent)):
    if p not in sys.path:
        sys.path.insert(0, p)

import fichas_core as fc
import github_cache as gh
from belle_core import BelleClient, slug, tokens_dos_secrets

UNIDADES = ["Rota", "Videira", "Joaçaba", "Treze Tílias", "Campos Novos"]
INICIO_PADRAO = dt.date(2021, 1, 1)   # as vendas de plano mais antigas ainda com sessões a usar

st.markdown('<div class="hero"><h1>Fichas de clientes</h1>'
            '<p>Ficha de sessões (PDF) de cada cliente com plano ativo, em pastas por cidade da credenciada.</p></div>',
            unsafe_allow_html=True)

tokens = tokens_dos_secrets()
c1, c2 = st.columns([1, 2])
unidade = c1.selectbox("Unidade", UNIDADES, help="Token da unidade que tem os planos das fichas.")
fonte = c2.radio("Fonte dos dados", ["Gerar online (API)", "Último cache do GitHub", "Enviar planilhas"],
                 horizontal=True)
caminho_cache = f"fichas/{slug(unidade)}.json"


def baixar_da_api(progresso):
    token = tokens.get(slug(unidade))
    if not token:
        raise RuntimeError(f"Sem token da unidade {unidade} nos secrets.")
    cli = BelleClient(token)
    hoje = dt.date.today()
    brutos = []
    janelas = list(cli.janelas(inicio, hoje))
    for i, (a, b) in enumerate(janelas):
        progresso(0.5 * i / len(janelas), f"Sessões de planos {a:%m/%Y}…")
        brutos += cli.get("relatorios/sessoes_planos", {
            "codEstab": 1, "dtVndIni": cli.d(a), "dtVndFim": cli.d(b), "listarSuspensos": ""})
    linhas, diag = fc.canon_sessoes_api(brutos)
    ids = {r["cli_id"] for r in linhas if r["cli_id"]}
    cad_brutos, pag = [], 0
    while pag < 500:                       # 100 clientes por página, a 1ª é a 0
        progresso(0.5 + 0.45 * min(pag / 150, 1), f"Cadastro de clientes (página {pag + 1})…")
        lote = cli.get("clientes", {"pagina": pag, "codEstab": 1})
        if not lote:
            break
        cad_brutos += lote
        pag += 1
    clientes = fc.canon_clientes_api(cad_brutos, ids)
    return {"versao": 1, "unidade": unidade, "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "inicio_vendas": inicio.isoformat(), "sessoes": linhas, "clientes": clientes}, diag


def buscar_inadimplentes(meses, progresso=None):
    """Clientes que o Belle marca como Inadimplente nos atendimentos de [hoje - meses, hoje + 30 dias]."""
    token = tokens.get(slug(unidade))
    if not token:
        raise RuntimeError(f"Sem token da unidade {unidade} nos secrets.")
    cli = BelleClient(token)
    hoje = dt.date.today()
    brutos = []
    jan = list(cli.janelas(hoje - dt.timedelta(days=30 * int(meses)), hoje + dt.timedelta(days=30)))
    for i, (a, b) in enumerate(jan):
        if progresso:
            progresso(i / len(jan), f"Inadimplentes {a:%m/%Y}…")
        brutos += cli.get("relatorios/relatorio_atendimentos", {
            "codEstab": 1, "dtInicio": cli.d(a), "dtFim": cli.d(b), "statusInadimplencia": "Inadimplente"})
    r = fc.canon_inadimplentes_api(brutos)
    r.update({"meses": int(meses), "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y %H:%M")})
    return r


linhas = clientes = None
if fonte.startswith("Gerar"):
    inicio = st.date_input("Vendas desde", INICIO_PADRAO, format="DD/MM/YYYY",
                           help="A API só aceita janelas de 90 dias; o app consulta uma janela por vez.")
    atual = None
    try:
        atual = gh.ler(caminho_cache)
    except Exception:
        pass
    if atual:
        st.caption(f"Cache atual: {atual['atualizado_em']} · {len(atual['sessoes'])} linhas de sessões")
    if st.button("Gerar fichas", type="primary"):
        barra = st.progress(0.0, text="Consultando a API…")
        try:
            dados, diag = baixar_da_api(lambda p, m: barra.progress(p, text=m))
            try:
                barra.progress(0.95, text="Buscando inadimplentes…")
                dados["inadimplentes"] = buscar_inadimplentes(st.session_state.get("fichas_inad_meses", 12))
            except Exception as e:
                dados["inadimplentes"] = None
                st.warning(f"Fichas geradas, mas não consegui buscar os inadimplentes: {e}")
            barra.progress(0.97, text="Salvando cache no GitHub…")
            onde = gh.salvar(caminho_cache, dados, f"cache fichas {unidade} {dados['atualizado_em']}")
            barra.empty()
            st.success(f"Dados baixados e cache salvo em {onde}.")
            st.session_state["fichas_dados"] = dados
            st.session_state["fichas_diag"] = diag
        except Exception as e:
            barra.empty()
            st.error(f"Não deu certo: {e}")
elif fonte.startswith("Último"):
    if st.button("Carregar cache", type="primary"):
        try:
            dados = gh.ler(caminho_cache)
            if dados:
                st.session_state["fichas_dados"] = dados
            else:
                st.warning("Ainda não há cache dessa unidade. Use **Gerar online** uma vez.")
        except Exception as e:
            st.error(f"Falha ao ler o GitHub: {e}")
else:
    st.caption("Plano B: os mesmos dois relatórios exportados do Belle (Sessões de Planos e Clientes).")
    fs = st.file_uploader("Sessões de planos (.xlsx)", type="xlsx")
    fcli = st.file_uploader("Clientes (.xlsx)", type="xlsx")
    if fs and fcli and st.button("Ler planilhas", type="primary"):
        l, cl = fc.canon_planilhas(fs, fcli)
        st.session_state["fichas_dados"] = {"unidade": unidade, "atualizado_em": "planilhas enviadas",
                                            "sessoes": l, "clientes": cl}

dados = st.session_state.get("fichas_dados")
if dados and dados.get("unidade") == unidade:
    st.markdown("---")
    todos = fc.agrupar(dados["sessoes"], dados["clientes"])
    o1, o2 = st.columns([1, 2])
    incluir = o1.radio("Incluir inadimplentes?", ["Sim", "Não"], horizontal=True, key="fichas_inad",
                       help="Não = gera só as fichas de clientes adimplentes (usa o status de inadimplência do Belle).")
    o2.number_input("Meses de atendimentos para detectar inadimplência", 1, 36, 12, key="fichas_inad_meses",
                    help="Cliente inadimplente sem nenhum atendimento nesse período não é detectado.")
    linhas_uso, inad = dados["sessoes"], dados.get("inadimplentes")
    if incluir == "Não":
        if not inad and fonte.startswith("Enviar"):
            st.warning("Com planilhas não dá para saber quem é inadimplente. Use 'Gerar online'.")
        else:
            if not inad:
                st.warning("Este cache ainda não tem a lista de inadimplentes.")
            if st.button("Atualizar lista de inadimplentes" if inad else "Buscar inadimplentes agora"):
                try:
                    with st.spinner("Consultando inadimplentes…"):
                        dados["inadimplentes"] = inad = buscar_inadimplentes(st.session_state["fichas_inad_meses"])
                    gh.salvar(caminho_cache, dados, f"inadimplentes {unidade} {inad['atualizado_em']}")
                    st.session_state["fichas_dados"] = dados
                    st.session_state.pop("fichas_zip", None)
                    st.rerun()
                except Exception as e:
                    st.error(f"Não consegui buscar os inadimplentes: {e}")
        if inad:
            linhas_uso = fc.sem_inadimplentes(dados["sessoes"], inad)
            st.caption(f"Lista de inadimplentes de {inad['atualizado_em']} ({len(inad['ids'])} clientes, "
                       f"atendimentos dos últimos {inad['meses']} meses).")
    grupos = fc.agrupar(linhas_uso, dados["clientes"])
    if incluir == "Não" and inad:
        tiradas = len(todos) - len(grupos)
        st.info(f"{tiradas} fichas de clientes inadimplentes foram removidas.")
        if todos and tiradas / len(todos) > 0.5:
            st.warning("Mais da metade das fichas foi removida. Confira no Explorador se o filtro "
                       "`statusInadimplencia=Inadimplente` está funcionando no seu Belle.")
    m1, m2, m3 = st.columns(3)
    m1.metric("Fichas", len(grupos))
    m2.metric("Cidades", len({g["pasta"] for g in grupos}))
    m3.metric("Dados de", dados["atualizado_em"])
    cidades = sorted({g["pasta"] for g in grupos})
    sel = st.multiselect("Cidades (vazio = todas)", cidades)
    alvo = [g for g in grupos if not sel or g["pasta"] in sel]
    if st.button(f"Montar ZIP com {len(alvo)} fichas"):
        barra = st.progress(0.0, text="Gerando PDFs…")
        z, n = fc.zip_fichas(alvo, lambda p: barra.progress(p))
        barra.empty()
        st.session_state["fichas_zip"] = (z, n)
    if "fichas_zip" in st.session_state:
        z, n = st.session_state["fichas_zip"]
        st.download_button(f"⬇️ Baixar {n} fichas (ZIP, {len(z) / 1e6:.0f} MB)", z,
                           f"fichas_{slug(unidade)}_{dt.date.today():%Y%m%d}.zip", "application/zip")
    with st.expander("Ficha individual"):
        escolha = st.selectbox("Cliente", alvo, format_func=lambda g: f"{g['nome']} — {g['pasta']}")
        if escolha:
            st.download_button("⬇️ Baixar PDF", fc.gerar_pdf(escolha), f"{escolha['nome']}.pdf", "application/pdf")
    diag = st.session_state.get("fichas_diag")
    if diag:
        with st.expander("Diagnóstico da API (campos reconhecidos)"):
            st.json(diag)
