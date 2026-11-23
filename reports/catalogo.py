"""catalogo.py — todos os relatórios online. Cada um: Relatorio(key, titulo, grupo, desc, params, periodo, run)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import timedelta
from typing import Callable

import pandas as pd

from belle_core import BelleClient
from reports.base import COD_ESTAB, CampoAusente, Ctx, P, Resultado, data, norm, num, to_df, whatsapp

J = 89  # janela máxima por consulta


@dataclass
class Relatorio:
    key: str
    titulo: str
    grupo: str
    desc: str
    run: Callable
    params: list = field(default_factory=list)
    periodo: tuple | None = (30, 0)   # (dias para trás, dias para frente) do período padrão; None = sem período
    pacote: bool = True               # entra no "pacote completo"


REGISTRO: dict[str, Relatorio] = {}


def reg(r: Relatorio):
    REGISTRO[r.key] = r
    return r


def _concat(partes: dict, nome="Unidade") -> pd.DataFrame:
    dfs = []
    for u, df in partes.items():
        if df is not None and not df.empty:
            d = df.copy()
            d.insert(0, nome, u)
            dfs.append(d)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# ════════════════════════════════════════════════════════════
# 1) CLIENTES SEM ATENDIMENTO HÁ +N DIAS
# ════════════════════════════════════════════════════════════
def run_inativos(ctx: Ctx, ini, fim, p):
    dias_min, lookback = int(p["dias"]), int(p["lookback"])
    hoje = ctx.hoje

    def por_unidade(u):
        reg_ = ctx.get_janelas(u, "relatorios/relatorio_atendimentos", hoje - timedelta(days=lookback), hoje, J,
                               extra={"statusAgendamento": "Atendido", "codEstab": COD_ESTAB})
        df = to_df(reg_)
        if df.empty:
            return df
        c_id = ctx.col(df, "a_id", ["idcliente", "codcliente", "codigocliente"], None, False)
        c_nome = ctx.col(df, "a_nome", ["nomecliente", "cliente", "nomedocliente"], "cliente")
        c_data = ctx.col(df, "a_data", ["dataagendamento", "dtagendamento", "dataatendimento", "data"], "data")
        c_serv = ctx.col(df, "a_serv", ["nomeservico", "servico", "nomedoservico"], "servico", False)
        c_prof = ctx.col(df, "a_prof", ["nomeprofissional", "profissional"], "profission", False)
        c_cel = ctx.col(df, "a_cel", ["celular", "telefone"], "celular", False)
        d = pd.DataFrame({
            "id": df[c_id].astype(str).str.strip() if c_id else "",
            "nome": df[c_nome].astype(str).str.strip(),
            "dt": data(df[c_data]), "servico": df[c_serv] if c_serv else "",
            "prof": df[c_prof] if c_prof else "", "cel": df[c_cel] if c_cel else ""}).dropna(subset=["dt"])
        d["chave"] = d["id"].where(d["id"].ne("") & d["id"].ne("nan"), d["nome"].str.upper())
        d = d.sort_values("dt")
        ult = d.groupby("chave").tail(1).set_index("chave")
        ult["qtd"] = d.groupby("chave").size()
        ult = ult.reset_index()
        ult["dias"] = (pd.Timestamp(hoje) - ult["dt"].dt.normalize()).dt.days
        ult = ult[ult["dias"] >= dias_min].copy()
        # agendados no futuro: sai da lista
        if p["excluir_futuros"]:
            fut = to_df(ctx.get_janelas(u, "agendamentos", hoje, hoje + timedelta(days=60), J,
                                        extra={"codEstab": COD_ESTAB}))
            if not fut.empty:
                f_id = ctx.col(fut, "f_id", ["idcliente", "codcliente", "codigocliente"], None, False)
                f_nome = ctx.col(fut, "f_nome", ["nomecliente", "cliente"], "cliente", False)
                fixos = set(fut[f_id].astype(str)) if f_id else set()
                fnomes = set(fut[f_nome].astype(str).str.upper().str.strip()) if f_nome else set()
                ult = ult[~(ult["id"].isin(fixos) | ult["nome"].str.upper().isin(fnomes))]
        # celular: do relatório, senão do cadastro
        if ult["cel"].astype(str).str.strip().replace({"nan": "", "None": ""}).eq("").mean() > 0.5:
            cad = ctx.clientes(u, p["usar_base"])
            if not cad.empty:
                por_id = dict(zip(cad["id"], cad["celular"]))
                por_nome = dict(zip(cad["nome"].str.upper(), cad["celular"]))
                ult["cel"] = [por_id.get(i) or por_nome.get(n.upper()) or c
                              for i, n, c in zip(ult["id"], ult["nome"], ult["cel"])]
        out = pd.DataFrame({
            "Cliente": ult["nome"], "Celular": ult["cel"].astype(str).replace({"nan": "", "None": ""}),
            "Último atendimento": ult["dt"].dt.strftime("%d/%m/%Y"), "Dias sem atendimento": ult["dias"],
            "Último serviço": ult["servico"], "Profissional": ult["prof"], "Atendimentos no período": ult["qtd"]})
        out["WhatsApp"] = out["Celular"].map(whatsapp)
        if p["so_celular"]:
            out = out[out["Celular"].str.strip() != ""]
        return out.sort_values("Dias sem atendimento", ascending=False)

    partes = ctx.paralelo(por_unidade)
    tudo = _concat(partes)
    sheets = {"Todos": tudo}
    for u, df in partes.items():
        if df is not None and not df.empty:
            sheets[u] = df
    kpis = [("Clientes", len(tudo))]
    if not tudo.empty:
        faixas = pd.cut(tudo["Dias sem atendimento"], [dias_min - 1, 60, 90, 180, 10_000],
                        labels=[f"{dias_min}-60 d", "61-90 d", "91-180 d", "180+ d"])
        res = tudo.assign(Faixa=faixas).groupby(["Unidade", "Faixa"], observed=True).size().unstack(fill_value=0)
        sheets = {"Resumo": res.reset_index(), **sheets}
        kpis.append(("Com celular", int((tudo["Celular"].str.strip() != "").sum())))
    return Resultado(sheets, kpis, [f"Sem atendimento há {dias_min}+ dias (olhando {lookback} dias de histórico). "
                                    "Quem nunca atendeu nesse histórico não aparece."])


reg(Relatorio("inativos", "Clientes sem atendimento há +N dias (celular e serviço)", "Clientes e retenção",
              "Lista de reativação: último atendimento, último serviço, celular e link de WhatsApp.",
              run_inativos, [P("dias", "Sem atendimento há (dias) ou mais", "num", 30),
                             P("lookback", "Histórico analisado (dias)", "num", 180),
                             P("excluir_futuros", "Tirar quem já tem agendamento futuro", "bool", True),
                             P("so_celular", "Só quem tem celular", "bool", False),
                             P("usar_base", "Usar base de clientes salva (rápido)", "bool", True,
                               help="Desmarque para recarregar o cadastro da API (lento, vários minutos).")],
              periodo=None))


# ════════════════════════════════════════════════════════════
# 2) ANIVERSARIANTES
# ════════════════════════════════════════════════════════════
def run_aniversarios(ctx: Ctx, ini, fim, p):
    mes = int(p["mes"])

    def por_unidade(u):
        cad = ctx.clientes(u, p["usar_base"])
        if cad.empty:
            return cad
        nasc = data(cad["nascimento"])
        d = cad.assign(_n=nasc).dropna(subset=["_n"])
        d = d[d["_n"].dt.month == mes].copy()
        d["Dia"] = d["_n"].dt.day
        d["Idade que faz"] = ctx.hoje.year - d["_n"].dt.year
        d = d.sort_values("Dia")
        out = pd.DataFrame({"Dia": d["Dia"], "Cliente": d["nome"], "Nascimento": d["_n"].dt.strftime("%d/%m/%Y"),
                            "Idade que faz": d["Idade que faz"], "Celular": d["celular"], "Cidade": d["cidade"]})
        out["WhatsApp"] = out["Celular"].map(whatsapp)
        return out

    partes = ctx.paralelo(por_unidade)
    tudo = _concat(partes)
    sheets = {"Todos": tudo, **{u: d for u, d in partes.items() if d is not None and not d.empty}}
    return Resultado(sheets, [("Aniversariantes", len(tudo))],
                     ["Usa o cadastro de clientes de cada unidade (o campo data de nascimento precisa estar preenchido)."])


reg(Relatorio("aniversarios", "Aniversariantes do mês", "Clientes e retenção",
              "Quem faz aniversário no mês, com celular e WhatsApp.", run_aniversarios,
              [P("mes", "Mês", "select", 0, [str(i) for i in range(1, 13)]),
               P("usar_base", "Usar base de clientes salva (rápido)", "bool", True)], periodo=None))


# ════════════════════════════════════════════════════════════
# 3) PLANOS: SALDO, VENCENDO, SEM AGENDAMENTO
# ════════════════════════════════════════════════════════════
def run_planos(ctx: Ctx, ini, fim, p):
    venc_dias, meses = int(p["vence_em"]), int(p["meses"])

    def por_unidade(u):
        a = ctx.hoje - timedelta(days=30 * meses)
        df = to_df(ctx.get_janelas(u, "relatorios/sessoes_planos", a, ctx.hoje, J, "dtVndIni", "dtVndFim",
                                   {"codEstab": COD_ESTAB}))
        if df.empty:
            return df
        m = {k: ctx.col(df, k, c, f, o) for k, c, f, o in [
            ("s_id", ["codcliente", "codigocliente", "idcliente", "codigodocliente"], None, False),
            ("s_nome", ["nomecliente", "cliente", "codigodocliente"], "cliente", True),
            ("s_serv", ["nomedoservico", "nomeservico", "servico"], "servico", False),
            ("s_plano", ["nomedoplano", "nomeplano", "plano"], None, False),
            ("s_venda", ["datadavenda", "datavenda", "dtvenda"], "venda", False),
            ("s_valid", ["validade", "datavalidade", "dtvalidade", "vencimento"], "valid", False),
            ("s_rest", ["sessoesrestantes", "sessaorestante", "restantes"], "restan", True),
            ("s_agend", ["sessoesagendadas", "agendadas"], "agend", False)]}
        d = pd.DataFrame({
            "Cliente": df[m["s_nome"]].astype(str).str.replace(r"^\d+\s*[-–]\s*", "", regex=True).str.strip(),
            "Plano": df[m["s_plano"]] if m["s_plano"] else "", "Serviço": df[m["s_serv"]] if m["s_serv"] else "",
            "Venda": data(df[m["s_venda"]]) if m["s_venda"] else pd.NaT,
            "Validade": data(df[m["s_valid"]]) if m["s_valid"] else pd.NaT,
            "Restantes": df[m["s_rest"]].map(num), "Agendadas": df[m["s_agend"]].map(num) if m["s_agend"] else 0.0})
        d["_id"] = df[m["s_id"]].astype(str).str.replace(r"\D.*$", "", regex=True) if m["s_id"] else ""
        return d[d["Restantes"] > 0]

    partes = ctx.paralelo(por_unidade)
    base = _concat(partes)
    if base.empty:
        return Resultado({"Planos": base}, [("Linhas", 0)], ["Nenhum plano com saldo."])
    base["Dias p/ vencer"] = (base["Validade"] - pd.Timestamp(ctx.hoje)).dt.days
    base["Sem agendar"] = (base["Restantes"] - base["Agendadas"]).clip(lower=0)
    fmt = base.assign(Venda=base["Venda"].dt.strftime("%d/%m/%Y"), Validade=base["Validade"].dt.strftime("%d/%m/%Y"))
    cols = ["Unidade", "Cliente", "Plano", "Serviço", "Venda", "Validade", "Dias p/ vencer", "Restantes", "Agendadas", "Sem agendar"]
    por_cli = (base.groupby(["Unidade", "Cliente"], as_index=False)
               .agg(Saldo=("Restantes", "sum"), Agendadas=("Agendadas", "sum"), SemAgendar=("Sem agendar", "sum"),
                    Vence=("Validade", "min")))
    por_cli["Vence"] = por_cli["Vence"].dt.strftime("%d/%m/%Y")
    por_cli = por_cli.rename(columns={"SemAgendar": "Sem agendar", "Vence": "Validade mais próxima"}).sort_values(
        "Sem agendar", ascending=False)
    sheets = {
        "Por cliente": por_cli,
        f"Vencem em {venc_dias} dias": fmt[(base["Dias p/ vencer"] <= venc_dias) & (base["Dias p/ vencer"] >= 0)][cols]
        .sort_values("Dias p/ vencer"),
        "Já vencidos com saldo": fmt[base["Dias p/ vencer"] < 0][cols].sort_values("Dias p/ vencer"),
        "Sem sessão agendada": fmt[base["Agendadas"] == 0][cols],
        "Todos os saldos": fmt[cols]}
    kpis = [("Clientes com saldo", por_cli["Cliente"].nunique()), ("Sessões a realizar", int(base["Restantes"].sum())),
            (f"Vencem em {venc_dias} d", int(((base["Dias p/ vencer"] <= venc_dias) & (base["Dias p/ vencer"] >= 0)).sum()))]
    return Resultado(sheets, kpis, ["Baseado em Sessões de Planos (vendas dos últimos meses informados)."])


reg(Relatorio("planos", "Planos: saldo, vencendo e sem agendamento", "Clientes e retenção",
              "Quem tem sessões a usar, o que vence logo e quem não agendou.", run_planos,
              [P("vence_em", "Vencem em (dias)", "num", 60), P("meses", "Vendas dos últimos (meses)", "num", 24)],
              periodo=None))


# ════════════════════════════════════════════════════════════
# 4) VENDAS (demonstrativo detalhado — campos conhecidos)
# ════════════════════════════════════════════════════════════
def _demonstrativo(ctx: Ctx, u, ini, fim):
    return ctx.get_janelas(u, "relatorios/demonstrativo_vendas_detalhado", ini, fim, J, "dtInicio", "dtFim",
                           {"codEstab": COD_ESTAB})


def _linhas_venda(ctx, ini, fim):
    partes = ctx.paralelo(lambda u: _demonstrativo(ctx, u, ini, fim))
    rows = []
    for u, regs in partes.items():
        for r in regs:
            pg = r.get("parcelas") or []
            rows.append({"Unidade": u, "Data": r.get("dataLancamento"), "idVenda": r.get("idVenda"),
                         "codCliente": r.get("codCliente"), "Cliente": r.get("nomeCliente"),
                         "Vendedora": r.get("responsavel"), "Item": r.get("nomeItem"),
                         "Bruto": num(r.get("valorBruto")), "Líquido": num(r.get("valorLiquido")),
                         "_parcelas": [(x.get("formaPagamento"), num(x.get("valor"))) for x in pg]})
    df = pd.DataFrame(rows)
    if not df.empty:
        df["_dt"] = data(df["Data"])
    return df


def run_boleto(ctx, ini, fim, p):
    palavras = [x.strip().lower() for x in p["palavras"].split(",") if x.strip()]
    df = _linhas_venda(ctx, ini, fim)
    if df.empty:
        return Resultado({"Boletos": df}, [("Registros", 0)], ["Sem vendas no período."])
    df["Valor boleto"] = df["_parcelas"].map(lambda ps: sum(v for f, v in ps if "boleto" in str(f).lower()))
    b = df[df["Valor boleto"] > 0].copy()
    b["Injetável"] = b["Item"].astype(str).str.lower().map(lambda t: any(w in t for w in palavras))
    cols = ["Unidade", "Data", "Cliente", "Item", "Valor boleto"]
    inj, todos = b[b["Injetável"]][cols], b[cols]
    resumo = b.groupby("Unidade").agg(Registros=("Item", "size"), Valor=("Valor boleto", "sum"),
                                      Injetáveis=("Injetável", "sum")).reset_index()
    kpis = [("Registros c/ boleto", len(b)), ("Valor total", f"R$ {b['Valor boleto'].sum():,.2f}"),
            ("Injetáveis", len(inj)), ("Valor injetáveis", f"R$ {inj['Valor boleto'].sum():,.2f}")]
    return Resultado({"Injetáveis c-boleto": inj, "Todos c-boleto": todos, "Resumo": resumo}, kpis,
                     [f"Injetável = item contém: {', '.join(palavras)}."],
                     pdf=[("Clientes com boleto e procedimentos injetáveis", inj), ("Todos os clientes com boleto", todos)])


reg(Relatorio("boleto_injetaveis", "Boletos e injetáveis", "Vendas e financeiro",
              "Versão online do injetaveis.py: vendas pagas em boleto e quais são injetáveis.", run_boleto,
              [P("palavras", "Palavras de injetáveis (vírgula)", "text",
                 "toxina, botulínica, preenchimento, skinbooster, mesoterapia, peim, bioestimulador, preenchedor")]))


def run_formas(ctx, ini, fim, p):
    df = _linhas_venda(ctx, ini, fim)
    if df.empty:
        return Resultado({"Formas": df}, [("Vendas", 0)], ["Sem vendas no período."])
    pg = df.explode("_parcelas").dropna(subset=["_parcelas"])
    pg["Forma"] = pg["_parcelas"].map(lambda x: x[0])
    pg["Valor"] = pg["_parcelas"].map(lambda x: x[1])
    tot = pg.groupby(["Unidade", "Forma"], as_index=False)["Valor"].sum().sort_values(["Unidade", "Valor"], ascending=[True, False])
    dia = pg.pivot_table(index="_dt", columns="Forma", values="Valor", aggfunc="sum", fill_value=0).reset_index()
    dia["_dt"] = dia["_dt"].dt.strftime("%d/%m/%Y")
    dia = dia.rename(columns={"_dt": "Dia"})
    return Resultado({"Por forma": tot, "Por dia (rede)": dia},
                     [("Total recebido", f"R$ {pg['Valor'].sum():,.2f}")], [])


reg(Relatorio("formas_pagamento", "Formas de pagamento", "Vendas e financeiro",
              "Quanto entrou por forma de pagamento, por unidade e por dia.", run_formas))


def run_fat_diario(ctx, ini, fim, p):
    df = _linhas_venda(ctx, ini, fim)
    if df.empty:
        return Resultado({"Faturamento": df}, [("Vendas", 0)], [])
    d = df.groupby(["Unidade", "_dt"], as_index=False).agg(Líquido=("Líquido", "sum"), Itens=("Item", "size"),
                                                           Vendas=("idVenda", "nunique"))
    d["Dia"] = d["_dt"].dt.strftime("%d/%m/%Y")
    d["Dia da semana"] = d["_dt"].dt.day_name()
    por_dia = d[["Unidade", "Dia", "Dia da semana", "Líquido", "Vendas", "Itens"]]
    mes = df.assign(Mês=df["_dt"].dt.strftime("%Y-%m")).groupby(["Unidade", "Mês"], as_index=False).agg(
        Líquido=("Líquido", "sum"), Vendas=("idVenda", "nunique"))
    mes["Ticket médio"] = (mes["Líquido"] / mes["Vendas"]).round(2)
    return Resultado({"Por mês": mes, "Por dia": por_dia}, [("Faturamento líquido", f"R$ {df['Líquido'].sum():,.2f}"),
                                                             ("Vendas", int(df["idVenda"].nunique()))], [])


reg(Relatorio("faturamento", "Faturamento por dia e por mês", "Vendas e financeiro",
              "Líquido, nº de vendas e ticket médio por unidade.", run_fat_diario))


def run_vendedoras(ctx, ini, fim, p):
    df = _linhas_venda(ctx, ini, fim)
    if df.empty:
        return Resultado({"Vendedoras": df}, [("Vendas", 0)], [])
    v = df.groupby(["Unidade", "Vendedora"], as_index=False).agg(Líquido=("Líquido", "sum"), Vendas=("idVenda", "nunique"),
                                                                 Itens=("Item", "size"))
    v["Ticket médio"] = (v["Líquido"] / v["Vendas"]).round(2)
    v = v.sort_values(["Unidade", "Líquido"], ascending=[True, False])
    it = df.groupby(["Unidade", "Item"], as_index=False).agg(Líquido=("Líquido", "sum"), Qtd=("Item", "size"))
    it = it.sort_values(["Unidade", "Líquido"], ascending=[True, False]).groupby("Unidade").head(int(p["top"]))
    return Resultado({"Vendedoras": v, "Itens mais vendidos": it}, [("Líquido", f"R$ {df['Líquido'].sum():,.2f}")], [])


reg(Relatorio("vendedoras", "Vendedoras e itens mais vendidos", "Vendas e financeiro",
              "Ranking por vendedora e top itens de cada unidade.", run_vendedoras,
              [P("top", "Top itens por unidade", "num", 30)]))


def run_top_clientes(ctx, ini, fim, p):
    df = _linhas_venda(ctx, ini, fim)
    if df.empty:
        return Resultado({"Clientes": df}, [("Clientes", 0)], [])
    g = df.groupby(["Unidade", "codCliente", "Cliente"], as_index=False).agg(
        Líquido=("Líquido", "sum"), Vendas=("idVenda", "nunique"), Itens=("Item", "size"), Última=("_dt", "max"))
    g["Ticket médio"] = (g["Líquido"] / g["Vendas"]).round(2)
    g["Última"] = g["Última"].dt.strftime("%d/%m/%Y")
    g = g.sort_values(["Unidade", "Líquido"], ascending=[True, False]).groupby("Unidade").head(int(p["top"]))
    return Resultado({"Top clientes": g}, [("Clientes listados", len(g))], [])


reg(Relatorio("top_clientes", "Melhores clientes (por faturamento)", "Vendas e financeiro",
              "Quem mais comprou no período, por unidade.", run_top_clientes, [P("top", "Top por unidade", "num", 50)]))


# ════════════════════════════════════════════════════════════
# 5) RELATÓRIOS SIMPLES: um endpoint = uma tabela (+ resumo automático)
# ════════════════════════════════════════════════════════════
def simples(key, titulo, grupo, desc, endpoint, ini_p="dtInicio", fim_p="dtFim", max_dias=J, cod=("codEstab",),
            fixos=None, extras=None, agrupar=(), valor=("valor", "total", "valorliquido"), periodo=(30, 0), pacote=True,
            por_unidade=True):
    extras = extras or []

    def run(ctx: Ctx, ini, fim, p):
        def f(u):
            prm = dict(fixos or {})
            for c in cod:
                prm[c] = COD_ESTAB
            for e in extras:
                if p.get(e.key) not in (None, ""):
                    prm[e.key] = p[e.key]
            if periodo is None:
                return to_df(ctx.get(u, endpoint, prm))
            return to_df(ctx.get_janelas(u, endpoint, ini, fim, max_dias, ini_p, fim_p, prm))
        partes = ctx.paralelo(f)
        tudo = _concat(partes)
        if tudo.empty:
            return Resultado({"Dados": tudo}, [("Linhas", 0)], ["A API não devolveu registros."])
        sheets, kpis = {"Dados": tudo}, [("Linhas", len(tudo))]
        gcol = next((c for c in tudo.columns if any(a in norm(c) for a in agrupar)), None) if agrupar else None
        vcol = next((c for c in tudo.columns if norm(c) in valor), None)
        if vcol is not None:
            tudo[vcol + " (num)"] = tudo[vcol].map(num)
            kpis.append((f"Soma de {vcol}", f"{tudo[vcol + ' (num)'].sum():,.2f}"))
        if gcol is not None:
            agg = {"Linhas": (gcol, "size")}
            if vcol is not None:
                agg["Soma"] = (vcol + " (num)", "sum")
            sheets["Resumo"] = tudo.groupby(["Unidade", gcol], as_index=False).agg(**agg).sort_values(
                ["Unidade", "Linhas"], ascending=[True, False])
            sheets = {"Resumo": sheets["Resumo"], "Dados": tudo.drop(columns=[c for c in tudo.columns if c.endswith("(num)")])}
        else:
            sheets["Dados"] = tudo.drop(columns=[c for c in tudo.columns if c.endswith("(num)")])
        return Resultado(sheets, kpis, [])
    return reg(Relatorio(key, titulo, grupo, desc, run, extras, periodo, pacote))


def _sel(key, label, opts, default=""):
    return P(key, label, "select", default, [""] + opts if default == "" else opts)


G1, G2, G3, G4, G5 = ("Clientes e retenção", "Vendas e financeiro", "Agenda e atendimento", "Qualidade (NPS)",
                      "Estoque, vouchers e contas")

simples("atividade_inativos", "Atividade dos clientes — inativos (Belle)", G1,
        "Relatório nativo do Belle de clientes inativos no período.", "relatorios/atividade_clientes_inativos")
simples("atividade_ativos", "Atividade dos clientes — ativos (Belle)", G1,
        "Relatório nativo do Belle de clientes ativos no período.", "relatorios/atividade_clientes", pacote=False)
simples("fidelizacao", "Fidelização", G1, "Fidelização por profissional.", "relatorios/fidelizacao", cod=(),
        agrupar=("profissional",))
simples("novos_clientes", "Novos clientes (cadastro) e resumo comercial", G1,
        "Clientes cadastrados no período e o que já compraram.", "listagem_comercial_clientes", cod=(), periodo=(30, 0))
simples("resumo_comercial", "Resumo comercial dos clientes", G1, "Compras por cliente no período.",
        "relatorios/resumo_comercial_clientes", "dtVendaIni", "dtVendaFim", pacote=False)
simples("ranking_servicos", "Ranking de planos e serviços vendidos", G2, "Mais vendidos no período.",
        "relatorios/ranking_vendas_servicos", agrupar=("servico", "nome", "plano"))
simples("ranking_produtos", "Ranking de produtos vendidos", G2, "Produtos mais vendidos.",
        "relatorios/ranking_vendas_produtos", cod=())
simples("aquisicao_planos", "Aquisição de planos e serviços", G2, "Planos/serviços adquiridos no período.",
        "relatorios/aquisicao_plano_servico", pacote=False)
simples("rescisao_planos", "Rescisões de planos", G1, "Planos cancelados/rescindidos.",
        "relatorios/rescisao_planos", "dtVndIni", "dtVndFim", periodo=(90, 0))
simples("comissao_prof", "Comissão de profissionais (técnicas)", G2, "Comissões por profissional e serviço.",
        "relatorios/comissao_profissional_detalhado", cod=("codEstabelecimento",), agrupar=("profissional",),
        valor=("comissao", "valorcomissao", "valor"))
simples("movimentacao", "Movimentação financeira detalhada", G2, "Entradas e saídas (caixa).",
        "relatorios/movimentacao_detalhado", fixos={"origemServico": 1, "origemPlano": 1, "origemProduto": 1,
                                                   "origemCRE": 1, "origemOutrasVendas": 1},
        extras=[P("tipoPeriodo", "Período por", "select", "Lançamento", ["Vencimento", "Lançamento", "Confirmação", "Competência"]),
                P("tipoMovimento", "Movimento", "select", "Todos", ["Todos", "Entrada", "Saída"]),
                P("situacao", "Situação", "select", "Todos", ["Todos", "Confirmado", "Não Confirmado"])], pacote=False)
simples("consumo_servicos", "Consumo de serviços (detalhado)", G3, "Serviços consumidos, por data de venda ou realização.",
        "relatorios/consumo_servicos_detalhado", cod=("estab",),
        extras=[P("tipoData", "Data de", "select", "V", ["V", "R"], help="V = venda, R = realização")], pacote=False)
simples("atendimentos", "Atendimentos detalhados", G3, "Atendimentos do período (máx. 30 dias por consulta, o app divide).",
        "relatorios/atendimentos_detalhado", max_dias=30, cod=(), pacote=False)
simples("agenda_cancelados", "Agendamentos cancelados / faltas", G3, "Cancelados, desmarcados e faltas.",
        "agendamentos/cancelados", agrupar=("status", "profissional"), extras=[
            _sel("status", "Status", ["Desmarcado", "Cancelado", "Falhou"])])
simples("agenda_futura", "Agenda futura (agendamentos em aberto)", G3, "O que está marcado nos próximos dias.",
        "agendamentos", agrupar=("profissional", "servico"), periodo=(0, 14), pacote=False)
simples("nps", "NPS", G4, "Notas de satisfação (até 180 dias por consulta).", "relatorios/analise_nps", max_dias=179,
        cod=(), agrupar=("classificacao", "profissional"), periodo=(90, 0),
        extras=[_sel("tipoNps", "Tipo", ["Atendimento", "Venda", "Profissional"])])
simples("avaliacoes_convertidas", "Avaliações convertidas em venda", G4, "Conversão das avaliações em vendas.",
        "relatorios/avaliacoes_convertidas_vendas", agrupar=("profissional",))
simples("inventario", "Inventário do estoque (resumo)", G5, "Posição atual do estoque.",
        "relatorios/inventario_estoque_resumo", cod=("estab",), periodo=None)
simples("estoque_saidas", "Estoque — saídas", G5, "Movimentações de saída.", "relatorios/estoque/saidas", pacote=False)
simples("estoque_entradas", "Estoque — entradas", G5, "Movimentações de entrada.", "relatorios/estoque/entradas", pacote=False)
simples("vouchers", "Vouchers", G5, "Vouchers emitidos/usados/vencidos.", "voucher/lista", cod=("estab",), extras=[
    P("status", "Status", "select", "Todos", ["Todos", "Em Aberto", "Usado", "Vencido", "Cancelado"])], periodo=(90, 0))
simples("contas_receber", "Contas a receber", G5, "Títulos a receber.", "contas_receber", cod=("estab",), extras=[
    P("tipoData", "Data de", "select", "vencimento", ["vencimento", "lancamento", "confirmacao", "competencia"])],
        periodo=(30, 30))
simples("contas_pagar", "Contas a pagar", G5, "Títulos a pagar.", "contas_pagar", cod=(), extras=[
    P("tipoData", "Data de", "select", "vencimento", ["vencimento", "lancamento", "confirmacao", "competencia"])],
        periodo=(30, 30))

GRUPOS = ["Clientes e retenção", "Vendas e financeiro", "Agenda e atendimento", "Qualidade (NPS)", "Estoque, vouchers e contas"]
