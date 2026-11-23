"""
relatorio.py — Gerador de Relatório de Vendas Leah Estética com Kommo
=======================================================================
Classifica clientes como "Novo" ou "Existente" usando histórico (rdv.xlsx).
Inclui análise de Campanha Origem (coluna K dos arquivos d.xlsx) e dados do Kommo.
"""

import calendar
import os
import re
import unicodedata
import tempfile
from datetime import date, datetime, timedelta
try:
    import kommo_module as kommo
except ImportError:
    kommo = None

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
import seaborn as sns
from fpdf import FPDF

# ─────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────

MAPEAMENTO = [
    ("v", "Videira"),
    ("j", "Joaçaba"),
    ("t", "Treze Tílias"),
    ("c", "Campos Novos"),
    ("r", "Rota"),
]

ARQ_RDV = "rdv.xlsx"

DIAS_PT = {
    0: "Segunda-feira", 1: "Terça-feira", 2: "Quarta-feira",
    3: "Quinta-feira",  4: "Sexta-feira", 5: "Sábado", 6: "Domingo"
}

COR_ROSA   = "#ed2a78"
COR_CIANO  = "#04c6c6"
COR_CINZA  = "#6b7280"
COR_FUNDO  = "#fdf2f7"

PALETA_ROSA  = [COR_ROSA, "#f06292", "#ec407a", "#d81b60", "#880e4f"]
PALETA_CIANO = [COR_CIANO, "#26d0d0", "#00acc1", "#0097a7", "#006978"]

MAPA_CATEGORIA = {
    "Depilação": None,
    "Rejuvenescimento": None,
    "Flacidez":            "Plexr",
    "Varicose":            "PEIM",
    "Remoção":             "Limpeza de Pele",
    "Maquiagem Permanente":"Micropigmentação",
    "Emagrecimento":       "Gordura/Emagrecimento",
}


# ─────────────────────────────────────────────
# GRUPOS DO RELATÓRIO (Facial / Corporal / Depilação) + COMISSÃO
# ─────────────────────────────────────────────

GRUPOS = ["Corporal", "Facial", "Depilação"]

# Comissão da vendedora sobre o valor líquido vendido em cada grupo.
COMISSAO = {"Corporal": 0.01, "Facial": 0.015, "Depilação": 0.02}

# Quando o cadastro do sistema não traz "Corporais / Faciais / Depilação", a categoria
# detalhada que o relatório já calcula é convertida por esta tabela (edite à vontade).
GRUPO_POR_CATEGORIA = {
    "Depilação Feminina": "Depilação", "Depilação Masculina": "Depilação",
    "Toxina Botulínica": "Facial", "Preenchimento": "Facial", "Peeling": "Facial",
    "Mesoterapia / Bioestimulador / Skinbooster": "Facial", "Rejuvenescimento": "Facial",
    "Limpeza de Pele": "Facial", "Microagulhamento": "Facial", "Micropigmentação": "Facial",
    "Plexr": "Facial",
    "PEIM": "Corporal", "Gordura/Emagrecimento": "Corporal", "Massagem/Drenagem": "Corporal",
    "Radiofrequência": "Corporal", "Glúteos": "Corporal", "Ozonioterapia": "Corporal",
    "Mesoterapia": "Corporal", "Termoterapia": "Corporal",
    "Ultrassom / Lipocavitação": "Corporal", "Criolipólise": "Corporal",
}


PALAVRAS_DEPILACAO = ("depila", "feminin", "masculin", "linha alba", "dedos dos", "virilha",
                      "axila", "buco", "barba")


# Regra de metas da vendedora, por categoria (a meta da categoria = 100%):
#   abaixo de 90%      -> não comissiona;
#   de 90% a 99,99%    -> de 0 até 50% da comissão, crescendo linearmente a partir de 90%;
#   100% ou mais       -> comissão inteira.
# Se preferir 50% fixo entre 90% e 99,99% (como na planilha de metas), troque RAMPA_90_99 para False.
FAIXA_INICIO = 90.0
FAIXA_MAX_FATOR = 0.5
RAMPA_90_99 = True

# Divisão padrão da meta total entre as categorias (editável na tela, por vendedora).
PCT_META_PADRAO = {"Corporal": 0.30, "Facial": 0.25, "Depilação": 0.45}

ARQ_METAS_VEND = os.path.join(os.path.dirname(os.path.abspath(__file__)), "metas_vendedoras.json")


def fator_comissao(perc):
    """Fração da comissão que vale para um atingimento de `perc` % da meta da categoria."""
    if perc >= 100.0:
        return 1.0
    if perc >= FAIXA_INICIO:
        if not RAMPA_90_99:
            return FAIXA_MAX_FATOR
        return FAIXA_MAX_FATOR * (perc - FAIXA_INICIO) / (100.0 - FAIXA_INICIO)
    return 0.0


def buscar_meta_vendedora(nome, metas_vend):
    """metas_vend = {nome: {"meta": total, "Corporal": pct, "Facial": pct, "Depilação": pct}}.
    O nome pode ser parcial ('Andrea' acha 'Andrea Caroline Oliveira do Amaral')."""
    n = normalizar_nome(nome)
    achados = [(k, v) for k, v in (metas_vend or {}).items()
               if normalizar_nome(k) and n.startswith(normalizar_nome(k))]
    if not achados or not achados[0][1].get("meta"):
        return None
    return max(achados, key=lambda kv: len(normalizar_nome(kv[0])))[1]


def comissao_vendedora(valores, qtds, cfg):
    """Detalhe por categoria. Sem meta cadastrada: comissão cheia (sem a regra de metas)."""
    out = {}
    for g in GRUPOS:
        valor = valores.get(g, 0.0)
        meta = (cfg["meta"] * cfg.get(g, 0.0)) if cfg else 0.0
        if meta > 0:
            perc = valor / meta * 100
            fator = fator_comissao(perc)
        else:
            perc, fator = None, 1.0
        out[g] = {"valor": valor, "qtd": qtds.get(g, 0), "meta": meta if meta > 0 else None,
                  "perc": perc, "fator": fator, "comissao": valor * COMISSAO[g] * fator}
    return out


def carregar_metas_vendedoras(unidade):
    """Lê metas_vendedoras.json (opcional): {"joacaba": {"Andrea": {"meta": 40000,
    "Corporal": 0.4, "Facial": 0.4, "Depilação": 0.2}}}"""
    try:
        import json
        with open(ARQ_METAS_VEND, encoding="utf-8") as f:
            return json.load(f).get(normalizar_nome(unidade), {})
    except Exception:
        return {}


def grupo_comissao(cat_sistema, nome_servico="", cat_final=""):
    """Facial / Corporal / Depilação de um item vendido, ou "Outros" se não der para saber.
    1º o cadastro do sistema (Faciais, Corporais, Depilação); 2º a tabela acima."""
    c = _sem_acento(re.sub(r"^\s*\d+\s*[-–]\s*", "", str(cat_sistema or ""))).strip()
    if c.startswith("facia"):
        return "Facial"
    if c.startswith("corpora"):
        return "Corporal"
    if c.startswith("depila"):
        return "Depilação"
    g = GRUPO_POR_CATEGORIA.get(cat_final)
    if g:
        return g
    n = _sem_acento(nome_servico)       # último recurso: palavras do nome (ex.: unidade Rota)
    if any(k in n for k in PALAVRAS_DEPILACAO):
        return "Depilação"
    return "Outros"


# ─────────────────────────────────────────────
# METAS (3 faixas) + META DIÁRIA DINÂMICA
# ─────────────────────────────────────────────

NOMES_FAIXAS = ["Meta Inicial", "Meta TOP", "Meta Turbo"]

# Metas mensais padrão por unidade (Inicial, TOP, Turbo). Na tela dá para alterar.
METAS_PADRAO = {
    "joacaba":      (100000.0, 115000.0, 135000.0),
    "camposnovos":  (85000.0, 100000.0, 120000.0),
    "videira":      (85000.0, 100000.0, 120000.0),
    "trezetilias":  (85000.0, 100000.0, 120000.0),
    # "rota": (…, …, …),   # <- preencha quando houver meta definida
}

# Dias da semana em que a clínica atende (0=segunda ... 6=domingo).
# Para voltar a contar todos os dias corridos: {0, 1, 2, 3, 4, 5, 6}
DIAS_ATENDIMENTO = {0, 1, 2, 3, 4, 5}

# Feriados que não contam como dia de atendimento: nacionais fixos + Sexta-feira Santa.
FERIADOS_FIXOS = {(1, 1), (4, 21), (5, 1), (9, 7), (10, 12), (11, 2), (11, 15), (11, 20), (12, 25)}
FERIADOS_EXTRAS = set()   # datas locais: {date(2026, 12, 24), date(2026, 12, 31)}


def metas_padrao_da_unidade(nome):
    return METAS_PADRAO.get(normalizar_nome(nome))


def parse_meta(texto):
    """Aceita '150000', '150.000', '150.000,50', 'R$ 150.000,00'."""
    t = re.sub(r"[^\d,.]", "", str(texto))
    if not t:
        return 0.0
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _feriados(ano):
    from dateutil.easter import easter
    fer = {date(ano, m, d) for m, d in FERIADOS_FIXOS}
    fer.add(easter(ano) - timedelta(days=2))          # Sexta-feira Santa
    return fer | {d for d in FERIADOS_EXTRAS if d.year == ano}


def dias_atendimento(ini, fim):
    """Quantidade de dias de atendimento em [ini, fim] (inclusive)."""
    if fim < ini:
        return 0
    fer, n, cur = {}, 0, ini
    while cur <= fim:
        if cur.year not in fer:
            fer[cur.year] = _feriados(cur.year)
        if cur.weekday() in DIAS_ATENDIMENTO and cur not in fer[cur.year]:
            n += 1
        cur += timedelta(days=1)
    return n


def calcular_metas(vendas_mes, mes_ref, hoje, metas, dias_restantes=None):
    """Acompanhamento das metas do mês `mes_ref` = (ano, mês).

    vendas_mes     : {date: faturamento do dia} (mês inteiro, não só o período do relatório)
    hoje           : data de hoje. No mês corrente, HOJE conta como dia ainda por vender;
                     o ritmo realizado usa só os dias já encerrados (até ontem).
    dias_restantes : força o nº de dias restantes (opcional; só vale no mês corrente).
    Devolve None se o mês é futuro ou não há meta válida.
    """
    ano, mes = mes_ref
    if (ano, mes) > (hoje.year, hoje.month):
        return None
    ini = date(ano, mes, 1)
    fim = date(ano, mes, calendar.monthrange(ano, mes)[1])
    aberto = (ano, mes) == (hoje.year, hoje.month)

    acumulado = sum(v for d, v in vendas_mes.items() if ini <= d <= min(fim, hoje))
    total_dias = dias_atendimento(ini, fim)
    if aberto:
        ontem = hoje - timedelta(days=1)
        fat_ate_ontem = sum(v for d, v in vendas_mes.items() if ini <= d <= ontem)
        decorridos = dias_atendimento(ini, ontem)
        restantes = dias_atendimento(hoje, fim)
        if dias_restantes is not None:
            restantes = max(0, int(dias_restantes))
    else:
        fat_ate_ontem, decorridos, restantes = acumulado, total_dias, 0

    media = fat_ate_ontem / decorridos if decorridos else 0.0
    projecao = fat_ate_ontem + media * restantes

    faixas = []
    for nome_faixa, valor in zip(NOMES_FAIXAS, metas or []):
        if not valor or valor <= 0:
            continue
        falta = max(0.0, valor - acumulado)
        nec = falta / restantes if restantes > 0 else 0.0
        if falta <= 0:
            situacao = "Atingida"
        elif not aberto:
            situacao = f"Não atingida (faltaram R$ {falta:,.2f})"
        elif restantes == 0:
            situacao = "Sem dias de atendimento restantes"
        elif decorridos == 0:
            situacao = "Início do mês: ainda sem ritmo para comparar"
        elif media >= nec:
            situacao = (f"No ritmo: precisa de R$ {nec:,.2f}/dia, "
                        f"R$ {media - nec:,.2f} abaixo do ritmo atual")
        else:
            situacao = (f"Acelerar: precisa de R$ {nec:,.2f}/dia, "
                        f"R$ {nec - media:,.2f} acima do ritmo atual")
        faixas.append({
            "nome": nome_faixa, "valor": valor, "perc": acumulado / valor * 100,
            "falta": falta, "atingida": falta <= 0,
            "diaria_original": valor / total_dias if total_dias else 0.0,
            "diaria_necessaria": nec, "situacao": situacao,
        })
    if not faixas:
        return None
    return {
        "faixas": faixas, "mes": (ano, mes), "aberto": aberto, "hoje": hoje,
        "acumulado": acumulado, "dias_total": total_dias, "dias_decorridos": decorridos,
        "dias_restantes": restantes, "media_diaria": media, "projecao": projecao,
    }

# ─────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────

def parse_br(v):
    if pd.isna(v) or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    texto = re.sub(r"[^\d,.]", "", str(v))
    if not texto:
        return 0.0
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0

def normalizar_nome(nome):
    if pd.isna(nome):
        return ""
    nome = str(nome).replace("\xa0", " ").strip().lower()
    nome = "".join(
        c for c in unicodedata.normalize("NFD", nome)
        if unicodedata.category(c) != "Mn"
    )
    nome = re.sub(r"[^a-z\s]", "", nome)
    return nome.replace(" ", "")

def carregar_excel(caminho):
    if not os.path.exists(caminho):
        return None
    try:
        return pd.read_excel(caminho)
    except Exception:
        try:
            return pd.read_excel(caminho, engine="xlrd")
        except Exception:
            return None

def carregar_historico_rdv(caminho_rdv):
    if not os.path.exists(caminho_rdv):
        print(f"  [aviso] Base de histórico não encontrada: {caminho_rdv}")
        return {}, {}

    df = pd.read_excel(caminho_rdv, header=0)
    if "ID Venda" not in df.columns:
        df = pd.read_excel(caminho_rdv, header=1)
    if "ID Venda" not in df.columns:
        print(f"  [erro] Coluna 'ID Venda' não encontrada em {caminho_rdv}")
        return {}, {}

    df["_id_venda"] = pd.to_numeric(df["ID Venda"], errors="coerce")
    df = df[df["_id_venda"].notna()].copy()

    if "Data Venda" not in df.columns:
        print(f"  [erro] Coluna 'Data Venda' não encontrada em {caminho_rdv}")
        return {}, {}

    df["_data_venda"] = pd.to_datetime(df["Data Venda"], dayfirst=True, errors="coerce")
    df["_id_cliente"] = df.get("ID Cliente", "").apply(
        lambda v: str(v).strip() if pd.notna(v) else ""
    )
    df["_nome_norm"] = df.get("Cliente", "").apply(normalizar_nome)

    df = df[df["_data_venda"].notna()]

    dict_id = (
        df[df["_id_cliente"] != ""]
        .groupby("_id_cliente")["_data_venda"]
        .min()
        .to_dict()
    )
    dict_nome = (
        df[df["_nome_norm"] != ""]
        .groupby("_nome_norm")["_data_venda"]
        .min()
        .to_dict()
    )
    return dict_id, dict_nome

def perfil_cliente_na_data(cliente_id, nome, data_venda, dict_id, dict_nome):
    if pd.isna(data_venda):
        return "Desconhecido"
    data_venda = pd.Timestamp(data_venda).normalize()

    primeira = None
    if cliente_id and cliente_id in dict_id:
        primeira = dict_id[cliente_id]
    else:
        nome_norm = normalizar_nome(nome)
        if nome_norm in dict_nome:
            primeira = dict_nome[nome_norm]

    if primeira is None or pd.isna(primeira):
        return "Novo"

    primeira = pd.Timestamp(primeira).normalize()
    if data_venda <= primeira:
        return "Novo"
    return "Existente"

def _sem_acento(txt):
    t = unicodedata.normalize("NFD", str(txt).lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")

def _subcategorizar_depilacao(nome_servico):
    t = nome_servico.lower()
    masculino = ["masculin", "barba", "peito masculino", "abdomen", "abdômen",
                 "costas", "ombro", "braco", "nuca"]
    if any(k in t for k in masculino):
        return "Depilação Masculina"
    return "Depilação Feminina"

def _subcategorizar_rejuvenescimento(nome_servico):
    t = nome_servico.lower()
    if any(k in t for k in ["toxina", "botul", "botox"]):
        return "Toxina Botulínica"
    if any(k in t for k in ["pree", "labial", "nasogenian", "olheira", "mandibul",
                            "papada", "queixo"]):
        return "Preenchimento"
    if any(k in t for k in ["meso", "bioest", "skinb", "victa"]):
        return "Mesoterapia / Bioestimulador / Skinbooster"
    if any(k in t for k in ["peeling", "peel"]):
        return "Peeling"
    return "Rejuvenescimento"

def resolver_categoria(cat_sistema, nome_servico):
    cat = str(cat_sistema).strip() if pd.notna(cat_sistema) else ""
    nome = str(nome_servico).strip() if pd.notna(nome_servico) else ""
    if "plexr" in nome.lower():
        return "Plexr"
    # Itens que o nome define sozinho (vale mesmo se a categoria do sistema vier
    # vazia, diferente ou errada, ex.: "Micropigmentação Labial" não é Preenchimento)
    nome_n = _sem_acento(nome)
    if "micropigment" in nome_n:
        return "Micropigmentação"
    if "peim" in nome_n:
        return "PEIM"
    if "limpeza de pele" in nome_n:
        return "Limpeza de Pele"
    if "microagulh" in nome_n:
        return "Microagulhamento"
    if cat not in ("Depilação", "Rejuvenescimento") and (
            cat not in MAPA_CATEGORIA or MAPA_CATEGORIA[cat] is None):
        if any(k in nome_n for k in ["toxina", "botul", "botox", "preench", "bioest",
                                     "skinb", "labial", "nasogen", "olheira", "mandibul"]):
            return _subcategorizar_rejuvenescimento(nome)
        if any(k in nome_n for k in ["barba", "bumbum", "coxa", "mento", "queixo", "peito",
                                     "rosto", "abdomen", "virilha", "axila", "buco",
                                     "perna", "braco", "costas", "nuca", "ombro",
                                     "depila"]):
            return _subcategorizar_depilacao(nome)
    if cat == "Depilação":
        return _subcategorizar_depilacao(nome)
    if cat == "Rejuvenescimento":
        return _subcategorizar_rejuvenescimento(nome)
    if cat in MAPA_CATEGORIA and MAPA_CATEGORIA[cat] is not None:
        return MAPA_CATEGORIA[cat]

    t = nome.lower()
    if any(k in t for k in ["massagem", "drenagem", "linfat"]):
        return "Massagem/Drenagem"
    if "radiofrequ" in t:
        return "Radiofrequência"
    if any(k in t for k in ["gluteo", "glúteo", "pump"]):
        return "Glúteos"
    if "ozon" in t:
        return "Ozonioterapia"
    if any(k in t for k in ["meso", "mesoterapia", "mesoterapía"]):
        return "Mesoterapia"
    if any(k in t for k in ["termo", "termoterapia"]):
        return "Termoterapia"
    if any(k in t for k in ["ultrassom", "ultra-som", "lipocavitação", "velox"]):
        return "Ultrassom / Lipocavitação"
    if "criolipólise" in t:
        return "Criolipólise"
    return "Outros"

def normalizar_forma_pgto(forma):
    f = str(forma).lower()
    if "plugbank" in f:
        return "Boleto PlugBank"
    if "boleto" in f:
        return "Boleto"
    if "pix" in f:
        return "PIX"
    if "dinheiro" in f:
        return "Dinheiro"
    if "maquininha" in f:
        return "Maquininha"
    if "débito" in f or "debito" in f:
        return "Cartão de Débito"
    if "venda link" in f and ("crédito" in f or "credito" in f):
        return "Link Crédito"
    if "crédito" in f or "credito" in f:
        return "Cartão de Crédito"
    return "Outros"

# ─────────────────────────────────────────────
# LEITURA DAS FONTES
# ─────────────────────────────────────────────

def ler_demonstrativo(caminho):
    df_raw = pd.read_excel(caminho, header=None)

    pgto_totais = {}
    pgto_quantidades = {}
    for i in range(2, min(200, len(df_raw))):
        label = str(df_raw.iloc[i, 0]).strip()
        val   = parse_br(df_raw.iloc[i, 1])
        if "Valor Líquido" in label and val > 0:
            forma_raw = label.replace(" - Valor Líquido", "").strip()
            if forma_raw in ("Plano - Aprovado", "Total Geral"):
                continue
            forma_norm = normalizar_forma_pgto(forma_raw)
            pgto_totais[forma_norm] = pgto_totais.get(forma_norm, 0.0) + val

            qtd = 0
            for col in [2, 3, 4]:
                if col < len(df_raw.columns):
                    qtd_val = df_raw.iloc[i, col]
                    if isinstance(qtd_val, (int, float)) and qtd_val > 0:
                        qtd = int(qtd_val)
                        break
                    elif isinstance(qtd_val, str):
                        num_match = re.search(r'(\d+)', qtd_val)
                        if num_match:
                            qtd = int(num_match.group(1))
                            break
            pgto_quantidades[forma_norm] = pgto_quantidades.get(forma_norm, 0) + qtd

    df = pd.read_excel(caminho, header=1)
    df["_id"] = pd.to_numeric(df["ID Venda"], errors="coerce")
    vendas = df[df["_id"].notna() & (df["_id"] < 30_000_000)].copy()

    vendas["valor_liq"]  = vendas["Valor Líquido"].apply(parse_br)
    vendas["valor_bruto"]= vendas["Valor Bruto"].apply(parse_br)
    vendas["data_venda"] = pd.to_datetime(
        vendas["Lançamento"], dayfirst=True, errors="coerce"
    )
    vendas["cliente_nome"] = (
        vendas["Cliente"]
        .astype(str)
        .str.replace(r"^\d+-", "", regex=True)
        .str.strip()
    )
    vendas["id_str"] = vendas["ID Venda"].astype(str)
    return vendas, pgto_totais, pgto_quantidades

def extrair_formas_pagamento_por_venda(caminho):
    """
    Lê o demonstrativo linha a linha e associa cada ID de Venda às suas
    formas de pagamento (pode haver mais de uma, em caso de parcelamento
    misto/split). Estrutura do arquivo: uma linha de "cabeçalho de venda"
    (com Lançamento preenchido e ID Venda numérico) seguida de 1+ linhas
    de "parcela/pagamento" (Lançamento vazio, forma de pagamento na
    coluna 'Cliente' e valor da parcela na coluna 'Origem Cliente').

    Retorna: { id_venda_str: [(forma_normalizada, valor), ...] }
    """
    df = pd.read_excel(caminho, header=1)
    formas_por_venda = {}
    venda_atual = None

    for _, row in df.iterrows():
        id_num = pd.to_numeric(row.get("ID Venda"), errors="coerce")
        lanc = row.get("Lançamento")

        if pd.notna(lanc) and pd.notna(id_num):
            # Linha de cabeçalho de uma venda
            venda_atual = str(int(id_num))
            formas_por_venda.setdefault(venda_atual, [])
        elif venda_atual is not None and pd.notna(row.get("Cliente")):
            # Linha de parcela/pagamento pertencente à venda atual
            forma_raw = str(row.get("Cliente")).strip()
            valor = parse_br(row.get("Origem Cliente"))
            forma_norm = normalizar_forma_pgto(forma_raw)
            formas_por_venda[venda_atual].append((forma_norm, valor))

    return formas_por_venda

def ler_detalhamento(caminho):
    df = pd.read_excel(caminho, header=1, skiprows=[2, 3, 4, 5])
    df["total_liq"]  = df["Total Líquido"].apply(parse_br)
    df["total_bruto"]= df["Total Bruto"].apply(parse_br)
    df["id_str"]     = df["ID Venda"].astype(str)
    df["categoria_final"] = df.apply(
        lambda r: resolver_categoria(r["Categoria Serviço"], r["Serviço/Produto"]),
        axis=1,
    )

    campanha_por_venda = {}
    if "Campanha Origem" in df.columns:
        for _, row in df.iterrows():
            vid = row["id_str"]
            camp = str(row["Campanha Origem"]).strip() if pd.notna(row["Campanha Origem"]) else "Não informado"
            if vid not in campanha_por_venda:
                campanha_por_venda[vid] = camp
    else:
        print(f"  [aviso] Coluna 'Campanha Origem' não encontrada em {caminho}")

    return df, campanha_por_venda

def _add_cat(rel, cat, valor, qtd, nome_norm, perfil, formas_valor=None):
    if cat not in rel:
        rel[cat] = {"vendas": 0, "faturado": 0.0, "novos": set(), "existentes": set(),
                    "formas": {}}
    rel[cat]["vendas"]  += qtd
    rel[cat]["faturado"] += valor
    if perfil == "Novo":
        rel[cat]["novos"].add(nome_norm)
    else:
        rel[cat]["existentes"].add(nome_norm)

    # formas_valor: lista de (forma_normalizada, valor_rateado) a somar na categoria
    if formas_valor:
        for forma, v in formas_valor:
            rel[cat]["formas"][forma] = rel[cat]["formas"].get(forma, 0.0) + v

def _formas_pagamento_da_venda(vid, valor_total_venda, formas_por_venda):
    """
    Retorna [(forma_normalizada, valor), ...] para a venda `vid`, com os
    valores já normalizados para somar `valor_total_venda` (usa a
    proporção de cada forma dentro do total pago na venda). Se não houver
    dado de forma de pagamento, cai em "Não informado" com o valor cheio.
    """
    entradas = formas_por_venda.get(vid, [])
    entradas = [(f, v) for f, v in entradas if v > 0]
    if not entradas:
        return [("Não informado", valor_total_venda)]

    soma = sum(v for _, v in entradas)
    if soma <= 0:
        return [("Não informado", valor_total_venda)]

    return [(forma, valor_total_venda * (v / soma)) for forma, v in entradas]

def nome_curto(nome, limite=26):
    """'Andrea Caroline Oliveira do Amaral' -> 'Andrea Amaral'. Cabe na coluna sem cortar no meio."""
    nome = " ".join(str(nome).split())
    if len(nome) <= limite:
        return nome
    partes = [x.strip() for x in nome.split(" - ")]
    prefixo = ""
    if partes[0].lower() == "master" and len(partes) > 1:
        prefixo, nome = "Master - ", partes[1]
    p = nome.split()
    curto = f"{p[0]} {p[-1]}" if len(p) > 1 else nome
    return (prefixo + curto)[:limite]


# Página 5: ainda sem conteúdo definido. Para usar, escreva uma função f(pdf, indicadores)
# que desenha a página e atribua aqui (ex.: PAGINA_5 = minha_funcao). Com None, o relatório
# pula direto da página 4 (vendedores) para a Análise Visual.
PAGINA_5 = None


def grupos_para_grafico(grupos):
    """Só as 3 categorias (o gráfico de barras ignora o que for zero)."""
    return {g: grupos[g] for g in GRUPOS if g in grupos}


# ─────────────────────────────────────────────
# GRÁFICOS
# ─────────────────────────────────────────────

def _estilo_base():
    plt.rcParams.update({
        "figure.facecolor": "white",
        "axes.facecolor":   "white",
        "axes.spines.top":  False,
        "axes.spines.right":False,
        "axes.grid":        True,
        "grid.color":       "#e5e7eb",
        "grid.linewidth":   0.6,
        "font.family":      "sans-serif",
        "font.size":        10,
    })

def grafico_categorias(relatorio_cat, total_geral):
    _estilo_base()
    dados = sorted(
        [(c, d["faturado"]) for c, d in relatorio_cat.items() if d["faturado"] > 0],
        key=lambda x: x[1],
    ) or [("Sem vendas", 0.0)]
    total_geral = total_geral or 1
    cats  = [c[:32] for c, _ in dados]
    vals  = [v for _, v in dados]
    percs = [v / total_geral * 100 for v in vals]

    n = len(cats)
    palette = sns.color_palette([COR_ROSA, COR_CIANO] * (n // 2 + 1))[:n]

    fig, ax = plt.subplots(figsize=(10, max(4, n * 0.55)))
    bars = ax.barh(cats, vals, color=palette, height=0.6, edgecolor="white", linewidth=0.5)

    for bar, val, perc in zip(bars, vals, percs):
        ax.text(
            val + max(vals) * 0.01, bar.get_y() + bar.get_height() / 2,
            f"R$ {val:,.0f}  ({perc:.1f}%)",
            va="center", fontsize=8.5, color=COR_CINZA,
        )

    ax.set_xlabel("Faturamento (R$)", color=COR_CINZA, fontsize=9)
    ax.set_title("Faturamento por Categoria", fontsize=13, color="#1f2937",
                 fontweight="bold", pad=12)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"R${x:,.0f}"))
    ax.tick_params(colors=COR_CINZA)
    fig.tight_layout()

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

def grafico_pagamentos(pgto_totais, total_geral):
    _estilo_base()
    dados = {k: v for k, v in pgto_totais.items() if v > 0}
    if not dados:
        return None

    labels = list(dados.keys())
    valores = list(dados.values())
    cores = (PALETA_CIANO + PALETA_ROSA)[:len(labels)]

    fig, ax = plt.subplots(figsize=(7, 7))
    wedges, texts, autotexts = ax.pie(
        valores,
        labels=None,
        autopct="%1.1f%%",
        colors=cores,
        pctdistance=0.78,
        startangle=90,
        wedgeprops={"linewidth": 1.5, "edgecolor": "white"},
    )
    for at in autotexts:
        at.set_fontsize(9)
        at.set_color("white")
        at.set_fontweight("bold")

    ax.add_artist(plt.Circle((0, 0), 0.55, fc="white"))

    legend_labels = [f"{l}  R${v:,.0f} ({v/total_geral*100:.1f}%)"
                     for l, v in zip(labels, valores)]
    patches = [mpatches.Patch(color=c, label=l) for c, l in zip(cores, legend_labels)]
    ax.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.18),
              ncol=2, fontsize=8, frameon=False)

    ax.set_title("Formas de Pagamento", fontsize=13, color="#1f2937",
                 fontweight="bold", pad=12)
    fig.tight_layout()

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

def grafico_perfil(perfil_clientes, faturamento_perfil):
    _estilo_base()
    tipos = ["Novos", "Existentes"]
    qtds  = [perfil_clientes.get("Novo", 0), perfil_clientes.get("Existente", 0)]
    fats  = [faturamento_perfil.get("Novo", 0.0), faturamento_perfil.get("Existente", 0.0)]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.5))

    for ax, vals, title, fmt in [
        (ax1, qtds, "Quantidade de Clientes", lambda v: str(int(v))),
        (ax2, fats, "Faturamento por Perfil",  lambda v: f"R$ {v:,.0f}"),
    ]:
        bars = ax.bar(tipos, vals,
                      color=[COR_ROSA, COR_CIANO],
                      width=0.45, edgecolor="white")
        for bar, v in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    v + max(vals) * 0.02,
                    fmt(v), ha="center", fontsize=10, color=COR_CINZA, fontweight="bold")
        ax.set_title(title, fontsize=12, color="#1f2937", fontweight="bold", pad=8)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"R${x:,.0f}") if "Fatur" in title
            else mticker.FuncFormatter(lambda x, _: f"{int(x)}")
        )
        ax.tick_params(colors=COR_CINZA)

    fig.tight_layout(pad=2)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

def grafico_vendedores(vendedor_stats):
    _estilo_base()
    dados = sorted(vendedor_stats.items(), key=lambda x: x[1]["faturamento"], reverse=True)
    nomes = [nome_curto(d[0], 20) for d in dados]
    fats  = [d[1]["faturamento"] for d in dados]
    qtds  = [d[1].get("comissao", d[1]["qtd"]) for d in dados]

    n = len(nomes)
    cores = (PALETA_ROSA + PALETA_CIANO)[:n]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, max(3.5, n * 0.8)))

    bars1 = ax1.barh(nomes, fats, color=cores, height=0.55, edgecolor="white")
    for bar, v in zip(bars1, fats):
        ax1.text(v + max(fats) * 0.01, bar.get_y() + bar.get_height() / 2,
                 f"R$ {v:,.0f}", va="center", fontsize=9, color=COR_CINZA)
    ax1.set_title("Faturamento por Vendedor", fontsize=12, color="#1f2937",
                  fontweight="bold", pad=8)
    ax1.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"R${x:,.0f}"))
    ax1.xaxis.set_major_locator(mticker.MaxNLocator(4))
    ax1.tick_params(colors=COR_CINZA)

    bars2 = ax2.barh(nomes, qtds, color=cores, height=0.55, edgecolor="white")
    for bar, v in zip(bars2, qtds):
        ax2.text(v + max(qtds) * 0.01, bar.get_y() + bar.get_height() / 2,
                 f"R$ {v:,.2f}", va="center", fontsize=9, color=COR_CINZA, fontweight="bold")
    ax2.set_title("Comissão por Vendedor", fontsize=12, color="#1f2937",
                  fontweight="bold", pad=8)
    ax2.tick_params(colors=COR_CINZA)

    fig.tight_layout(pad=2)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

def grafico_diario(vendas_diarias, vendas_semanais):
    _estilo_base()
    if not vendas_diarias:
        return None

    dias  = sorted(vendas_diarias.keys())
    vals  = [vendas_diarias[d] for d in dias]
    media = sum(vals) / len(vals)

    fig, ax = plt.subplots(figsize=(13, 4))
    ax.bar([str(d.strftime("%d/%m")) for d in dias], vals,
           color=COR_CIANO, edgecolor="white", width=0.65)
    ax.axhline(media, color=COR_ROSA, linestyle="--", linewidth=1.2,
               label=f"Média diária R$ {media:,.0f}")
    ax.set_title("Resultado Diário", fontsize=13, color="#1f2937",
                 fontweight="bold", pad=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"R${x:,.0f}"))
    ax.tick_params(axis="x", rotation=45, labelsize=8, colors=COR_CINZA)
    ax.tick_params(axis="y", colors=COR_CINZA)
    ax.legend(fontsize=9, frameon=False)
    fig.tight_layout()

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

def grafico_campanhas(campanha_stats, total_geral):
    if not campanha_stats:
        return None
    dados = {k: v["faturamento"] for k, v in campanha_stats.items() if v["faturamento"] > 0}
    if not dados:
        return None

    _estilo_base()
    labels = list(dados.keys())
    valores = list(dados.values())
    cores = (PALETA_ROSA + PALETA_CIANO)[:len(labels)]

    fig, ax = plt.subplots(figsize=(7, 7))
    wedges, texts, autotexts = ax.pie(
        valores,
        labels=None,
        autopct="%1.1f%%",
        colors=cores,
        pctdistance=0.78,
        startangle=90,
        wedgeprops={"linewidth": 1.5, "edgecolor": "white"},
    )
    for at in autotexts:
        at.set_fontsize(9)
        at.set_color("white")
        at.set_fontweight("bold")

    ax.add_artist(plt.Circle((0, 0), 0.55, fc="white"))

    legend_labels = [f"{l}  R${v:,.0f} ({v/total_geral*100:.1f}%)"
                     for l, v in zip(labels, valores)]
    patches = [mpatches.Patch(color=c, label=l) for c, l in zip(cores, legend_labels)]
    ax.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.18),
              ncol=2, fontsize=8, frameon=False)

    ax.set_title("Faturamento por Campanha Origem", fontsize=13, color="#1f2937",
                 fontweight="bold", pad=12)
    fig.tight_layout()

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    fig.savefig(tmp.name, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return tmp.name

# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────

FONT_DIR = "C:/Windows/Fonts/"

class LeahPDF(FPDF):
    def __init__(self, cidade):
        super().__init__()
        self.cidade = cidade.encode("latin-1", "replace").decode("latin-1")
        self.set_auto_page_break(auto=True, margin=18)
        try:
            self.add_font("LS",  "",  FONT_DIR + "LiberationSans-Regular.ttf",  uni=True)
            self.add_font("LS",  "B", FONT_DIR + "LiberationSans-Bold.ttf",     uni=True)
            self.add_font("LS",  "I", FONT_DIR + "LiberationSans-Italic.ttf",   uni=True)
            self.add_font("LS",  "BI", FONT_DIR + "LiberationSans-BoldItalic.ttf",uni=True)
        except:
            # fallback para fontes padrão
            pass

    def header(self):
        self.set_fill_color(237, 42, 120)
        self.rect(0, 0, 210, 10, "F")
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(255, 255, 255)
        self.set_xy(0, 1)
        self.cell(0, 8, f"Leah Estética  ·  {self.cidade}", align="C")
        self.set_text_color(0, 0, 0)
        self.ln(30)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(160, 160, 160)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

    def titulo_secao(self, texto):
        self.ln(4)
        self.set_fill_color(253, 242, 247)
        self.set_draw_color(237, 42, 120)
        self.set_line_width(0.6)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(180, 20, 80)
        self.cell(0, 9, f"  {texto}", border="L", fill=True, ln=True)
        self.set_line_width(0.2)
        self.set_draw_color(200, 200, 200)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def cabecalho_tabela(self, cols):
        self.set_fill_color(237, 42, 120)
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 8)
        for w, label in cols:
            self.cell(w, 7, label, border=1, align="C", fill=True)
        self.ln()
        self.set_text_color(0, 0, 0)

    def linha_tabela(self, cols, linha_par):
        if self.get_y() > 250:
            self.add_page()
        self.set_font("Helvetica", "", 8)
        if linha_par:
            self.set_fill_color(253, 242, 247)
        else:
            self.set_fill_color(255, 255, 255)
        for w, texto, align in cols:
            self.cell(w, 6.5, texto, border=1, align=align, fill=True)
        self.ln()

    def linha_tabela_com_cabecalho(self, cols, cabecalho_cols, linha_par):
        if self.get_y() > 250:
            self.add_page()
            self.ln(4)
            self.cabecalho_tabela(cabecalho_cols)
        self.set_font("Helvetica", "", 8)
        if linha_par:
            self.set_fill_color(253, 242, 247)
        else:
            self.set_fill_color(255, 255, 255)
        for w, texto, align in cols:
            self.cell(w, 6.5, texto, border=1, align=align, fill=True)
        self.ln()

    def linha_total(self, cols):
        if self.get_y() > 250:
            self.add_page()
        self.set_font("Helvetica", "B", 9)
        self.set_fill_color(240, 240, 240)
        for w, texto, align in cols:
            self.cell(w, 7, texto, border=1, align=align, fill=True)
        self.ln()

    def kpi(self, label, valor, x, y, w=58):
        self.set_xy(x, y)
        self.set_fill_color(253, 242, 247)
        self.set_draw_color(237, 42, 120)
        self.set_line_width(0.4)
        self.rect(x, y, w, 14, "DF")
        self.set_font("Helvetica", "", 7)
        self.set_text_color(120, 120, 120)
        self.set_xy(x, y + 1.5)
        self.cell(w, 4, label, align="C")
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(180, 20, 80)
        self.set_xy(x, y + 6)
        self.cell(w, 6, valor, align="C")
        self.set_text_color(0, 0, 0)
        self.set_draw_color(200, 200, 200)
        self.set_line_width(0.2)

# ─────────────────────────────────────────────
# GERADOR PDF PRINCIPAL
# ─────────────────────────────────────────────

def gerar_pdf(
    cidade, relatorio_cat, pgto_totais, pgto_quantidades,
    perfil_clientes, faturamento_perfil, total_geral,
    ticket_geral, ticket_novos, ticket_exist,
    vendas_diarias, vendas_semanais, vendedor_stats,
    campanha_stats,
    dados_kommo,
    indicadores=None
):
    pdf = LeahPDF(cidade)
    pdf.add_page()

    # Título
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(180, 20, 80)
    pdf.cell(0, 10, "Relatório de Vendas", ln=True, align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, pdf.cidade, ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(4)

    # KPIs
    n_novos  = perfil_clientes.get("Novo", 0)
    n_exist  = perfil_clientes.get("Existente", 0)
    n_total  = n_novos + n_exist
    total_servicos = sum(d["vendas"] for d in relatorio_cat.values())

    y_kpi = pdf.get_y()
    pdf.kpi("Faturamento Total",    f"R$ {total_geral:,.2f}", 10,  y_kpi, 45)
    pdf.kpi("Serviços Vendidos",    str(total_servicos),      59,  y_kpi, 45)
    pdf.kpi("Total de Clientes",    str(n_total),             108, y_kpi, 45)
    pdf.kpi("Ticket Médio",         f"R$ {ticket_geral:,.2f}", 157, y_kpi, 43)
    pdf.ln(20)

    # Metas (3 faixas + meta diária dinâmica)
    m = (indicadores or {}).get("metas")
    if m:
        nomes_mes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
                     "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        ano_m, mes_m = m["mes"]
        pdf.titulo_secao(f"Metas de {nomes_mes[mes_m - 1]}/{ano_m}")
        cols_res = [(110, "Acompanhamento"), (80, "Valor")]
        pdf.cabecalho_tabela(cols_res)
        rotulo_acum = (f"Faturamento acumulado (01/{mes_m:02d} a {m['hoje']:%d/%m})"
                       if m["aberto"] else "Faturamento do mês")
        linhas_res = [
            (rotulo_acum,                                       f"R$ {m['acumulado']:,.2f}"),
            ("Dias de atendimento: decorridos / no mês",        f"{m['dias_decorridos']} / {m['dias_total']}"),
            ("Dias de atendimento restantes (contando hoje)" if m["aberto"] else "Dias restantes",
                                                                str(m["dias_restantes"])),
            ("Média diária realizada (dias encerrados)",        f"R$ {m['media_diaria']:,.2f}"),
        ]
        if m["aberto"]:
            linhas_res.append(("Projeção de fechamento (ritmo atual)", f"R$ {m['projecao']:,.2f}"))
        for i, (desc, val) in enumerate(linhas_res):
            pdf.linha_tabela_com_cabecalho([(110, desc, "L"), (80, val, "R")],
                                           cols_res, (i % 2 == 0))
        pdf.ln(4)

        cols_meta = [(26, "Faixa"), (30, "Meta (R$)"), (20, "% Alcanç."),
                     (34, "Falta (R$)"), (40, "Meta diária original"), (40, "Meta diária necessária")]
        pdf.cabecalho_tabela(cols_meta)
        for i, f in enumerate(m["faixas"]):
            if f["atingida"]:
                falta, nec = "Atingida", "-"
            else:
                falta = f"R$ {f['falta']:,.2f}"
                nec = f"R$ {f['diaria_necessaria']:,.2f}" if m["aberto"] and m["dias_restantes"] > 0 else "-"
            pdf.linha_tabela_com_cabecalho([
                (26, f["nome"], "L"),
                (30, f"R$ {f['valor']:,.2f}", "R"),
                (20, f"{f['perc']:.1f}%", "R"),
                (34, falta, "R"),
                (40, f"R$ {f['diaria_original']:,.2f}", "R"),
                (40, nec, "R"),
            ], cols_meta, (i % 2 == 0))
        pdf.ln(3)

        cols_sit = [(26, "Faixa"), (164, "Situação do ritmo")]
        pdf.cabecalho_tabela(cols_sit)
        for i, f in enumerate(m["faixas"]):
            pdf.linha_tabela_com_cabecalho([(26, f["nome"], "L"), (164, f["situacao"], "L")],
                                           cols_sit, (i % 2 == 0))
        pdf.set_font("Helvetica", "I", 7)
        pdf.set_text_color(120, 120, 120)
        pdf.multi_cell(0, 4, "Meta diária original = meta / dias de atendimento do mês. Meta diária necessária = "
                             "(meta - faturamento acumulado) / dias de atendimento restantes, contando hoje. "
                             "Se o ritmo realizado já supera o necessário, dá para reduzir o ritmo; "
                             "se está abaixo, é preciso acelerar. Dias de atendimento: segunda a sábado, "
                             "sem feriados nacionais.")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)

    ind = indicadores or {}
    grupos = ind.get("grupos", {})

    # ═══════════════ PÁGINA 1 — METAS (acima) ═══════════════

    # ═══════════════ PÁGINA 2 — FACIAL · CORPORAL · DEPILAÇÃO ═══════════════
    pdf.add_page()
    pdf.titulo_secao("Faturamento por Categoria")
    cols_h = [(55, "Categoria"), (16, "Vendas"), (38, "Faturamento (R$)"),
              (27, "Cli. Novos"), (27, "Cli. Exist."), (27, "% Repr.")]
    pdf.cabecalho_tabela(cols_h)

    vazio = {"vendas": 0, "faturado": 0.0, "novos": set(), "existentes": set()}
    tot_vendas, tot_fat, tot_novos, tot_exist = 0, 0.0, set(), set()
    for i, g in enumerate(GRUPOS):
        d = grupos.get(g, vazio)
        perc = d["faturado"] / total_geral * 100 if total_geral else 0
        pdf.linha_tabela_com_cabecalho([
            (55, g,                         "L"),
            (16, str(d["vendas"]),          "C"),
            (38, f"{d['faturado']:,.2f}",   "R"),
            (27, str(len(d["novos"])),      "C"),
            (27, str(len(d["existentes"])), "C"),
            (27, f"{perc:.2f}%",            "R"),
        ], cols_h, (i % 2 == 0))
        tot_vendas += d["vendas"]
        tot_fat += d["faturado"]
        tot_novos |= d["novos"]
        tot_exist |= d["existentes"]
    perc_tot = tot_fat / total_geral * 100 if total_geral else 0
    pdf.linha_total([
        (55, "TOTAL",                    "L"),
        (16, str(tot_vendas),            "C"),
        (38, f"R$ {tot_fat:,.2f}",       "R"),
        (27, str(len(tot_novos)),        "C"),
        (27, str(len(tot_exist)),        "C"),
        (27, f"{perc_tot:.2f}%",         "R"),
    ])
    outros = grupos.get("Outros")
    if outros and outros["faturado"] > 0.005:
        pdf.set_font("Helvetica", "I", 7)
        pdf.set_text_color(120, 120, 120)
        pdf.ln(1)
        pdf.multi_cell(0, 4, f"Não entram nas 3 categorias: R$ {outros['faturado']:,.2f} em itens sem "
                             f"categoria Facial, Corporal ou Depilação (ficam fora da comissão).")
        pdf.set_text_color(0, 0, 0)

    # ═══════════════ PÁGINA 3 — PERFIL · TICKET · SERVIÇOS POR CLIENTE ═══════════════
    pdf.add_page()
    pdf.titulo_secao("Perfil dos Clientes  (Novos vs Existentes)")
    cols_perf = [(90, "Tipo"), (25, "Qtd"), (45, "Faturamento (R$)"), (30, "% Qtd")]
    pdf.cabecalho_tabela(cols_perf)
    fat_novo  = faturamento_perfil.get("Novo", 0.0)
    fat_exist = faturamento_perfil.get("Existente", 0.0)
    for i, (tipo, n, fat) in enumerate([
        ("Clientes Novos",      n_novos, fat_novo),
        ("Clientes Existentes", n_exist, fat_exist),
    ]):
        perc = n / n_total * 100 if n_total else 0
        pdf.linha_tabela_com_cabecalho([
            (90, tipo,              "L"),
            (25, str(n),            "C"),
            (45, f"R$ {fat:,.2f}",  "R"),
            (30, f"{perc:.1f}%",    "R"),
        ], cols_perf, (i % 2 == 0))
    pdf.ln(6)

    pdf.titulo_secao("Ticket Médio")
    cols_tk = [(110, "Descrição"), (80, "Valor (R$)")]
    pdf.cabecalho_tabela(cols_tk)
    for i, (desc, val) in enumerate([
        ("Ticket Médio Geral",                 ticket_geral),
        ("Ticket Médio - Clientes Novos",      ticket_novos),
        ("Ticket Médio - Clientes Existentes", ticket_exist),
    ]):
        pdf.linha_tabela_com_cabecalho([(110, desc, "L"), (80, f"R$ {val:,.2f}", "R")],
                                        cols_tk, (i % 2 == 0))
    pdf.ln(6)

    if ind.get("clientes_unicos"):
        pdf.titulo_secao("Serviços por Cliente e Valor Médio")
        cols_sc = [(110, "Indicador"), (80, "Valor")]
        pdf.cabecalho_tabela(cols_sc)
        linhas_sc = [
            ("Serviços vendidos no período",                    str(ind["servicos"])),
            ("Clientes únicos atendidos",                       str(ind["clientes_unicos"])),
            ("Serviços por cliente",                            f"{ind['serv_por_cliente']:.2f}"),
            ("Valor médio por serviço",                         f"R$ {ind['valor_medio']:,.2f}"),
            ("Valor médio por cliente (serviços/cliente x valor médio)",
                                                                f"R$ {ind['ticket_cliente']:,.2f}"),
        ]
        for i, (desc, val) in enumerate(linhas_sc):
            pdf.linha_tabela_com_cabecalho([(110, desc, "L"), (80, val, "R")],
                                           cols_sc, (i % 2 == 0))
        pdf.ln(6)

    # ═══════════════ PÁGINA 4 — VENDEDORAS: RESUMO + UM BLOCO POR VENDEDORA ═══════════════
    pdf.add_page()
    pdf.titulo_secao("Desempenho por Vendedor")
    cols_vend = [(60, "Vendedor"), (20, "Vendas"), (40, "Faturamento"),
                 (30, "% do Total"), (40, "Comissão")]
    pdf.cabecalho_tabela(cols_vend)

    vend_ord = sorted(vendedor_stats.items(), key=lambda x: x[1]["faturamento"], reverse=True)
    t_fat = t_com = 0.0
    for i, (nome_v, stats) in enumerate(vend_ord):
        perc_v = stats["faturamento"] / total_geral * 100 if total_geral else 0
        pdf.linha_tabela_com_cabecalho([
            (60, nome_curto(nome_v) + (" *" if stats["sem_meta"] else ""), "L"),
            (20, str(stats["qtd"]),                   "C"),
            (40, f"R$ {stats['faturamento']:,.2f}",   "R"),
            (30, f"{perc_v:.1f}%",                    "R"),
            (40, f"R$ {stats['comissao']:,.2f}",      "R"),
        ], cols_vend, (i % 2 == 0))
        t_fat += stats["faturamento"]
        t_com += stats["comissao"]
    pdf.linha_total([
        (60, "TOTAL", "L"), (20, "", "C"), (40, f"R$ {t_fat:,.2f}", "R"),
        (30, "", "R"), (40, f"R$ {t_com:,.2f}", "R"),
    ])
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(120, 120, 120)
    pdf.ln(1)
    taxas = "  ".join(f"{g} {COMISSAO[g]*100:g}%".replace(".", ",") for g in GRUPOS)
    pdf.multi_cell(0, 4,
        f"Taxas: {taxas}. Regra por categoria: abaixo de 90% da meta da categoria não comissiona; "
        "de 90% a 99,99% paga de 0 a 50% da comissão, crescendo a partir de 90%; "
        "a partir de 100% paga a comissão inteira. * = vendedora sem meta cadastrada: "
        "comissão cheia, sem a regra de metas.")
    per = ind.get("periodo")
    if per and (per[0].day != 1 or per[1] < per[1].replace(day=calendar.monthrange(per[1].year, per[1].month)[1])):
        pdf.set_x(pdf.l_margin)   # fpdf2 recente deixa o cursor na margem direita após multi_cell
        pdf.multi_cell(0, 4, "Atenção: a meta é mensal e o % atingido usa o vendido no período do "
                             "relatório; antes do fechamento do mês o % e a comissão são parciais.")
    pdf.set_text_color(0, 0, 0)

    cols_b = [(28, "Categoria"), (14, "Vendas"), (30, "Faturamento"), (18, "% Repr."),
              (28, "Meta (R$)"), (20, "% Meta"), (18, "Fator"), (34, "Comissão")]
    for nome_v, stats in vend_ord:
        if stats["faturamento"] <= 0:
            continue
        if pdf.get_y() > 195:
            pdf.add_page()
        pdf.titulo_secao(f"Vendedora: {nome_curto(nome_v, 40)}" + ("  (sem meta cadastrada)" if stats["sem_meta"] else ""))
        pdf.cabecalho_tabela(cols_b)
        tot_q, tot_c = 0, 0.0
        for i, g in enumerate(GRUPOS):
            d = stats["detalhe"][g]
            repr_ = d["valor"] / stats["faturamento"] * 100 if stats["faturamento"] else 0
            pdf.linha_tabela_com_cabecalho([
                (28, g,                                                    "L"),
                (14, str(d["qtd"]),                                        "C"),
                (30, f"R$ {d['valor']:,.2f}",                              "R"),
                (18, f"{repr_:.1f}%",                                      "R"),
                (28, f"R$ {d['meta']:,.2f}" if d["meta"] else "-",         "R"),
                (20, f"{d['perc']:.1f}%" if d["perc"] is not None else "-", "R"),
                (18, f"{d['fator']*100:.0f}%" if d["perc"] is not None else "100%*", "R"),
                (34, f"R$ {d['comissao']:,.2f}",                           "R"),
            ], cols_b, (i % 2 == 0))
            tot_q += d["qtd"]
            tot_c += d["comissao"]
        pdf.linha_total([
            (28, "TOTAL", "L"), (14, str(tot_q), "C"),
            (30, f"R$ {sum(stats['detalhe'][g]['valor'] for g in GRUPOS):,.2f}", "R"),
            (18, "", "R"), (28, "", "R"), (20, "", "R"), (18, "", "R"),
            (34, f"R$ {tot_c:,.2f}", "R"),
        ])
        pdf.ln(4)

    # ═══════════════ PÁGINA 5 — reservada (veja PAGINA_5 no topo do arquivo) ═══════════════
    if PAGINA_5 is not None:
        pdf.add_page()
        PAGINA_5(pdf, ind)

    # ═══════════════ PÁGINA 6 EM DIANTE — ANÁLISE VISUAL E RESULTADO DIÁRIO ═══════════════
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(180, 20, 80)
    pdf.cell(0, 10, "Análise Visual", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)

    tmp_files = []

    f_cat = grafico_categorias(grupos_para_grafico(grupos), total_geral)
    tmp_files.append(f_cat)
    pdf.image(f_cat, x=10, w=190)
    pdf.ln(4)

    f_perf = grafico_perfil(perfil_clientes, faturamento_perfil)
    tmp_files.append(f_perf)
    if pdf.get_y() > 200:
        pdf.add_page()
    pdf.image(f_perf, x=10, w=190)
    pdf.ln(4)

    pdf.add_page()
    f_vend = grafico_vendedores(vendedor_stats)
    tmp_files.append(f_vend)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(180, 20, 80)
    pdf.cell(0, 10, "Desempenho por Vendedor", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(2)
    pdf.image(f_vend, x=10, w=190)
    pdf.ln(6)

    f_dia = grafico_diario(vendas_diarias, vendas_semanais)
    if f_dia:
        tmp_files.append(f_dia)
        if pdf.get_y() > 190:
            pdf.add_page()
        pdf.image(f_dia, x=10, w=190)
        pdf.ln(4)

    pdf.add_page()
    pdf.titulo_secao("Resultado Semanal (Domingo a Sábado)")
    cols_sem = [(100, "Período"), (90, "Faturamento (R$)")]
    pdf.cabecalho_tabela(cols_sem)
    sems_ordenadas = sorted(vendas_semanais)
    for i, sem in enumerate(sems_ordenadas):
        if pdf.get_y() > 250 and i < len(sems_ordenadas) - 1:
            pdf.add_page()
            pdf.titulo_secao("Resultado Semanal (Domingo a Sábado)")
            pdf.cabecalho_tabela(cols_sem)
        fim = sem + timedelta(days=6)
        txt = f"De {sem.strftime('%d/%m')} a {fim.strftime('%d/%m/%Y')}"
        pdf.linha_tabela_com_cabecalho(
            [(100, txt, "L"), (90, f"R$ {vendas_semanais[sem]:,.2f}", "R")],
            cols_sem, (i % 2 == 0)
        )
    pdf.ln(6)

    pdf.titulo_secao("Resultado Diário")
    cols_dia = [(100, "Data"), (90, "Faturamento (R$)")]
    pdf.cabecalho_tabela(cols_dia)
    dias_ordenados = sorted(vendas_diarias)
    for i, dia in enumerate(dias_ordenados):
        if pdf.get_y() > 250 and i < len(dias_ordenados) - 1:
            pdf.add_page()
            pdf.titulo_secao("Resultado Diário")
            pdf.cabecalho_tabela(cols_dia)
        txt = f"{dia.strftime('%d/%m/%Y')} ({DIAS_PT.get(dia.weekday(), '')})"
        pdf.linha_tabela_com_cabecalho(
            [(100, txt, "L"), (90, f"R$ {vendas_diarias[dia]:,.2f}", "R")],
            cols_dia, (i % 2 == 0)
        )

    # ─── SEÇÃO KOMMO (só aparece se houver dados) ──────────────
    if dados_kommo and dados_kommo.get("total", 0) > 0:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_text_color(180, 20, 80)
        pdf.cell(0, 10, "Métricas do Kommo", ln=True, align="C")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)

        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(60, 8, f"Total de Leads: {dados_kommo['total']}", ln=0)
        pdf.cell(60, 8, f"Ganhos: {dados_kommo['ganhos']}", ln=0)
        pdf.cell(60, 8, f"Perdidos: {dados_kommo['perdidos']}", ln=1)
        pdf.cell(60, 8, f"Abertos: {dados_kommo['abertos']}", ln=0)
        pdf.cell(60, 8, f"Valor Total Ganho: R$ {dados_kommo['valor_total_ganho']:,.2f}", ln=0)
        pdf.cell(60, 8, f"Ticket Médio: R$ {dados_kommo['ticket_medio']:,.2f}", ln=1)
        if dados_kommo['tempo_medio_fechamento']:
            tempo_str = kommo.formatar_tempo(int(dados_kommo['tempo_medio_fechamento']))
            pdf.cell(0, 8, f"Tempo Médio de Fechamento: {tempo_str}", ln=1)
        pdf.ln(4)

        for titulo, chave, cab in [("Leads por Funil", "por_funil", "Funil"),
                                   ("Leads por Etapa", "por_etapa", "Etapa"),
                                   ("Motivos de Perda", "motivos_perda", "Motivo")]:
            if dados_kommo[chave]:
                pdf.titulo_secao(titulo)
                cols_k = [(100, cab), (90, "Quantidade")]
                pdf.cabecalho_tabela(cols_k)
                for i, (nome_k, qtd) in enumerate(dados_kommo[chave].items()):
                    pdf.linha_tabela_com_cabecalho(
                        [(100, nome_k[:40], "L"), (90, str(qtd), "C")], cols_k, (i % 2 == 0))
                pdf.ln(4)

    # Finaliza e salva
    pdf_bytes = bytes(pdf.output())

    for f in tmp_files:
        if f and os.path.exists(f):
            os.remove(f)

    return pdf_bytes

# ─────────────────────────────────────────────
# PROCESSAMENTO DE CADA UNIDADE
# ─────────────────────────────────────────────

def processar_unidade(letra, cidade, dict_id, dict_nome, metas=None):
    arq_demo = f"{letra}v.xlsx"
    arq_det  = f"{letra}d.xlsx"

    if not os.path.exists(arq_demo):
        print(f"  [skip] {arq_demo} não encontrado.")
        return
    if not os.path.exists(arq_det):
        print(f"  [skip] {arq_det} não encontrado.")
        return

    print(f"Processando {cidade}…")

    vendas_df, pgto_totais, pgto_quantidades = ler_demonstrativo(arq_demo)
    det_df, campanha_por_venda = ler_detalhamento(arq_det)

    formas_por_venda = extrair_formas_pagamento_por_venda(arq_demo)
    pdf = agregar_e_gerar(cidade, vendas_df, det_df, campanha_por_venda, formas_por_venda,
                          pgto_totais, pgto_quantidades, dict_id, dict_nome,
                          metas=metas if metas is not None else metas_padrao_da_unidade(cidade))
    nome_arquivo = f"{cidade.replace(' ', '_')}.pdf"
    with open(nome_arquivo, "wb") as f:
        f.write(pdf)
    print(f"  ✓ Gerado: {nome_arquivo}")

def agregar_e_gerar(cidade, vendas_df, det_df, campanha_por_venda, formas_por_venda,
                    pgto_totais, pgto_quantidades, dict_id, dict_nome,
                    dados_kommo=None, id_cliente_por_venda=None,
                    metas=None, vendas_mes=None, mes_ref=None, hoje=None, dias_restantes=None,
                    metas_vend=None, periodo=None):
    """Agrega os dados (já carregados de planilhas OU da API) e devolve os bytes do PDF.

    metas          : (inicial, top, turbo) do mês, ou None para não mostrar a seção de metas
    vendas_mes     : {date: faturamento do dia} do mês inteiro (padrão: o próprio período)
    mes_ref        : (ano, mês) das metas (padrão: mês da última venda)
    hoje           : data de hoje (padrão: date.today())
    dias_restantes : força os dias restantes do mês (opcional)
    """
    vendas_df = vendas_df.copy()

    if dados_kommo is None and kommo is not None:
        data_min, data_max = vendas_df["data_venda"].min(), vendas_df["data_venda"].max()
        if pd.notna(data_min) and pd.notna(data_max):
            try:
                dados_kommo = kommo.obter_dados_kommo(
                    cidade, int(data_min.timestamp()), int(data_max.timestamp()))
            except Exception as e:
                print(f"  [aviso] Kommo indisponível: {e}")

    if id_cliente_por_venda is None:
        id_cliente_por_venda = {}
        for _, row in det_df.iterrows():
            vid = row["id_str"]
            if vid not in id_cliente_por_venda:
                id_cliente = row.get("ID Cliente", "")
                if pd.notna(id_cliente):
                    id_cliente_por_venda[vid] = str(id_cliente).strip()

    total_geral = vendas_df["valor_liq"].sum()

    # ── Perfil novo vs existente ──
    vendas_df["perfil"] = vendas_df.apply(
        lambda r: perfil_cliente_na_data(
            id_cliente_por_venda.get(r["id_str"], ""),
            r["cliente_nome"],
            r["data_venda"],
            dict_id,
            dict_nome
        ),
        axis=1
    )

    perfil_clientes    = vendas_df["perfil"].value_counts().to_dict()
    faturamento_perfil = vendas_df.groupby("perfil")["valor_liq"].sum().to_dict()

    # ── Campanha Origem ──
    campanha_stats = {}
    for _, venda in vendas_df.iterrows():
        vid = venda["id_str"]
        camp = campanha_por_venda.get(vid, "Não informado")
        if camp not in campanha_stats:
            campanha_stats[camp] = {"faturamento": 0.0, "vendas": 0}
        campanha_stats[camp]["faturamento"] += venda["valor_liq"]
        campanha_stats[camp]["vendas"] += 1


    # ── Faturamento por categoria ──
    relatorio_cat = {}
    relatorio_grp = {}          # Facial / Corporal / Depilação (+ "Outros" se não classificado)
    vend_grp = {}               # {vendedora: {grupo: valor líquido}}
    vend_qtd = {}               # {vendedora: {grupo: nº de serviços}}
    for _, venda in vendas_df.iterrows():
        vid   = venda["id_str"]
        vliq  = venda["valor_liq"]
        perf  = venda["perfil"]
        nome  = normalizar_nome(venda["cliente_nome"])

        servicos = det_df[det_df["id_str"] == vid]
        vend_nome = str(venda["Responsável"]).strip()
        vg = vend_grp.setdefault(vend_nome, {})
        vq = vend_qtd.setdefault(vend_nome, {})

        if servicos.empty:
            cat = "Outros"
            formas_valor = _formas_pagamento_da_venda(vid, vliq, formas_por_venda)
            _add_cat(relatorio_cat, cat, vliq, 1, nome, perf, formas_valor)
            _add_cat(relatorio_grp, "Outros", vliq, 1, nome, perf)
            vg["Outros"] = vg.get("Outros", 0.0) + vliq
            vq["Outros"] = vq.get("Outros", 0) + 1
            continue

        soma_bruto = servicos["total_bruto"].sum()
        for _, srv in servicos.iterrows():
            cat = srv["categoria_final"]
            prop = (srv["total_bruto"] / soma_bruto) if soma_bruto > 0 else (1 / len(servicos))
            valor_rateado = vliq * prop
            formas_valor = _formas_pagamento_da_venda(vid, valor_rateado, formas_por_venda)
            _add_cat(relatorio_cat, cat, valor_rateado, 1, nome, perf, formas_valor)
            g = srv["grupo"] if "grupo" in srv.index and pd.notna(srv["grupo"]) \
                else grupo_comissao("", "", cat)
            _add_cat(relatorio_grp, g, valor_rateado, 1, nome, perf)
            vg[g] = vg.get(g, 0.0) + valor_rateado
            vq[g] = vq.get(g, 0) + 1

    # ── Vendedor ──
    vendedor_stats = {}
    for _, venda in vendas_df.iterrows():
        nome_vend = str(venda["Responsável"]).strip()
        if nome_vend not in vendedor_stats:
            vendedor_stats[nome_vend] = {"qtd": 0, "faturamento": 0.0}
        vendedor_stats[nome_vend]["qtd"]         += 1
        vendedor_stats[nome_vend]["faturamento"] += venda["valor_liq"]
    for nome_vend, st in vendedor_stats.items():
        grp = vend_grp.get(nome_vend, {})
        st["grupos"] = {g: grp.get(g, 0.0) for g in GRUPOS}
        st["outros"] = grp.get("Outros", 0.0)
        cfg = buscar_meta_vendedora(nome_vend, metas_vend)
        st["detalhe"] = comissao_vendedora(st["grupos"], vend_qtd.get(nome_vend, {}), cfg)
        st["sem_meta"] = cfg is None
        st["comissao_grupos"] = {g: st["detalhe"][g]["comissao"] for g in GRUPOS}
        st["comissao"] = sum(st["comissao_grupos"].values())

    # ── Diário / semanal ──
    vendas_diarias = {}
    vendas_semanais = {}
    for _, venda in vendas_df.iterrows():
        dt = venda["data_venda"]
        if pd.isna(dt):
            continue
        dia_key = dt.date()
        vendas_diarias[dia_key] = vendas_diarias.get(dia_key, 0.0) + venda["valor_liq"]
        domingo  = dt - timedelta(days=(dt.weekday() + 1) % 7)
        sem_key  = domingo.date()
        vendas_semanais[sem_key] = vendas_semanais.get(sem_key, 0.0) + venda["valor_liq"]

    # ── Ticket médio ──
    n_novos  = perfil_clientes.get("Novo", 0)
    n_exist  = perfil_clientes.get("Existente", 0)
    n_total  = len(vendas_df)
    ticket_geral  = total_geral / n_total if n_total > 0 else 0
    ticket_novos  = faturamento_perfil.get("Novo", 0) / n_novos   if n_novos > 0  else 0
    ticket_exist  = faturamento_perfil.get("Existente", 0) / n_exist if n_exist > 0 else 0

    # ── Serviços por cliente / valor médio ──
    servicos_total = sum(d["vendas"] for d in relatorio_cat.values())
    chaves_cli = {
        (id_cliente_por_venda.get(v["id_str"]) or normalizar_nome(v["cliente_nome"]))
        for _, v in vendas_df.iterrows()
    } - {""}
    clientes_unicos = len(chaves_cli)
    serv_por_cliente = servicos_total / clientes_unicos if clientes_unicos else 0.0
    valor_medio = total_geral / servicos_total if servicos_total else 0.0
    ticket_cliente = total_geral / clientes_unicos if clientes_unicos else 0.0

    # ── Metas ──
    info_metas = None
    if metas and any(m and m > 0 for m in metas):
        hoje = hoje or date.today()
        if mes_ref is None:
            ult = vendas_df["data_venda"].max()
            mes_ref = (ult.year, ult.month) if pd.notna(ult) else (hoje.year, hoje.month)
        info_metas = calcular_metas(vendas_mes if vendas_mes is not None else vendas_diarias,
                                    mes_ref, hoje, metas, dias_restantes)

    indicadores = {
        "servicos": servicos_total, "clientes_unicos": clientes_unicos,
        "serv_por_cliente": serv_por_cliente, "valor_medio": valor_medio,
        "ticket_cliente": ticket_cliente, "metas": info_metas,
        "grupos": relatorio_grp, "periodo": periodo,
    }

    # ── Gerar PDF ──
    return gerar_pdf(
        cidade, relatorio_cat, pgto_totais, pgto_quantidades,
        perfil_clientes, faturamento_perfil, total_geral,
        ticket_geral, ticket_novos, ticket_exist,
        vendas_diarias, vendas_semanais, vendedor_stats,
        campanha_stats,
        dados_kommo,
        indicadores
    )

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print(f"Carregando histórico de vendas ({ARQ_RDV})…")
    dict_id, dict_nome = carregar_historico_rdv(ARQ_RDV)
    print(f"  {len(dict_id)} clientes identificados por ID, {len(dict_nome)} por nome.\n")

    for letra, cidade in MAPEAMENTO:
        processar_unidade(letra, cidade, dict_id, dict_nome)

    print("\nTodos os relatórios gerados.")