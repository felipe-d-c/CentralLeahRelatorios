"""
belle_api.py — Cliente da API Belle Software + montagem dos dados do relatório
==============================================================================
Substitui as planilhas (Xv.xlsx, Xd.xlsx, rdv.xlsx) por dados da API.

Fluxo rápido:
  1) precache.py (GitHub Action, todo dia 00:05) baixa os dados brutos da API e salva em
     cache/<unidade>.json  -> baixar_dados() + salvar_cache()
  2) o app lê esse arquivo e monta o PDF na hora, sem chamar a API
     -> carregar_cache() + gerar_relatorio_de_dados()
Tokens são lidos de key.ini (ou st.secrets["tokens"] como alternativa).
"""

import configparser
import json
import os
import re
import threading
import time
import unicodedata
from collections import deque
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests

import relatorio as rel

BASE = "https://app.bellesoftware.com.br/api/release/controller/IntegracaoExterna/v1.0/"
COD_ESTAB = 1
JANELA_DIAS = 89          # a API limita 90 dias por consulta
JANELA_MESES = 8           # quantos meses o cache noturno guarda (cobre 6 meses com folga)
HIST_DIAS = 730           # quanto histórico olhar para decidir "Novo" x "Existente"
UNIDADES = [c for _, c in rel.MAPEAMENTO]


def slug(txt):
    t = unicodedata.normalize("NFD", str(txt).lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t)


# ─────────────────────────────────────────────
# TOKENS
# ─────────────────────────────────────────────

def carregar_tokens(caminho=None):
    """Lê key.ini no formato:
        [tokens]
        videira = xxxxx
        joacaba = xxxxx
    Retorna {slug_da_cidade: token}. Se não houver key.ini, tenta st.secrets["tokens"].
    """
    tokens = {}
    caminho = caminho or os.path.join(os.path.dirname(os.path.abspath(__file__)), "key.ini")
    if os.path.exists(caminho):
        cp = configparser.ConfigParser(interpolation=None)
        cp.read(caminho, encoding="utf-8")
        for sec in cp.sections():
            for k, v in cp.items(sec):
                if v.strip():
                    tokens[slug(k)] = v.strip()
    else:
        try:
            import streamlit as st
            for k, v in st.secrets.get("tokens", {}).items():
                tokens[slug(k)] = str(v).strip()
        except Exception:
            pass
    return tokens



# ─────────────────────────────────────────────
# CLIENTE HTTP (40 req/min por token)
# ─────────────────────────────────────────────

class BelleClient:
    def __init__(self, token, limite_por_min=36):
        self.token = token
        self.limite = limite_por_min
        self._marcas = deque()
        self._lock = threading.Lock()

    def _esperar(self):
        with self._lock:
            agora = time.time()
            while self._marcas and agora - self._marcas[0] > 60:
                self._marcas.popleft()
            if len(self._marcas) >= self.limite:
                time.sleep(max(0.0, 60 - (agora - self._marcas[0]) + 0.5))
            self._marcas.append(time.time())

    def get(self, caminho, params=None, tentativas=4):
        params = {k: v for k, v in (params or {}).items() if v not in (None, "")}
        ultimo = None
        for i in range(tentativas):
            self._esperar()
            try:
                r = requests.get(BASE + caminho, params=params, timeout=120,
                                 headers={"Authorization": self.token})
                if r.status_code == 429:
                    time.sleep(30)
                    ultimo = "429 (limite de requisições)"
                    continue
                if r.status_code in (401, 403):
                    raise PermissionError("Token recusado pela API (verifique o token).")
                r.raise_for_status()
                dados = r.json()
                if not dados:
                    return []
                if isinstance(dados, dict):
                    raise RuntimeError(f"Resposta inesperada: {str(dados)[:300]}")
                return dados
            except (requests.Timeout, requests.ConnectionError) as e:
                ultimo = e
                time.sleep(3 * (i + 1))
        raise RuntimeError(f"Falha ao consultar {caminho}: {ultimo}")

    @staticmethod
    def _janelas(ini, fim, dias=JANELA_DIAS):
        cur = ini
        while cur <= fim:
            fim_j = min(cur + timedelta(days=dias - 1), fim)
            yield cur, fim_j
            cur = fim_j + timedelta(days=1)

    @staticmethod
    def _d(d):
        return d.strftime("%d/%m/%Y")

    def demonstrativo(self, ini, fim):
        out = []
        for a, b in self._janelas(ini, fim):
            out += self.get("relatorios/demonstrativo_vendas_detalhado", {
                "dtInicio": self._d(a), "dtFim": self._d(b), "codEstab": COD_ESTAB})
        return out

    def vendas_planos(self, ini, fim):
        out = []
        for a, b in self._janelas(ini, fim):
            out += self.get("venda_planos", {
                "tipoPeriodo": "DataVenda", "dtInicio": self._d(a),
                "dtFim": self._d(b), "codEstab": COD_ESTAB})
        return out

    def servicos(self):
        return self.get("servico/listar", {})

    def vendas_resumo(self, ini, fim):
        out = []
        for a, b in self._janelas(ini, fim):
            out += self.get("vendas", {
                "estab": COD_ESTAB, "dtInicio": self._d(a), "dtFim": self._d(b)})
        return out


# ─────────────────────────────────────────────
# MONTAGEM DOS DADOS
# ─────────────────────────────────────────────

DIAGNOSTICO = {}   # {cidade: {...}} preenchido a cada geração (aba "Diagnóstico" do app)

_CHAVES_CATEGORIA = ("categoria", "nomeCategoria", "categoriaServico", "descCategoria",
                     "descricaoCategoria", "categoria_servico", "nome_categoria", "grupo")
_CHAVES_NOME = ("nome", "nomeServico", "descricao", "servico")
_CHAVES_COD = ("codServico", "cod_servico", "codigo", "cod", "id", "idServico")
_CAT_CANONICAS = {slug(k): k for k in rel.MAPA_CATEGORIA}


def _limpa_nome(txt):
    """Tira prefixos tipo '123 - ' e normaliza para comparação."""
    t = re.sub(r"^\s*\d+\s*[-–]\s*", "", str(txt or "")).strip()
    return slug(t)


def _texto(v):
    if isinstance(v, dict):
        v = v.get("nome") or v.get("descricao") or ""
    return "" if v is None else str(v)


def _extrai_categoria(d):
    """Pega a categoria de um registro (serviço ou linha do demonstrativo)."""
    for k in _CHAVES_CATEGORIA:
        if k in d and _texto(d[k]).strip():
            cat = re.sub(r"^\s*\d+\s*[-–]\s*", "", _texto(d[k])).strip()
            return _CAT_CANONICAS.get(slug(cat), cat)   # "depilacao" -> "Depilação"
    return ""


# Itens que não dá para classificar pelo nome -> categoria final (edite à vontade)
OVERRIDE_NOME = {
    "planopersonalizado": "Plexr",
}

# (palavras no nome do item, categoria "do sistema" que o resolver_categoria entende)
REGRAS_NOME = [
    (("plexr",), "Flacidez"),
    (("peim", "varize", "vasinho", "escleroterapia"), "Varicose"),
    (("depila", "virilha", "axila", "buco", "buço", "perna", "laser"), "Depilação"),
    (("preench", "toxina", "botox", "botul", "bioestim", "skinbooster", "peeling",
      "labial", "olheira", "papada", "mandibul"), "Rejuvenescimento"),
    (("emagrec", "gordura"), "Emagrecimento"),
    (("micropigment", "maquiagem"), "Maquiagem Permanente"),
    (("limpeza de pele",), "Remoção"),
]


def categoria_por_nome(nome):
    t = unicodedata.normalize("NFD", str(nome).lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    for palavras, cat in REGRAS_NOME:
        if any(unicodedata.normalize("NFD", p).encode("ascii", "ignore").decode() in t
               for p in palavras):
            return cat
    return ""


class CatServicos:
    """Resolve a categoria do sistema para um item vendido (por código, nome exato,
    nome normalizado ou nome contido)."""

    def __init__(self, servicos):
        self.por_nome, self.por_cod = {}, {}
        for s in servicos:
            cat = _extrai_categoria(s)
            for k in _CHAVES_NOME:
                if s.get(k):
                    self.por_nome[_limpa_nome(s[k])] = cat
                    break
            for k in _CHAVES_COD:
                if s.get(k) not in (None, ""):
                    self.por_cod[str(s[k]).strip()] = cat
                    break
        self.amostra = servicos[:1]

    def buscar(self, nome, cod=None, linha=None):
        if linha:
            c = _extrai_categoria(linha)
            if c:
                return c
        if cod not in (None, "") and str(cod).strip() in self.por_cod:
            return self.por_cod[str(cod).strip()]
        n = _limpa_nome(nome)
        if n in self.por_nome:
            return self.por_nome[n]
        cands = [k for k in self.por_nome if k and n and (k in n or n in k)]
        if cands:
            return self.por_nome[max(cands, key=len)]
        return ""


def _historico_primeira_compra(vendas):
    dict_id = {}
    for v in vendas:
        cid = str(v.get("cod_cliente", "")).strip()
        dt = pd.to_datetime(v.get("data_venda"), errors="coerce")
        if cid and pd.notna(dt) and (cid not in dict_id or dt < dict_id[cid]):
            dict_id[cid] = dt
    return dict_id


def _parcelas_validas(linhas, liq_venda):
    """O demonstrativo pode repetir as parcelas em cada item da venda ou trazê-las
    por item. Escolhe a interpretação cuja soma bate com o valor líquido da venda."""
    todas = [p for ps in linhas for p in ps]
    if not todas:
        return []
    soma_todas = sum(rel.parse_br(p.get("valor")) for p in todas)
    primeira = linhas[0]
    soma_1 = sum(rel.parse_br(p.get("valor")) for p in primeira)
    if abs(soma_todas - liq_venda) <= 0.05 * max(liq_venda, 1):
        return todas
    if primeira and abs(soma_1 - liq_venda) <= 0.05 * max(liq_venda, 1):
        return primeira
    return todas




# ─────────────────────────────────────────────
# CACHE EM DISCO (cache/<unidade>.json) — gerado todo dia pelo precache.py
# ─────────────────────────────────────────────

CACHE_DIR = Path(__file__).resolve().parents[2] / "cache"   # cache único da Central (raiz do repositório)


def agora_br():
    """Horário de Brasília (o servidor do Streamlit/GitHub roda em UTC)."""
    return datetime.utcnow() - timedelta(hours=3)


def hoje_br():
    return agora_br().date()


def inicio_janela(hoje=None, meses=JANELA_MESES):
    """1º dia do mês, `meses` meses atrás."""
    hoje = hoje or hoje_br()
    m = hoje.year * 12 + (hoje.month - 1) - meses
    return date(m // 12, m % 12 + 1, 1)


def _fatia_demonstrativo(r):
    return {
        "idVenda": r.get("idVenda"), "dataLancamento": r.get("dataLancamento"),
        "codCliente": r.get("codCliente"), "nomeCliente": r.get("nomeCliente"),
        "responsavel": r.get("responsavel"), "nomeItem": r.get("nomeItem"),
        "valorBruto": r.get("valorBruto"), "valorLiquido": r.get("valorLiquido"),
        "parcelas": [{"formaPagamento": p.get("formaPagamento"), "valor": p.get("valor")}
                     for p in (r.get("parcelas") or [])],
    }


def _fatia_plano(p):
    return {
        "idVenda": p.get("idVenda"), "nomeCampanhaOrigem": p.get("nomeCampanhaOrigem"),
        "servicos": [{"codigoServico": s.get("codigoServico"), "nomeServico": s.get("nomeServico"),
                      "valorTotalServico": s.get("valorTotalServico")}
                     for s in (p.get("servicos") or [])],
    }


def baixar_dados(token, ini: date, fim: date, anterior=None, progresso=None):
    """Baixa da API os dados brutos de uma unidade para [ini, fim].
    `anterior` = cache antigo: se existir, o histórico de 1ª compra é atualizado só
    a partir da data em que parou (muito mais rápido)."""
    def avisa(msg):
        if progresso:
            progresso(msg)

    c = BelleClient(token)
    avisa("Vendas do período…")
    dem = [_fatia_demonstrativo(r) for r in c.demonstrativo(ini, fim)]
    avisa("Planos e campanhas…")
    planos = [_fatia_plano(p) for p in c.vendas_planos(ini, fim)]
    avisa("Cadastro de serviços…")
    servicos = [{"codServico": s.get("codServico"), "nome": s.get("nome"),
                 "categoria": s.get("categoria")} for s in c.servicos()]

    avisa("Histórico de clientes (novo x existente)…")
    primeira = dict((anterior or {}).get("primeira_compra", {}))
    if primeira and (anterior or {}).get("hist_ate"):
        h_ini = date.fromisoformat(anterior["hist_ate"]) - timedelta(days=3)
    else:
        primeira = {}
        h_ini = ini - timedelta(days=HIST_DIAS)
    for cid, d in _historico_primeira_compra(c.vendas_resumo(h_ini, fim)).items():
        iso = d.strftime("%Y-%m-%d")
        if cid not in primeira or iso < primeira[cid]:
            primeira[cid] = iso

    return {
        "versao": 1,
        "atualizado_em": agora_br().strftime("%Y-%m-%d %H:%M"),
        "cobertura": [ini.isoformat(), fim.isoformat()],
        "hist_ate": fim.isoformat(),
        "demonstrativo": dem, "planos": planos, "servicos": servicos,
        "primeira_compra": primeira,
    }


def caminho_cache(cidade):
    return CACHE_DIR / f"{slug(cidade)}.json"


def salvar_cache(cidade, dados):
    CACHE_DIR.mkdir(exist_ok=True)
    tmp = caminho_cache(cidade).with_suffix(".tmp")
    tmp.write_text(json.dumps(dados, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
                   encoding="utf-8")
    tmp.replace(caminho_cache(cidade))


def carregar_cache(cidade):
    """Dados salvos da unidade ou None."""
    p = caminho_cache(cidade)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def cache_do_dia(dados):
    """True se o cache foi baixado hoje (depois da meia-noite, horário de Brasília)."""
    return bool(dados) and str(dados.get("atualizado_em", ""))[:10] == hoje_br().isoformat()


def atualizar_cache(cidade, token, progresso=None):
    """Baixa os dados novos da API (incremental) e grava em cache/. Devolve os dados."""
    hoje = hoje_br()
    dados = baixar_dados(token, inicio_janela(hoje), hoje,
                         anterior=carregar_cache(cidade), progresso=progresso)
    salvar_cache(cidade, dados)
    return dados


def cache_cobre(dados, ini: date, fim: date):
    if not dados:
        return False
    a, b = (date.fromisoformat(x) for x in dados["cobertura"])
    return a <= ini and fim <= b


# ─────────────────────────────────────────────
# MONTAGEM DO RELATÓRIO A PARTIR DOS DADOS
# ─────────────────────────────────────────────

def vendas_diarias_do_mes(dados, ano, mes):
    """{date: faturamento líquido do dia} do mês inteiro, direto do cache (independe do período)."""
    out = {}
    for r in dados["demonstrativo"]:
        d = pd.to_datetime(r.get("dataLancamento"), dayfirst=True, errors="coerce")
        if pd.notna(d) and d.year == ano and d.month == mes:
            out[d.date()] = out.get(d.date(), 0.0) + rel.parse_br(r.get("valorLiquido"))
    return out


def gerar_relatorio_de_dados(cidade, dados, ini: date, fim: date, metas=None, dias_restantes=None,
                             metas_vend=None):
    """Monta o PDF da unidade para [ini, fim] usando dados já baixados (sem API).
    metas = (inicial, top, turbo) do mês de `fim`; o acompanhamento usa o mês inteiro até hoje."""
    dt_ini, dt_fim = pd.Timestamp(ini), pd.Timestamp(fim)
    linhas = []
    for r in dados["demonstrativo"]:
        d = pd.to_datetime(r.get("dataLancamento"), dayfirst=True, errors="coerce")
        if pd.notna(d) and dt_ini <= d <= dt_fim:
            linhas.append(r)
    if not linhas:
        raise ValueError("Nenhuma venda encontrada no período.")

    campanha_por_venda, servicos_plano = {}, {}
    for p in dados["planos"]:
        vid = str(p.get("idVenda", "")).strip()
        nome = str(p.get("nomeCampanhaOrigem") or "").strip()
        if vid:
            campanha_por_venda[vid] = nome or "Não informado"
            servicos_plano.setdefault(vid, []).extend(p.get("servicos") or [])
    lista_serv = dados["servicos"]
    cat_serv = CatServicos(lista_serv)
    dict_id = {cid: pd.Timestamp(d) for cid, d in dados["primeira_compra"].items()}

    vendas, itens, formas_por_venda = {}, [], {}
    for r in linhas:
        vid = str(r.get("idVenda", "")).strip()
        liq = rel.parse_br(r.get("valorLiquido"))
        bruto = rel.parse_br(r.get("valorBruto"))
        v = vendas.setdefault(vid, {
            "id_str": vid,
            "data_venda": pd.to_datetime(r.get("dataLancamento"), dayfirst=True, errors="coerce"),
            "cod_cliente": str(r.get("codCliente", "")).strip(),
            "cliente_nome": str(r.get("nomeCliente", "")).strip(),
            "Responsável": str(r.get("responsavel", "")).strip(),
            "valor_liq": 0.0, "valor_bruto": 0.0, "_parc": [],
        })
        v["valor_liq"] += liq
        v["valor_bruto"] += bruto
        v["_parc"].append(r.get("parcelas") or [])
        nome = str(r.get("nomeItem", "")).strip()
        cat_sis = cat_serv.buscar(nome) or categoria_por_nome(nome)
        cat_final = OVERRIDE_NOME.get(slug(nome)) or rel.resolver_categoria(cat_sis, nome)
        itens.append({"id_str": vid, "total_bruto": bruto, "total_liq": liq,
                      "categoria_final": cat_final, "_nome": nome,
                      "grupo": rel.grupo_comissao(cat_sis, nome, cat_final)})

    pgto_totais, pgto_quantidades = {}, {}
    for vid, v in vendas.items():
        parcelas = _parcelas_validas(v.pop("_parc"), v["valor_liq"])
        formas_por_venda[vid] = []
        usadas = set()
        for p in parcelas:
            forma = rel.normalizar_forma_pgto(p.get("formaPagamento", ""))
            val = rel.parse_br(p.get("valor"))
            formas_por_venda[vid].append((forma, val))
            pgto_totais[forma] = pgto_totais.get(forma, 0.0) + val
            if forma not in usadas:          # parcelas da mesma forma = 1 pagamento
                usadas.add(forma)
                pgto_quantidades[forma] = pgto_quantidades.get(forma, 0) + 1

    # Itens por serviço (como no antigo Xd.xlsx): venda de plano -> serviços do plano
    por_venda = {}
    for it in itens:
        por_venda.setdefault(it["id_str"], []).append(it)
    itens, n_plano, sem_categoria = [], 0, set()
    for vid, lst in por_venda.items():
        svs = [x for x in servicos_plano.get(vid, []) if x.get("nomeServico")]
        if svs:
            n_plano += 1
            lst = []
            for x in svs:
                nm = str(x.get("nomeServico", "")).strip()
                val = rel.parse_br(x.get("valorTotalServico"))
                cs = cat_serv.buscar(nm, x.get("codigoServico")) or categoria_por_nome(nm)
                cf = OVERRIDE_NOME.get(slug(nm)) or rel.resolver_categoria(cs, nm)
                lst.append({"id_str": vid, "total_bruto": val, "total_liq": val,
                            "categoria_final": cf, "_nome": nm,
                            "grupo": rel.grupo_comissao(cs, nm, cf)})
        for it in lst:
            if it["grupo"] == "Outros":
                sem_categoria.add(it.get("_nome") or "(item sem nome)")
        itens += lst

    DIAGNOSTICO[cidade] = {
        "vendas_com_servicos_do_plano": n_plano,
        "sem_categoria": sorted(sem_categoria),
        "cache_atualizado_em": dados.get("atualizado_em"),
        "cobertura": dados.get("cobertura"),
        "n_servicos": len(lista_serv),
        "nomes_itens": sorted({str(r.get("nomeItem", "")).strip() for r in linhas})[:15],
    }

    vendas_df = pd.DataFrame(vendas.values())
    det_df = pd.DataFrame(itens)
    id_cliente_por_venda = {vid: v["cod_cliente"] for vid, v in vendas.items()}
    return rel.agregar_e_gerar(
        cidade, vendas_df, det_df, campanha_por_venda, formas_por_venda,
        pgto_totais, pgto_quantidades, dict_id, {}, None, id_cliente_por_venda,
        metas=metas, vendas_mes=vendas_diarias_do_mes(dados, fim.year, fim.month),
        mes_ref=(fim.year, fim.month), hoje=hoje_br(), dias_restantes=dias_restantes,
        metas_vend=metas_vend, periodo=(ini, fim),
    )


def gerar_relatorio_unidade(cidade, token, ini: date, fim: date, progresso=None,
                            metas=None, dias_restantes=None, metas_vend=None):
    """Plano B (sem cache): baixa da API só o necessário e gera o PDF.
    Baixa desde o dia 1 do mês de `fim`, para o acompanhamento das metas ficar completo."""
    dados = baixar_dados(token, min(ini, fim.replace(day=1)), fim, progresso=progresso)
    return gerar_relatorio_de_dados(cidade, dados, ini, fim, metas=metas,
                                    dias_restantes=dias_restantes, metas_vend=metas_vend)


def vendedoras_do_periodo(dados, ini: date, fim: date):
    """Nomes das vendedoras com venda em [ini, fim], da que mais vendeu para a que menos."""
    tot = {}
    for r in (dados or {}).get("demonstrativo", []):
        d = pd.to_datetime(r.get("dataLancamento"), dayfirst=True, errors="coerce")
        if pd.notna(d) and ini <= d.date() <= fim:
            nome = str(r.get("responsavel", "")).strip()
            tot[nome] = tot.get(nome, 0.0) + rel.parse_br(r.get("valorLiquido"))
    return [n for n, v in sorted(tot.items(), key=lambda x: -x[1]) if n and v > 0]
