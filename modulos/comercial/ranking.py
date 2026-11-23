#!/usr/bin/env python3
"""
ranking.py — Gera o ranking.json (top 5 vendedoras do mês, rede toda) a partir do cache/.

Só sai nome, clínica e posição (sem valores e sem dados de cliente), porque o arquivo
vai para o site público do RH.

    python ranking.py                 # mês atual -> ranking.json
    python ranking.py saida.json      # grava em outro caminho
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

CACHE = Path(__file__).resolve().parents[2] / "cache"
TOP_N = 5

# Clínicas que entram no ranking (a Rota/call center fica de fora).
UNIDADES = {"videira": "Videira", "joacaba": "Joaçaba",
            "trezetilias": "Treze Tílias", "camposnovos": "Campos Novos"}

# Quem NÃO é vendedora de balcão (contas de call center, financeiro etc.).
# Compara sem acento e sem maiúsculas; use o nome como aparece no sistema.
EXCLUIR_NOMES = {"gislaine", "juliana toro"}
EXCLUIR_PREFIXOS = ("master",)


def sem_acento(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


def chave(nome: str) -> str:
    return re.sub(r"\s+", " ", sem_acento(nome).strip().lower())


def parse_br(v) -> float:
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[^\d,.]", "", str(v))
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def excluida(nome: str) -> bool:
    k = chave(nome)
    return not k or k in EXCLUIR_NOMES or k.startswith(EXCLUIR_PREFIXOS)


def main():
    hoje = dt.datetime.utcnow() - dt.timedelta(hours=3)          # Brasília
    sufixo = f"/{hoje.month:02d}/{hoje.year}"
    total = defaultdict(float)                                    # chave -> R$
    nomes = {}                                                    # chave -> nome de exibição
    por_cli = defaultdict(lambda: defaultdict(float))             # chave -> {clínica: R$}
    atualizacoes = []

    for slug, cli in UNIDADES.items():
        p = CACHE / f"{slug}.json"
        if not p.exists():
            print(f"aviso: {p.name} não encontrado, clínica ignorada")
            continue
        d = json.loads(p.read_text(encoding="utf-8"))
        atualizacoes.append(str(d.get("atualizado_em", ""))[:16])
        for r in d.get("demonstrativo", []):
            if not str(r.get("dataLancamento", "")).endswith(sufixo):
                continue
            resp = [re.sub(r"\s+", " ", n).strip() for n in str(r.get("responsavel") or "").split(",")]
            resp = [n for n in resp if not excluida(n)]
            if not resp:
                continue
            parte = parse_br(r.get("valorLiquido")) / len(resp)   # venda dividida entre as vendedoras
            for n in resp:
                k = chave(n)
                nomes.setdefault(k, n)
                total[k] += parte
                por_cli[k][cli] += parte

    ordem = sorted((k for k in total if total[k] > 0), key=lambda k: -total[k])[:TOP_N]
    top = [{"pos": i + 1, "nome": nomes[k], "cli": max(por_cli[k], key=por_cli[k].get)}
           for i, k in enumerate(ordem)]
    atualizado = min((a for a in atualizacoes if a), default=hoje.strftime("%Y-%m-%d %H:%M"))

    saida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ranking.json")
    saida.write_text(json.dumps({"mes": f"{hoje.year}-{hoje.month:02d}", "atualizado_em": atualizado,
                                 "top": top}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{saida}: " + " | ".join(f"{t['pos']}º {t['nome']} ({t['cli']})" for t in top))


if __name__ == "__main__":
    main()
