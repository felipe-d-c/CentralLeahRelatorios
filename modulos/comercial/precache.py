#!/usr/bin/env python3
"""
precache.py — Baixa da API Belle os dados brutos de cada clínica (últimos ~8 meses)
e salva em cache/<unidade>.json. O app lê esses arquivos e monta os PDFs na hora.

Roda todo dia 00:05 (Brasília) pelo GitHub Actions; pode rodar à mão:
    python precache.py                # todas as clínicas
    python precache.py videira joacaba   # só algumas
"""
from __future__ import annotations

import datetime as dt
import os
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from belle_api import UNIDADES, atualizar_cache, carregar_tokens, hoje_br, inicio_janela, slug

# Secrets do GitHub (mesmos nomes de antes) -> unidade
ENV_TOKENS = {
    "videira": "TOKEN_VIDEIRA", "joacaba": "TOKEN_JOACABA",
    "trezetilias": "TOKEN_TREZE_TILIAS", "camposnovos": "TOKEN_CAMPOS_NOVOS",
    "rota": "TOKEN_ROTA",
}


def tokens_disponiveis():
    tokens = carregar_tokens()
    for chave, env in ENV_TOKENS.items():
        if os.environ.get(env, "").strip():
            tokens[chave] = os.environ[env].strip()
    return tokens


def atualizar(cidade, token, ini, fim):
    dados = atualizar_cache(cidade, token,
                            progresso=lambda m: print(f"  [{cidade}] {m}", flush=True))
    return len(dados["demonstrativo"]), len(dados["primeira_compra"])


def main():
    filtro = {slug(a) for a in sys.argv[1:]}
    tokens = tokens_disponiveis()
    if not tokens:
        print("ERRO: nenhum token encontrado (key.ini ou variáveis TOKEN_*).")
        sys.exit(1)

    hoje = hoje_br()
    ini, fim = inicio_janela(hoje), hoje
    unidades = [u for u in UNIDADES if slug(u) in tokens and (not filtro or slug(u) in filtro)]
    print(f"Atualizando cache: {ini:%d/%m/%Y} → {fim:%d/%m/%Y} · {', '.join(unidades)}\n")

    falhas = 0
    with ThreadPoolExecutor(max_workers=len(unidades) or 1) as ex:   # 1 thread por clínica (cada token tem seu limite)
        futuros = {u: ex.submit(atualizar, u, tokens[slug(u)], ini, fim) for u in unidades}
        for u, f in futuros.items():
            try:
                n_vendas, n_clientes = f.result()
                print(f"✓ {u}: {n_vendas} linhas de venda, {n_clientes} clientes no histórico")
            except Exception as e:                       # mantém o cache antigo dessa clínica
                falhas += 1
                print(f"✗ {u}: {e}")

    print(f"\nConcluído: {len(unidades) - falhas} ok, {falhas} falha(s)")
    sys.exit(1 if falhas else 0)


if __name__ == "__main__":
    main()
