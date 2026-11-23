"""belle_core.py — cliente mínimo e genérico da API Belle (usado pelo módulo Fichas e por futuros relatórios).

Respeita o limite de 40 requisições/min por token (usa 36) e tenta de novo em 429/timeout.
"""
from __future__ import annotations

import re
import threading
import time
import unicodedata
from collections import deque
from datetime import date, timedelta

import requests

BASE = "https://app.bellesoftware.com.br/api/release/controller/IntegracaoExterna/v1.0/"
JANELA_DIAS = 89   # a maioria dos relatórios limita a 90 dias por consulta


def slug(txt) -> str:
    t = unicodedata.normalize("NFD", str(txt).lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t)


def tokens_dos_secrets() -> dict:
    """{slug_da_unidade: token} lido de st.secrets['tokens'] (mesmo formato dos outros relatórios)."""
    try:
        import streamlit as st
        return {slug(k): str(v).strip() for k, v in st.secrets.get("tokens", {}).items() if str(v).strip()}
    except Exception:
        return {}


class BelleClient:
    def __init__(self, token: str, limite_por_min: int = 36):
        self.token = token
        self.limite = limite_por_min
        self._marcas: deque = deque()
        self._lock = threading.Lock()

    def _esperar(self):
        with self._lock:
            agora = time.time()
            while self._marcas and agora - self._marcas[0] > 60:
                self._marcas.popleft()
            if len(self._marcas) >= self.limite:
                time.sleep(max(0.0, 60 - (agora - self._marcas[0]) + 0.5))
            self._marcas.append(time.time())

    def get(self, caminho: str, params: dict | None = None, tentativas: int = 4, timeout: int = 120) -> list:
        """GET que sempre devolve uma lista de registros (a API às vezes embrulha a lista num dict)."""
        params = {k: v for k, v in (params or {}).items() if v not in (None, "")}
        ultimo = None
        for i in range(tentativas):
            self._esperar()
            try:
                r = requests.get(BASE + caminho, params=params, timeout=timeout,
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
                if isinstance(dados, list):
                    return dados
                if isinstance(dados, dict):
                    for v in dados.values():            # {"dados": [...]} / {"clientes": [...]}
                        if isinstance(v, list):
                            return v
                    return [dados]
                return []
            except (requests.Timeout, requests.ConnectionError) as e:
                ultimo = e
                time.sleep(3 * (i + 1))
        raise RuntimeError(f"Falha ao consultar {caminho}: {ultimo}")

    def get_raw(self, caminho: str, params: dict | None = None):
        """GET sem tratamento: (status_http, json_ou_texto). Só para o Explorador."""
        params = {k: v for k, v in (params or {}).items() if v not in (None, "")}
        self._esperar()
        r = requests.get(BASE + caminho.lstrip("/"), params=params, timeout=120,
                         headers={"Authorization": self.token})
        try:
            return r.status_code, r.json()
        except ValueError:
            return r.status_code, r.text[:2000]

    @staticmethod
    def janelas(ini: date, fim: date, dias: int = JANELA_DIAS):
        cur = ini
        while cur <= fim:
            fim_j = min(cur + timedelta(days=dias - 1), fim)
            yield cur, fim_j
            cur = fim_j + timedelta(days=1)

    @staticmethod
    def d(dt: date) -> str:
        return dt.strftime("%d/%m/%Y")
