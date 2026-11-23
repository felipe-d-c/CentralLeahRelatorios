"""github_cache.py — guarda/lê o cache das Fichas no GitHub, numa branch própria.

Por que uma branch separada? O Streamlit Cloud reimplanta o app a cada commit na branch principal;
um commit no meio da geração derrubaria a sessão. Na branch `cache-fichas` (criada sozinha na
primeira gravação) o app não é reimplantado.

st.secrets:
    [github]
    token  = "ghp_..."            # PAT com permissão Contents: read/write neste repositório
    repo   = "usuario/central-leah"
    branch = "cache-fichas"       # opcional

Sem [github] nos secrets, o cache vai para a pasta local cache_fichas/ (some quando o app reinicia).
"""
from __future__ import annotations

import base64
import json
from pathlib import Path

import requests

API = "https://api.github.com"
LOCAL = Path(__file__).resolve().parent / "cache_fichas"


def _cfg():
    try:
        import streamlit as st
        g = st.secrets.get("github", {})
        if g.get("token") and g.get("repo"):
            return {"token": g["token"], "repo": g["repo"], "branch": g.get("branch", "cache-fichas")}
    except Exception:
        pass
    return None


def _h(cfg, raw=False):
    return {"Authorization": f"Bearer {cfg['token']}",
            "Accept": "application/vnd.github.raw+json" if raw else "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28"}


def destino() -> str:
    c = _cfg()
    return f"GitHub ({c['repo']}, branch {c['branch']})" if c else "pasta local (temporária)"


def _garantir_branch(cfg):
    u = f"{API}/repos/{cfg['repo']}"
    if requests.get(f"{u}/branches/{cfg['branch']}", headers=_h(cfg), timeout=30).status_code == 200:
        return
    padrao = requests.get(u, headers=_h(cfg), timeout=30)
    padrao.raise_for_status()
    base = padrao.json()["default_branch"]
    ref = requests.get(f"{u}/git/ref/heads/{base}", headers=_h(cfg), timeout=30)
    ref.raise_for_status()
    r = requests.post(f"{u}/git/refs", headers=_h(cfg), timeout=30,
                      json={"ref": f"refs/heads/{cfg['branch']}", "sha": ref.json()["object"]["sha"]})
    r.raise_for_status()


def ler(caminho: str):
    """Objeto JSON salvo em `caminho` (ex.: 'fichas/rota.json') ou None."""
    cfg = _cfg()
    if not cfg:
        p = LOCAL / caminho
        return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None
    r = requests.get(f"{API}/repos/{cfg['repo']}/contents/{caminho}", params={"ref": cfg["branch"]},
                     headers=_h(cfg, raw=True), timeout=60)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def salvar(caminho: str, obj, mensagem: str) -> str:
    """Grava `obj` como JSON. Devolve onde ficou (texto para mostrar ao usuário)."""
    texto = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    cfg = _cfg()
    if not cfg:
        p = LOCAL / caminho
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(texto, encoding="utf-8")
        return "pasta local (configure [github] nos secrets para guardar no GitHub)"
    _garantir_branch(cfg)
    url = f"{API}/repos/{cfg['repo']}/contents/{caminho}"
    atual = requests.get(url, params={"ref": cfg["branch"]}, headers=_h(cfg), timeout=30)
    corpo = {"message": mensagem, "branch": cfg["branch"],
             "content": base64.b64encode(texto.encode("utf-8")).decode()}
    if atual.status_code == 200:
        corpo["sha"] = atual.json()["sha"]
    r = requests.put(url, headers=_h(cfg), json=corpo, timeout=120)
    r.raise_for_status()
    return destino()
