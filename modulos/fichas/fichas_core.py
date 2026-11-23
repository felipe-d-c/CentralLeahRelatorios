"""fichas_core.py — fichas de sessões (PDF) dos clientes com plano ativo.

Mesma regra do fichas.py original (uma ficha por cliente × cidade da credenciada, só quem tem
sessões restantes), mas sem arquivos soltos: entra uma lista de sessões + cadastro de clientes
(vindos da API Belle ou de planilhas) e saem PDFs em memória.
"""
from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

LOGO = Path(__file__).resolve().parent / "logo.png"
CABECALHO = [  # edite aqui os dados da empresa que saem no topo da ficha
    "<b>DEPILEI</b> - 60.287.839/0001-84",
    "Telefone: (49) 99980-3310 | E-mail: financeiro@leahestetica.com.br",
    "R PEDRO NELCIDO KAFER, 75 - Treze Tílias/SC",
]


def _norm(k) -> str:
    t = unicodedata.normalize("NFD", str(k).lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t)


def _num(v) -> float:
    s = str(v if v is not None else "").strip().replace(".", "").replace(",", ".") \
        if isinstance(v, str) and "," in str(v) else str(v if v is not None else "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


# ─────────────────────────────────────────────
# ENTRADA: API Belle  →  linhas canônicas
# ─────────────────────────────────────────────
# O PDF da API não traz o formato da resposta, então cada campo aceita vários nomes possíveis
# (comparados sem acento/maiúscula) e, se nada casar, o erro mostra as chaves que vieram.
_SESSOES = {
    "cli_id":     ["codcliente", "codigocliente", "idcliente", "codigodocliente"],
    "cli_nome":   ["nomecliente", "cliente", "nomedocliente", "codigodocliente"],
    "servico":    ["nomedoservico", "nomeservico", "servico", "descricaoservico"],
    "data_venda": ["datadavenda", "datavenda", "dtvenda", "dtvnd", "dataemissao"],
    "indicacao":  ["indicacao", "nomeindicacao", "credenciada", "nomecredenciada"],
    "restantes":  ["sessoesrestantes", "sessaorestante", "qtdsessoesrestantes", "saldosessoes", "restantes"],
    "vendidas":   ["sessoesvendidas", "qtdsessoesvendidas", "totalsessoes"],
    "plano":      ["nomedoplano", "nomeplano", "plano"],
}
_CLIENTES = {
    "cli_id":   ["id", "codcliente", "codigocliente", "idcliente", "codigo"],
    "cli_nome": ["nome", "nomecliente"],
    "cidade":   ["ufcidade", "cidade", "municipio"],
    "uf":       ["uf", "estado"],
    "telefone": ["celular", "telefone", "fone", "celular1"],
}
_FRAGMENTOS = {"restantes": "restan", "vendidas": "vendid", "servico": "servico", "indicacao": "indica"}


def mapear(registro: dict, esquema: dict) -> dict:
    """{campo_lógico: chave_real_ou_None} para um registro da API."""
    reais = {_norm(k): k for k in registro}
    out = {}
    for campo, cands in esquema.items():
        achada = next((reais[c] for c in cands if c in reais), None)
        frag = _FRAGMENTOS.get(campo)
        if achada is None and frag:
            achada = next((r for n, r in reais.items() if frag in n), None)
        out[campo] = achada
    return out


def _id_nome(texto):
    """'1852502-Cleudes Brunetto' -> ('1852502', 'CLEUDES BRUNETTO')."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return "", ""
    t = str(texto).strip()
    m = re.match(r"^(\d+)\s*[-–]?\s*(.*)$", t)
    if m:
        return m.group(1).strip(), m.group(2).strip().upper()
    return "", re.sub(r"^[\d\-\s]+", "", t).strip().upper()


def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def canon_sessoes_api(brutos: list) -> tuple[list, dict]:
    """Registros brutos de relatorios/sessoes_planos -> linhas canônicas (só com sessões restantes)."""
    if not brutos:
        return [], {"chaves": [], "mapa": {}}
    mapa = mapear(brutos[0], _SESSOES)
    falta = [c for c in ("restantes", "servico") if not mapa[c]] + \
            ([] if (mapa["cli_id"] or mapa["cli_nome"]) else ["cliente"])
    if falta:
        raise ValueError("Não encontrei na resposta da API os campos: " + ", ".join(falta) +
                         ". Chaves recebidas: " + ", ".join(brutos[0].keys()))
    linhas = []
    for r in brutos:
        rest = _num(r.get(mapa["restantes"]))
        if rest <= 0:
            continue
        cid, nome = "", ""
        for campo in ("cli_id", "cli_nome"):
            k = mapa[campo]
            if not k:
                continue
            i, n = _id_nome(r.get(k))
            cid = cid or i or (_s(r.get(k)) if campo == "cli_id" and _s(r.get(k)).isdigit() else "")
            nome = nome or n
        linhas.append({
            "cli_id": cid, "cli_nome": nome,
            "servico": _s(r.get(mapa["servico"])),
            "data_venda": _s(r.get(mapa["data_venda"])) if mapa["data_venda"] else "",
            "indicacao": _s(r.get(mapa["indicacao"])) if mapa["indicacao"] else "",
            "restantes": rest,
            "plano": _s(r.get(mapa["plano"])) if mapa["plano"] else "",
        })
    return linhas, {"chaves": list(brutos[0].keys()), "mapa": mapa}


def canon_clientes_api(brutos: list, ids_necessarios: set | None = None) -> dict:
    """Registros brutos de /clientes -> {id: {nome, cidade, telefone}} (só os ids pedidos, para o cache ficar pequeno)."""
    if not brutos:
        return {}
    m = mapear(brutos[0], _CLIENTES)
    out = {}
    for r in brutos:
        cid = _s(r.get(m["cli_id"])) if m["cli_id"] else ""
        if not cid or (ids_necessarios is not None and cid not in ids_necessarios):
            continue
        cidade = _s(r.get(m["cidade"])) if m["cidade"] else ""
        uf = _s(r.get(m["uf"])) if m["uf"] else ""
        if uf and cidade and " - " not in cidade:
            cidade = f"{uf} - {cidade}"
        out[cid] = {"nome": _s(r.get(m["cli_nome"])).upper() if m["cli_nome"] else "",
                    "cidade": cidade,
                    "telefone": _s(r.get(m["telefone"])) if m["telefone"] else ""}
    return out


# ─────────────────────────────────────────────
# ENTRADA: planilhas (plano B, igual ao fluxo antigo)
# ─────────────────────────────────────────────
def canon_planilhas(sessoes_xlsx, clientes_xlsx) -> tuple[list, dict]:
    df = pd.read_excel(sessoes_xlsx, skiprows=1)
    linhas = []
    for _, r in df.iterrows():
        rest = _num(r.get("Sessões Restantes"))
        if rest <= 0:
            continue
        cid, nome = _id_nome(r.get("Código do Cliente"))
        data = pd.to_datetime(r.get("Data da Venda"), dayfirst=True, errors="coerce")
        linhas.append({"cli_id": cid, "cli_nome": nome, "servico": _s(r.get("Nome do Serviço")),
                       "data_venda": data.strftime("%d/%m/%Y") if pd.notna(data) else _s(r.get("Data da Venda")),
                       "indicacao": _s(r.get("Indicação")), "restantes": rest,
                       "plano": _s(r.get("Nome do Plano"))})
    cli = pd.read_excel(clientes_xlsx, header=0, dtype=str)
    c = cli.columns.tolist()
    clientes = {}
    for _, r in cli.iterrows():
        cid = _s(r[c[0]])
        clientes[cid] = {"nome": _s(r[c[1]]).upper(), "cidade": _s(r[c[7]]), "telefone": _s(r[c[9]])}
    return linhas, clientes


# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────
def _sanitizar(s) -> str:
    return re.sub(r'[<>:"/\\|?*]', "", str(s)).strip()


def _credenciada(texto, cidade_cliente):
    """(pasta_cidade, label_credenciada) — mesma regra do fichas.py original."""
    texto = _s(texto)
    cidade_cliente = "" if cidade_cliente.strip(" -") == "" else cidade_cliente
    if not texto:
        cidade = cidade_cliente or "SEM CIDADE"
        return _sanitizar(cidade), _sanitizar(cidade)
    m = re.search(r"-\s*(.+)$", texto)
    if m:
        return _sanitizar(m.group(1)), _sanitizar(texto)
    cidade = cidade_cliente or "SEM CIDADE"
    return _sanitizar(cidade), _sanitizar(f"{texto} - {cidade}")


def _ficha_pdf(nome, telefone, label, servicos) -> bytes | None:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=20, bottomMargin=20)
    est = getSampleStyleSheet()
    mini = est["Normal"]
    mini.fontSize, mini.leading = 8, 10
    el = []
    if LOGO.exists():
        logo = Image(str(LOGO), width=90, height=65)
        logo.hAlign = "LEFT"
        el.append(logo)
    else:
        el.append(Paragraph("<b>LEAH ESTÉTICA</b>", est["Normal"]))
    cab = [[Paragraph(t, mini)] for t in CABECALHO] + \
          [[Paragraph(f"Emissão: {datetime.now():%d/%m/%Y %H:%M}", mini)]]
    el += [Table(cab, colWidths=[500]), Spacer(1, 12)]
    linha = f"<b>CLIENTE:</b> {nome}" + (f"&nbsp;&nbsp;&nbsp;<b>Tel:</b> {telefone}" if telefone else "")
    el.append(Paragraph(linha, est["Heading3"]))
    if label:
        el.append(Paragraph(f"<b>CREDENCIADA:</b> {label}", est["Normal"]))
    el.append(Spacer(1, 10))
    dados = [["Serviço", "Data da Venda", "Sessão", "Profissional", "Assinatura"]]
    for servico, data, total in servicos:
        n = int(total)
        for i in range(n):
            dados.append([Paragraph(servico, mini), str(data), f"{i + 1}/{n}", "", ""])
    if len(dados) <= 1:
        return None
    t = Table(dados, colWidths=[195, 65, 40, 80, 110], splitByRow=True, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6), ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke)]))
    el.append(t)
    doc.build(el)
    return buf.getvalue()


def agrupar(linhas: list, clientes: dict) -> list[dict]:
    """Uma entrada por cliente × pasta (cidade da credenciada), sem gerar o PDF ainda."""
    grupos: dict = {}
    for r in linhas:
        cad = clientes.get(r["cli_id"], {}) if r["cli_id"] else {}
        nome = cad.get("nome") or r["cli_nome"]
        if not nome or len(nome) < 3:
            continue
        pasta, label = _credenciada(r["indicacao"], cad.get("cidade", ""))
        g = grupos.setdefault((nome, pasta), {"nome": nome, "pasta": pasta, "label": label,
                                              "telefone": cad.get("telefone", ""), "servicos": []})
        g["servicos"].append((r["servico"], r["data_venda"], r["restantes"]))
    return sorted(grupos.values(), key=lambda g: (g["pasta"], g["nome"]))


def gerar_pdf(grupo: dict) -> bytes | None:
    return _ficha_pdf(grupo["nome"], grupo["telefone"], grupo["label"], grupo["servicos"])


def zip_fichas(grupos: list[dict], progresso=None) -> tuple[bytes, int]:
    """ZIP com uma pasta por cidade. Devolve (bytes, quantidade de fichas)."""
    buf, n = io.BytesIO(), 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, g in enumerate(grupos):
            pdf = gerar_pdf(g)
            if pdf:
                z.writestr(f"{g['pasta']}/{_sanitizar(g['nome'])}.pdf", pdf)
                n += 1
            if progresso:
                progresso((i + 1) / len(grupos))
    return buf.getvalue(), n
