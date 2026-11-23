"""base.py — peças comuns dos relatórios: contexto (API + memória + cache), resolvedor de colunas, exportação."""
from __future__ import annotations

import io
import json
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

import github_cache as gh
from belle_core import BelleClient, slug

COD_ESTAB = 1


def norm(k) -> str:
    t = unicodedata.normalize("NFD", str(k).lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t)


def num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def data(serie) -> pd.Series:
    return pd.to_datetime(serie, dayfirst=True, errors="coerce")


def whatsapp(tel) -> str:
    d = re.sub(r"\D", "", str(tel or ""))
    if len(d) in (10, 11):
        d = "55" + d
    return f"https://wa.me/{d}" if len(d) >= 12 else ""


def to_df(registros: list) -> pd.DataFrame:
    if not registros:
        return pd.DataFrame()
    df = pd.json_normalize(registros, sep="_")
    for c in df.columns:     # listas/dicts viram texto (cabem em planilha)
        if df[c].map(lambda x: isinstance(x, (list, dict))).any():
            df[c] = df[c].map(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else x)
    return df


class CampoAusente(Exception):
    def __init__(self, campo, colunas):
        super().__init__(f"Não achei a coluna '{campo}'")
        self.campo, self.colunas = campo, list(colunas)


@dataclass
class P:   # parâmetro de tela de um relatório
    key: str
    label: str
    kind: str = "num"          # num | select | text | bool
    default: object = 0
    options: list = field(default_factory=list)
    help: str = ""


@dataclass
class Resultado:
    sheets: dict
    kpis: list = field(default_factory=list)
    notas: list = field(default_factory=list)
    pdf: list | None = None    # [(titulo, DataFrame)] — padrão: 1ª aba

    @property
    def principal(self):
        return next(iter(self.sheets.values())) if self.sheets else pd.DataFrame()


class Ctx:
    """Tudo que um relatório precisa: tokens, unidades, memória de chamadas, ajustes de colunas, cache GitHub."""

    def __init__(self, tokens, unidades, memo=None, overrides=None, hoje=None):
        self.tokens, self.unidades = tokens, list(unidades)
        self.memo = memo if memo is not None else {}
        self.overrides = overrides or {}
        self.hoje = hoje or date.today()
        self.erros: dict = {}
        self._cli: dict = {}

    def cli(self, u) -> BelleClient:
        if u not in self._cli:
            t = self.tokens.get(slug(u))
            if not t:
                raise RuntimeError(f"Sem token da unidade {u} nos secrets.")
            self._cli[u] = BelleClient(t)
        return self._cli[u]

    def get(self, u, caminho, params=None) -> list:
        k = (u, caminho, tuple(sorted((params or {}).items())))
        if k not in self.memo:
            self.memo[k] = self.cli(u).get(caminho, params)
        return self.memo[k]

    def get_janelas(self, u, caminho, ini, fim, dias, nome_ini="dtInicio", nome_fim="dtFim", extra=None):
        out = []
        for a, b in BelleClient.janelas(ini, fim, dias):
            out += self.get(u, caminho, {**(extra or {}), nome_ini: BelleClient.d(a), nome_fim: BelleClient.d(b)})
        return out

    def col(self, df, campo, cands, frag=None, obrigatorio=True):
        ov = self.overrides.get(campo)
        if ov in df.columns:
            return ov
        reais = {norm(c): c for c in df.columns}
        for c in cands:
            if c in reais:
                return reais[c]
        if frag:
            for n, r in reais.items():
                if frag in n:
                    return r
        if obrigatorio:
            raise CampoAusente(campo, df.columns)
        return None

    def paralelo(self, fn) -> dict:
        """Roda fn(unidade) em paralelo (um token por unidade = limites independentes). Erros vão para self.erros."""
        out = {}
        with ThreadPoolExecutor(max_workers=max(1, len(self.unidades))) as ex:
            fut = {u: ex.submit(fn, u) for u in self.unidades}
            for u, f in fut.items():
                try:
                    out[u] = f.result()
                except CampoAusente:
                    raise
                except Exception as e:
                    self.erros[u] = str(e)
        return out

    # ── base de clientes (cadastro) — lenta (100 por página); fica salva no GitHub ──
    def clientes(self, u, usar_salva=True) -> pd.DataFrame:
        k = ("__clientes__", u)
        if k in self.memo:
            return self.memo[k]
        caminho = f"base/clientes_{slug(u)}.json"
        if usar_salva:
            try:
                salvo = gh.ler(caminho)
                if salvo:
                    self.memo[k] = pd.DataFrame(salvo["clientes"])
                    return self.memo[k]
            except Exception:
                pass
        brutos, pag = [], 0
        while pag < 600:
            lote = self.cli(u).get("clientes", {"pagina": pag, "codEstab": COD_ESTAB})
            if not lote:
                break
            brutos += lote
            pag += 1
        df = to_df(brutos)
        if df.empty:
            self.memo[k] = pd.DataFrame(columns=["id", "nome", "celular"])
            return self.memo[k]
        m = {c: self.col(df, c, cands, frag, obr) for c, cands, frag, obr in [
            ("c_id", ["id", "codcliente", "codigocliente", "idcliente", "codigo"], None, True),
            ("c_nome", ["nome", "nomecliente"], None, True),
            ("c_cel", ["celular", "celular1", "telefone", "fone"], "celular", False),
            ("c_nasc", ["datanascimento", "nascimento", "dtnascimento"], "nasc", False),
            ("c_cidade", ["cidade", "ufcidade", "municipio"], None, False),
            ("c_email", ["email"], None, False)]}
        can = pd.DataFrame({
            "id": df[m["c_id"]].astype(str).str.strip(), "nome": df[m["c_nome"]].astype(str).str.strip(),
            "celular": df[m["c_cel"]].astype(str) if m["c_cel"] else "",
            "nascimento": df[m["c_nasc"]].astype(str) if m["c_nasc"] else "",
            "cidade": df[m["c_cidade"]].astype(str) if m["c_cidade"] else "",
            "email": df[m["c_email"]].astype(str) if m["c_email"] else ""}).replace({"None": "", "nan": ""})
        try:
            gh.salvar(caminho, {"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "clientes": can.to_dict("records")}, f"base de clientes {u}")
        except Exception as e:
            self.erros[f"{u} (salvar base)"] = str(e)
        self.memo[k] = can
        return can


# ───────────── exportação ─────────────
def _nome_aba(n, usados):
    n = re.sub(r"[\[\]\*\?/\\:]", "", str(n))[:31] or "Aba"
    base, i = n, 2
    while n in usados:
        n = f"{base[:28]}_{i}"
        i += 1
    usados.add(n)
    return n


def xlsx_bytes(sheets: dict, titulo="Relatório") -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    buf, usados = io.BytesIO(), set()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, df in (sheets or {"Vazio": pd.DataFrame({"aviso": ["sem dados"]})}).items():
            if df is None or df.empty:
                df = pd.DataFrame({"aviso": ["sem dados"]})
            aba = _nome_aba(nome, usados)
            df.to_excel(w, sheet_name=aba, index=False)
            ws = w.sheets[aba]
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="ED2A78")
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, 1):
                larg = max([len(str(col))] + [len(str(x)) for x in df[col].head(300)])
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, larg + 2), 55)
    return buf.getvalue()


def pdf_bytes(titulo: str, secoes: list, max_linhas=3000) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    est = getSampleStyleSheet()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    el = [Paragraph(f"<b>{titulo}</b>", est["Title"]),
          Paragraph(f"Léah Estética · gerado em {datetime.now():%d/%m/%Y %H:%M}", est["Normal"]), Spacer(1, 10)]
    for sub, df in secoes:
        el.append(Paragraph(f"<b>{sub}</b>", est["Heading3"]))
        if df is None or df.empty:
            el += [Paragraph("Sem registros.", est["Normal"]), Spacer(1, 8)]
            continue
        d = df.drop(columns=[c for c in df.columns if str(c).lower() == "whatsapp"], errors="ignore").head(max_linhas)
        cols = list(d.columns)[:9]
        larg = (landscape(A4)[0] - 48) / len(cols)
        linhas = [cols] + [[Paragraph(str(v)[:60], est["BodyText"].clone("x", fontSize=7, leading=8))
                            for v in r] for r in d[cols].itertuples(index=False)]
        t = Table(linhas, colWidths=[larg] * len(cols), repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ed2a78")),
                               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTSIZE", (0, 0), (-1, 0), 7),
                               ("GRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
                               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fdf2f7")])]))
        el += [t, Spacer(1, 10)]
        if len(df) > max_linhas:
            el.append(Paragraph(f"(mostrando {max_linhas} de {len(df)} linhas — veja o Excel para tudo)", est["Italic"]))
    doc.build(el)
    return buf.getvalue()


def salvar_resultado(chave: str, res: Resultado) -> str:
    sheets = {k: json.loads(v.to_json(orient="records", date_format="iso")) for k, v in res.sheets.items()}
    obj = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "kpis": res.kpis, "notas": res.notas, "sheets": sheets}
    return gh.salvar(f"relatorios/{chave}.json", obj, f"relatório {chave}")


def carregar_resultado(chave: str):
    o = gh.ler(f"relatorios/{chave}.json")
    if not o:
        return None
    return o["gerado_em"], Resultado({k: pd.DataFrame(v) for k, v in o["sheets"].items()},
                                     [tuple(x) for x in o["kpis"]], o["notas"])
