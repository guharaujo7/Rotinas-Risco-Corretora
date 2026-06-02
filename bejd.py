import os, sys, re, random, tkinter as tk, threading, time, tempfile, shutil, pythoncom, webbrowser
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader
import win32com.client as win32
from datetime import datetime, date
 
try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_OK = True
except Exception:
    sync_playwright = None
    PLAYWRIGHT_OK = False
 
LOGO_FILENAME = "itau-logo-png_seeklogo-74122.png"
 
def _parse_brl_to_decimal(raw: str):
    s = (raw or "").strip()
    if not s:
        return None
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in {".", ",", "-", "-.", "-,"}:
        return None
    last_dot = s.rfind(".")
    last_comma = s.rfind(",")
    sep_idx = max(last_dot, last_comma)
    try:
        if sep_idx == -1:
            int_part = re.sub(r"[^\d\-]", "", s)
            d = Decimal(int_part)
        else:
            decimal_sep = s[sep_idx]
            int_part = s[:sep_idx]
            dec_part = s[sep_idx + 1:]
            int_part_clean = re.sub(r"[^\d\-]", "", int_part)
            dec_part_clean = re.sub(r"[^\d]", "", dec_part)
            if not dec_part_clean:
                return None
            normalized = (int_part_clean if int_part_clean not in {"", "-"} else "0") + "." + dec_part_clean
            d = Decimal(normalized)
    except (InvalidOperation, ValueError):
        return None
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
 
 
def _format_brl_currency_from_decimal(d: Decimal) -> str:
    d2 = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sign = "-" if d2 < 0 else ""
    d_abs = abs(d2)
    s = f"{d_abs:.2f}"
    int_s, dec_s = s.split(".")
    int_s_grouped = "{:,}".format(int(int_s)).replace(",", ".")
    return f"R$ {sign}{int_s_grouped},{dec_s}"
 
 
def _format_brl_currency(raw: str) -> str:
    d = _parse_brl_to_decimal(raw)
    if d is None:
        return (raw or "").strip()
    return _format_brl_currency_from_decimal(d)
 
 
def _format_brl_plain_for_web(raw: str) -> str:
    d = _parse_brl_to_decimal(raw)
    if d is None:
        return ""
    d2 = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sign = "-" if d2 < 0 else ""
    d_abs = abs(d2)
    s = f"{d_abs:.2f}"
    int_s, dec_s = s.split(".")
    int_s_grouped = "{:,}".format(int(int_s)).replace(",", ".")
    return f"{sign}{int_s_grouped},{dec_s}"
 
 
def _bpm_round_rect(canvas: tk.Canvas, x1, y1, x2, y2, radius=18, **kwargs):
    try:
        x1, y1, x2, y2 = float(x1), float(y1), float(x2), float(y2)
    except Exception:
        return canvas.create_rectangle(x1, y1, x2, y2, **kwargs)
    if x2 <= x1 or y2 <= y1:
        return canvas.create_rectangle(x1, y1, x2, y2, **kwargs)
    r = min(radius, (x2 - x1) / 2.0, (y2 - y1) / 2.0)
    points = [
        x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r,
        x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2,
        x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r,
        x1, y2-r, x1, y1+r, x1, y1+r, x1, y1, x1+r, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)
 
 
class BPMUserCancelled(Exception):
    """Usuário cancelou a sequência de BPM na interface."""
 
 
BPM_CLIENT_DATA = {
    "Transdourada": {"CNPJ": "01259730000174", "PLATAFORMA": "2939", "AG": "1643", "CONTA": "99451-8"},
    "RPB": {"CNPJ": "07075892000139", "PLATAFORMA": "8973", "AG": "6627", "CONTA": "06471-7"},
    "Posto Arinos": {"CNPJ": "05798923000154", "PLATAFORMA": "8250", "AG": "1364", "CONTA": "98355-9"},
    "Brasnorte": {"CNPJ": "00514301000133", "PLATAFORMA": "8250", "AG": "1364", "CONTA": "98654-5"},
    "Mirian Varzea": {"CNPJ": "16519674000137", "PLATAFORMA": "8250", "AG": "1689", "CONTA": "05136-3"},
    "Mirian Cuiaba": {"CNPJ": "41240105000103", "PLATAFORMA": "8250", "AG": "1689", "CONTA": "58145-0"},
    "Petrocal": {"CNPJ": "12781233000158", "PLATAFORMA": "7948", "AG": "8251", "CONTA": "44190-6"},
    "Posto Sapucaia": {"CNPJ": "22787055000126", "PLATAFORMA": "0352", "AG": "1334", "CONTA": "57853-9"},
    "PetroMix": {"CNPJ": "05684913000198", "PLATAFORMA": "7948", "AG": "8251", "CONTA": "99886-3"},
    "Auto Posto M Timbozao": {"CNPJ": "04632746000179", "PLATAFORMA": "0352", "AG": "8296", "CONTA": "18100-4"},
    "PetroVel": {"CNPJ": "01294927000144", "PLATAFORMA": "7948", "AG": "8251", "CONTA": "99887-1"},
    "Posto Gasol Timbo III": {"CNPJ": "32179707000101", "PLATAFORMA": "0352", "AG": "8296", "CONTA": "06655-1"},
    "Posto Timbozao Itaperuna": {"CNPJ": "25032853000136", "PLATAFORMA": "0352", "AG": "1334", "CONTA": "49413-3"},
    "Posto Pioneiro": {"CNPJ": "23184831000166", "PLATAFORMA": "0352", "AG": "5255", "CONTA": "12888-5"},
}
 
LIMITE_CLIENT_URLS = {
    "Transdourada":  "https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1R10K6&tppes=J&subgrupo=01259730000174&plataforma=2939&ambiente=PRD#/omni/",
    "RPB":           "https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1D1AK3&tppes=J&subgrupo=26727190000137&plataforma=8973&ambiente=PRD#/omni/",
    "Posto Arinos":  "https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1H1UK3&tppes=J&subgrupo=00514301000133&plataforma=8250&ambiente=PRD",
    "Mirian Varzea": "https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1H1UK3&tppes=J&subgrupo=16519674000137&plataforma=8250&ambiente=PRD",
    "Petrocal":      "https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1H1UK3&tppes=J&subgrupo=12781233000158&plataforma=7948&ambiente=PRD",
    "Posto Sapucaia":"https://digital.itau/CF_Digital/IntegracaoQK/IntegracaoQK/redir.aspx?usuario=987400146&senha=1H1UK3&tppes=J&subgrupo=32179707000101&plataforma=0352&ambiente=PRD#/omni/",
}
 
LIMITE_SHARED_RESULTS = {
    "Posto Arinos":   ["Brasnorte"],
    "Mirian Varzea":  ["Mirian Cuiaba"],
    "Petrocal":       ["PetroMix", "PetroVel"],
    "Posto Sapucaia": ["Auto Posto M Timbozao", "Posto Gasol Timbo III", "Posto Timbozao Itaperuna", "Posto Pioneiro"],
}
 
RE_SPACES = re.compile(r"\s+")
RE_CNPJ = re.compile(r"(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})")
RE_CNPJ_LABEL = re.compile(r"CNPJ[^\d]*([\d./-]{14,20})", re.IGNORECASE)
RE_VALOR_1 = re.compile(r"Valor\s+da\s+Opera[cç][aã]o[^\d]*(R\$\s*[\d\.,]+)", re.IGNORECASE)
RE_VALOR_2 = re.compile(r"R\$\s*[\d\.]{1,},\d{2}")
RE_VALOR_3 = re.compile(r"(R\$\s*[\d\.,]{3,})")
RE_PLATAFORMA = re.compile(r"Plataforma\D{0,20}(\d{4})", re.IGNORECASE)
RE_REGIAO_PLAT = re.compile(r"Regi[aã]o(?:\s+da)?\s+Plataforma\D{0,20}(\d{2})", re.IGNORECASE)
RE_REGIAO = re.compile(r"Regi[aã]o\D{0,20}(\d{2})", re.IGNORECASE)
RE_SPREAD_1 = re.compile(r"(?:Taxa\/Spread|Spread)\D{0,25}([\d\.,]+)\s*%?", re.IGNORECASE)
RE_SPREAD_2 = re.compile(r"Spread\s+m[íi]nimo\D{0,25}([\d\.,]+)\s*%?", re.IGNORECASE)
RE_PRAZO_MIN_1 = re.compile(r"Prazo\s+M[ií]n[ií]mo\s+NF\D{0,25}(\d{1,4})", re.IGNORECASE)
RE_PRAZO_MAX_1 = re.compile(r"Prazo\s+M[áa]ximo\s+NF\D{0,25}(\d{1,4})", re.IGNORECASE)
RE_MODALIDADE_1 = re.compile(r"Modalidade[:\s]+([^\n\.;]+)", re.IGNORECASE)
RE_AUTORIZADAS_CSV = re.compile(r"\bautorizadas\s+csv\b", re.IGNORECASE)
RE_AUTORIZADAS_SISPAG = re.compile(r"\bautorizadas\s+sispag\b", re.IGNORECASE)
RE_RAZAO_1 = re.compile(r"Raz[aã]o\s+Social[:\s]*([^\n]+)", re.IGNORECASE)
RE_RAZAO_ATE_PROXIMO = re.compile(
    r"Raz[aã]o\s+Social\s*[:\s]*\s*(.+?)(?=\n\s*(?:CNPJ|Plataforma|Modalidade|Regi[aã]o|Conta\s+Corrente|Valor\s+da|Spread|Prazo)\b)",
    re.IGNORECASE | re.DOTALL,
)
RE_CONTA_AG_CC = re.compile(r"\b(\d{4})\s*[/\s]\s*(\d{3,10}(?:[-‑–—−]\d)?)\b")
RE_CONTA_LABEL = re.compile(r"conta\s+corrente(?:\s+do\s+cliente)?\s*[:\s-]*", re.IGNORECASE)
RE_LIQ_CRED = re.compile(r"Cr[ée]dito\s+em\s+CC", re.IGNORECASE)
RE_PREMIO = re.compile(r"com\s+pr[êe]mio", re.IGNORECASE)
 
REGIAO_TRADER_ESPEC = {
    "21": ("Debora", "Vinicios Luz"),
    "22": ("Thiago", "Paula Costa"),
    "23": ("Thiago", "Paula Costa"),
    "24": ("Gabriel", "Luiz Gustavo Sarmento"),
    "25": ("Giovanna", "Lucas Capeli"),
    "26": ("Gabriel", "Vinicios Luz"),
    "28": ("Thiago", "Paula Costa"),
    "29": ("Debora", "Renata Leviski"),
    "30": ("Adriana", "Luiz Gustavo Sarmento"),
    "32": ("Debora", "Renata Leviski"),
    "33": ("Giovanna", "Lucas Capeli"),
}
 
 
def app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
 
 
def resource_path(p):
    base = getattr(sys, "_MEIPASS", app_base_dir())
    return os.path.join(base, p)
 
 
def only_digits(s):
    return re.sub(r"\D", "", s or "")
 
 
def format_cnpj(c):
    return only_digits(c or "")
 
 
def normalize_hyphens(s):
    return (s or "").replace("‑", "-").replace("–", "-").replace("—", "-").replace("−", "-")
 
 
def normalize_razao_social_text(s):
    if not s:
        return s
    t = normalize_hyphens(s)
    t = t.replace("\u00a0", " ").replace("\u200b", "").replace("\ufeff", "")
    t = RE_SPACES.sub(" ", t).strip()
    return t
 
 
def normalize_text_variants(t):
    return RE_SPACES.sub(" ", t).strip(), re.sub(r"\s+", "", t or "")
 
 
def extract_text_from_pdf(p):
    r = PdfReader(p, strict=False)
    layout_parts = []
    plain_parts = []
    for pg in r.pages:
        layout = ""
        plain = ""
        try:
            try:
                layout = pg.extract_text(extraction_mode="layout") or ""
            except TypeError:
                layout = pg.extract_text() or ""
        except Exception:
            layout = ""
        try:
            plain = pg.extract_text() or ""
        except Exception:
            plain = layout
        layout_parts.append(layout)
        plain_parts.append(plain)
    return "\n".join(layout_parts), "\n".join(plain_parts)
 
 
def find_first(P, t):
    for pat in P:
        m = pat.search(t)
        if m:
            return m
    return None
 
 
def normalize_modalidade(m):
    s = (m or "").strip()
    s = RE_SPACES.sub(" ", s)
    sl = s.lower()
    if "sispag" in sl.replace(" ", ""):
        return "Autorizadas SISPAG"
    if RE_AUTORIZADAS_SISPAG.search(s):
        return "Autorizadas SISPAG"
    if RE_AUTORIZADAS_CSV.search(s):
        return "Autorizadas CSV"
    if re.search(r"\bsispag\b", s, re.IGNORECASE):
        return "Autorizadas SISPAG"
    if re.search(r"\bcsv\b", s, re.IGNORECASE):
        return "Autorizadas CSV"
    return s
 
 
def infer_troca_from_modalidade(m):
    m = (m or "").strip()
    ml = m.lower()
    compact = ml.replace(" ", "")
    if "sispag" in compact:
        return "Sispag"
    if re.search(r"\bsispag\b", m, re.IGNORECASE):
        return "Sispag"
    if re.search(r"\bcsv\b", m, re.IGNORECASE) or "csv" in compact:
        return "CSV"
    return "CSV"
 
 
def normalize_percent_br(p):
    if not p:
        return ""
    v = p.strip().replace(" ", "").replace(",", ".")
    try:
        n = float(re.sub(r"[^\d\.]", "", v))
        return f"{int(n)},00% a.a." if n.is_integer() else f"{n:.2f}".replace(".", ",") + "% a.a."
    except Exception:
        v = p.strip().replace(".", ",")
        if not v.endswith("%"):
            v += "%"
        return v + ("" if "a.a" in v.lower() else " a.a.")
 
 
def sanitize_razao_social_display(s: str) -> str:
    if not s:
        return s
    s = normalize_razao_social_text(s)
    s = re.sub(r"\s*\(\s*/[^)]+\)", "", s)
    s = re.sub(r"\s*\(\s*[^)]*[/\\][^)]+\)", "", s)
    s = re.sub(r"\s*\(\s*[A-Za-z0-9]{14,}\s*\)", "", s)
    s = re.split(r"\bCNPJ\b", s, maxsplit=1, flags=re.I)[0]
    s = re.sub(r"\s*\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\s*$", "", s)
    s = RE_SPACES.sub(" ", s).strip(" :-–—.,;")
    return s
 
 
def _razao_end_markers():
    return [
        re.compile(r"\bCNPJ\b", re.I),
        re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}"),
        re.compile(r"\bPlataforma\b", re.I),
        re.compile(r"\bModalidade\b", re.I),
        re.compile(r"\bRegi[aã]o\b", re.I),
        re.compile(r"Conta\s+Corrente", re.I),
        re.compile(r"\bValor\s+da\s+Opera", re.I),
    ]
 
 
def _extract_razao_between_labels(t: str) -> str | None:
    m = re.search(r"raz[aã]o\s+social\s*[:\s]*", t, re.I)
    if not m:
        return None
    rest = t[m.end():]
    end = len(rest)
    for pat in _razao_end_markers():
        rm = pat.search(rest)
        if rm is not None and rm.start() < end:
            end = rm.start()
    chunk = rest[:end].strip()
    if "\n\n" in chunk:
        chunk = chunk.split("\n\n")[0].strip()
    if len(chunk) < 2:
        return None
    return chunk
 
 
def _is_razao_stop_line(ln: str) -> bool:
    s = (ln or "").strip()
    if not s:
        return False
    low = s.lower()
    if re.match(r"^\s*cnpj\b", low): return True
    if re.match(r"^\s*plataforma\b", low): return True
    if re.match(r"^\s*modalidade\b", low): return True
    if re.match(r"^\s*regi[aã]o\b", low): return True
    if re.match(r"^\s*conta\s+corrente", low): return True
    if re.match(r"^\s*valor\s+da\s+opera", low): return True
    if re.match(r"^\s*spread\b", low): return True
    if re.match(r"^\s*prazo\b", low): return True
    if re.match(r"^\s*liquida", low): return True
    if len(only_digits(s)) >= 14 and len(s) < 40: return True
    return False
 
 
def _extract_razao_line_block(lines: list[str]) -> str | None:
    for i, line in enumerate(lines):
        if not re.search(r"raz[aã]o\s+social", line, re.I):
            continue
        m = re.search(r"raz[aã]o\s+social\s*[:\s]*\s*(.+)$", line, re.I)
        if m:
            rest = m.group(1).strip()
            if rest and len(rest) >= 2 and not re.match(r"^[\d./\s\-]+$", rest):
                if not _is_razao_stop_line(rest):
                    return rest
        parts = []
        for j in range(i + 1, min(i + 25, len(lines))):
            ln = lines[j].strip()
            if not ln:
                if parts: break
                continue
            if _is_razao_stop_line(ln): break
            if re.search(r"[A-Za-zÀ-ÿ0-9]", ln):
                parts.append(ln)
        if parts:
            return " ".join(parts)
    return None
 
 
def _extract_razao_regex_blocks(t: str) -> str | None:
    m = RE_RAZAO_ATE_PROXIMO.search(t)
    if m:
        c = RE_SPACES.sub(" ", m.group(1).strip()).strip()
        if len(c) >= 2: return c
    m = RE_RAZAO_1.search(t)
    if m:
        c = m.group(1).strip()
        if not re.match(r"^[(/\s]*\d", c):
            c = RE_SPACES.sub(" ", c).strip()
            if len(c) >= 2: return c
    m = re.search(r"Raz[aã]o\s+Social\s*[:\s]*\s*(.+?)(?=\n\s*\n|\Z)", t, re.I | re.DOTALL)
    if m:
        c = RE_SPACES.sub(" ", m.group(1).strip()).strip()
        if len(c) >= 2: return c
    return None
 
 
def extract_razao_social(t, lines, tn, tc, t_plain=None):
    for block in ([t] + ([t_plain] if t_plain else [])):
        if not block: continue
        r = _extract_razao_between_labels(block)
        if r: return sanitize_razao_social_display(r)
        r = _extract_razao_line_block(block.splitlines())
        if r: return sanitize_razao_social_display(r)
        r = _extract_razao_regex_blocks(block)
        if r: return sanitize_razao_social_display(r)
    idx = tc.lower().find("razaosocial")
    if idx != -1:
        tail = tc[idx: idx + 220]
        m2 = re.search(r"razaosocial[:\-]?(.*?)(cnpj|\d{11,})", tail, re.I)
        if m2:
            c = m2.group(1).strip(" :-")
            if len(c) >= 4: return sanitize_razao_social_display(c)
    return None
 
 
def extract_cnpj(t, tn, tc):
    m = RE_CNPJ_LABEL.search(t)
    if m: return m.group(1).strip()
    m = RE_CNPJ.search(t)
    if m: return m.group(1).strip()
    m2 = re.search(r"cnpj.*?(\d{14})", tc, re.IGNORECASE)
    if m2: return m2.group(1)
    return None
 
 
def trader_espec_from_regiao(reg: str):
    reg = (reg or "").strip()
    d = only_digits(reg)
    if len(d) >= 2:
        key = d[-2:]
        return REGIAO_TRADER_ESPEC.get(key, ("", ""))
    m = RE_REGIAO_PLAT.search(reg) or RE_REGIAO.search(reg)
    if m: return REGIAO_TRADER_ESPEC.get(m.group(1), ("", ""))
    return ("", "")
 
 
def premio_str(prem: str) -> str:
    if re.search(r"\bcom\s+pr", prem or "", re.IGNORECASE):
        return "com prêmio"
    return "sem prêmio"
 
 
def _sanitize_conta_corrente(val: str) -> str:
    val = (val or "").strip()
    if not val: return val
    for splitter in (r"\bCNPJ\b", r"\bPlataforma\b", r"\bModalidade\b", r"\bRegi[aã]o\b", r"\bRaz[aã]o\s+Social\b", r"\bValor\s+da\s+Opera"):
        val = re.split(splitter, val, maxsplit=1, flags=re.I)[0].strip()
    val = re.sub(r"\s*\([^)]*[/\\][^)]*\)", "", val)
    val = re.sub(r"\s*\(\s*[A-Za-z0-9]{12,}\s*\)", "", val)
    val = RE_SPACES.sub(" ", val).strip(" :-–—.,;")
    return val
 
 
def _conta_stop_line(ln: str) -> bool:
    low = (ln or "").strip().lower()
    if not low: return True
    return bool(re.match(r"^\s*(?:cnpj|plataforma|modalidade|regi[aã]o|raz[aã]o\s+social|valor|spread|prazo)\b", low))
 
 
def extract_conta_corrente(lines, tn, tc):
    for hay in (tn, tc):
        if not hay: continue
        m = RE_CONTA_AG_CC.search(hay)
        if m: return f"{m.group(1)} / {m.group(2)}"
    for i, line in enumerate(lines):
        if not re.search(r"conta\s+corrente", line, re.I): continue
        mm = RE_CONTA_LABEL.search(line)
        if not mm: continue
        rest = line[mm.end():].strip()
        if rest and not _conta_stop_line(rest): return _sanitize_conta_corrente(rest)
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if nxt and not _conta_stop_line(nxt): return _sanitize_conta_corrente(nxt)
    m = RE_CONTA_LABEL.search(tn)
    if m:
        tail = tn[m.end(): m.end() + 120]
        tail = tail.split("\n")[0].strip()
        if tail: return _sanitize_conta_corrente(tail)
    return None
 
 
def extract_plataforma(t, tn, tc):
    m = RE_PLATAFORMA.search(tn) or RE_PLATAFORMA.search(t)
    return m.group(1) if m else None
 
 
def extract_regiao_plataforma(t, tn, tc):
    m = RE_REGIAO_PLAT.search(tn) or RE_REGIAO_PLAT.search(t)
    if m: return m.group(1)
    m = RE_REGIAO.search(tn) or RE_REGIAO.search(t)
    return m.group(1) if m else None
 
 
def extract_valor_operacao(t, tn, tc):
    m = find_first([RE_VALOR_1, RE_VALOR_2, RE_VALOR_3], tn)
    return m.group(1).strip() if m else None
 
 
def extract_spread(t, tn, tc):
    m = find_first([RE_SPREAD_1, RE_SPREAD_2], tn)
    return m.group(1).strip() if m else None
 
 
def extract_prazo_minimo_nf(t, tn, tc):
    m = RE_PRAZO_MIN_1.search(tn)
    return m.group(1) if m else None
 
 
def extract_prazo_maximo_nf(t, tn, tc):
    m = RE_PRAZO_MAX_1.search(tn)
    return m.group(1) if m else None
 
 
def extract_modalidade(t, tn, tc):
    for hay in (tn, t, tc):
        if not hay: continue
        m = RE_MODALIDADE_1.search(hay)
        if m: return normalize_modalidade(m.group(1).strip())
    return None
 
 
class HomeFrame(ttk.Frame):
    CLIENTS = [
        "Posto Arinos", "Auto Posto M Timbozao", "Posto Sapucaia", "Brasnorte",
        "Mirian Varzea", "Mirian Cuiaba", "RPB", "Petrocal", "PetroMix",
        "Transdourada", "PetroVel", "Posto Gasol Timbo III",
        "Posto Timbozao Itaperuna", "Posto Pioneiro",
    ]
 
    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self._mode_dialog = None
        self._mode_overlay = None
        self._dialog_body = None
        self._selected_clients = {}
        self._confirm_btn = None
        self._cred_continue_btn = None
        self._bpm_func_var = None
        self._bpm_senha_var = None
        self._build_ui()
 
    def _build_ui(self):
        bg = self.COLORS["bg"]

        outer = tk.Frame(self, bg=bg)
        outer.pack(fill="both", expand=True)
        center = tk.Frame(outer, bg=bg)
        center.place(relx=0.5, rely=0.5, anchor="center")

        # Logo + titulo
        logo_row = tk.Frame(center, bg=bg)
        logo_row.pack()
        self.logo_label = tk.Label(logo_row, bg=bg)
        lp = resource_path(LOGO_FILENAME)
        if os.path.exists(lp):
            try:
                img = tk.PhotoImage(file=lp)
                img2 = img.subsample(2, 2)
                self._logo_ref = img2
                self.logo_label.configure(image=img2)
                self.logo_label.pack(side="left", padx=(0, 12))
            except Exception:
                pass
        title_col = tk.Frame(logo_row, bg=bg)
        title_col.pack(side="left")
        title_row = tk.Frame(title_col, bg=bg)
        title_row.pack(anchor="w")
        tk.Label(title_row, text="Mesa", bg=bg, fg="#f2f2f2",
                 font=("Segoe UI", 20, "bold")).pack(side="left")
        tk.Label(title_row, text="Itaú", bg=bg, fg="#ec7000",
                 font=("Segoe UI", 20, "bold")).pack(side="left", padx=(6, 0))
        tk.Label(title_col, text="Risco Sacado", bg=bg, fg="#ec7000",
                 font=("Segoe UI", 10)).pack(anchor="w")

        # Divisor
        tk.Frame(center, bg="#1e1e1e", height=1).pack(fill="x", pady=(24, 20))

        tk.Label(center, text="MODULOS", bg=bg, fg="#2e2e2e",
                 font=("Segoe UI", 8, "bold")).pack(pady=(0, 12))

        # Grid 2 colunas usando grid geometry manager
        cards_frame = tk.Frame(center, bg=bg)
        cards_frame.pack()
        cards_frame.columnconfigure(0, weight=1, uniform="c")
        cards_frame.columnconfigure(1, weight=1, uniform="c")

        def _make_card(parent, col, title, subtitle, accent, command):
            border = "#1a3d28" if accent else "#1f1f1f"
            card_bg = "#111c15" if accent else "#161616"
            bar_color = "#2fb875" if accent else "#242424"
            title_fg = "#e8e8e8" if accent else "#b0b0b0"
            sub_fg = "#2a7a50" if accent else "#3a3a3a"

            wrap = tk.Frame(parent, bg=card_bg, highlightthickness=1,
                            highlightbackground=border, cursor="hand2")
            wrap.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 6) if col == 0 else (6, 0))

            tk.Frame(wrap, bg=bar_color, height=2).pack(fill="x")

            body = tk.Frame(wrap, bg=card_bg, padx=16, pady=14)
            body.pack(fill="both", expand=True)

            tk.Label(body, text=title, bg=card_bg, fg=title_fg,
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(body, text=subtitle, bg=card_bg, fg=sub_fg,
                     font=("Segoe UI", 8), wraplength=140,
                     justify="left").pack(anchor="w", pady=(4, 0))

            def on_enter(_):
                hbg = "#152416" if accent else "#1c1c1c"
                hbd = "#2fb875" if accent else "#2e2e2e"
                wrap.configure(bg=hbg, highlightbackground=hbd)
                body.configure(bg=hbg)
                for w in body.winfo_children():
                    try: w.configure(bg=hbg)
                    except Exception: pass

            def on_leave(_):
                wrap.configure(bg=card_bg, highlightbackground=border)
                body.configure(bg=card_bg)
                for w in body.winfo_children():
                    try: w.configure(bg=card_bg)
                    except Exception: pass

            for widget in (wrap, body):
                widget.bind("<Button-1>", lambda _e, cmd=command: cmd())
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)
            for child in body.winfo_children():
                child.bind("<Button-1>", lambda _e, cmd=command: cmd())
                child.bind("<Enter>", on_enter)
                child.bind("<Leave>", on_leave)

        _make_card(cards_frame, 0, "Cadastro Share", "Analise pdf - Share",
                   accent=True, command=lambda: self.controller.show_frame("Share"))
        _make_card(cards_frame, 1, "BPM", "Abertura das solicitações",
                   accent=False, command=self.open_bpm_mode_dialog)

        # Limites Invertido full-width
        lim_bg = "#0f1a2a"
        lim_border = "#1a2e45"
        lim = tk.Frame(center, bg=lim_bg, highlightthickness=1,
                       highlightbackground=lim_border, cursor="hand2")
        lim.pack(fill="x", pady=(10, 0))
        tk.Frame(lim, bg="#1e4060", height=2).pack(fill="x")
        lim_inner = tk.Frame(lim, bg=lim_bg, padx=16, pady=12)
        lim_inner.pack(fill="x")
        tk.Label(lim_inner, text="Limites Invertido", bg=lim_bg, fg="#7ec8e3",
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        tk.Label(lim_inner, text="Consulta do LTC e limites disponiveis",
                 bg=lim_bg, fg="#1e4060", font=("Segoe UI", 8)).pack(side="left", padx=(10, 0))

        def lim_cmd(): self.controller.show_frame("LimitesInvertido")
        def lim_enter(_):
            lim.configure(bg="#0d1f35", highlightbackground="#2a4d70")
            lim_inner.configure(bg="#0d1f35")
            for w in lim_inner.winfo_children():
                try: w.configure(bg="#0d1f35")
                except Exception: pass
        def lim_leave(_):
            lim.configure(bg=lim_bg, highlightbackground=lim_border)
            lim_inner.configure(bg=lim_bg)
            for w in lim_inner.winfo_children():
                try: w.configure(bg=lim_bg)
                except Exception: pass
        for widget in (lim, lim_inner):
            widget.bind("<Button-1>", lambda _e: lim_cmd())
            widget.bind("<Enter>", lim_enter)
            widget.bind("<Leave>", lim_leave)
        for child in lim_inner.winfo_children():
            child.bind("<Button-1>", lambda _e: lim_cmd())
            child.bind("<Enter>", lim_enter)
            child.bind("<Leave>", lim_leave)

        # Rodape
        tk.Label(center, text="Mesa de Operacao  ·  Risco Sacado  ·  Middle",
                 bg=bg, fg="#242424", font=("Segoe UI", 8)).pack(pady=(26, 0))
 
    def open_bpm_mode_dialog(self):
        if self._mode_overlay is not None and self._mode_overlay.winfo_exists():
            self._mode_overlay.lift()
            return
        self.update_idletasks()
        overlay = tk.Frame(self, bg="#121212", bd=0, highlightthickness=0)
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        overlay.lift()
        self._mode_overlay = overlay
        dialog = tk.Frame(overlay, bg="#1a1a1a", highlightthickness=1, highlightbackground="#2e2e2e", bd=0, padx=24, pady=28)
        dialog.place(relx=0.5, rely=0.5, anchor="center", width=468)
        self._mode_dialog = dialog
        self._dialog_body = tk.Frame(dialog, bg="#1a1a1a")
        self._dialog_body.pack(fill="both", expand=True)
        self._render_mode_step()
        overlay.bind("<Button-1>", lambda *_: self.close_bpm_mode_dialog())
        dialog.bind("<Button-1>", lambda e: "break")
 
    def close_bpm_mode_dialog(self):
        if self._mode_overlay is not None and self._mode_overlay.winfo_exists():
            self._mode_overlay.destroy()
        self._mode_dialog = None
        self._mode_overlay = None
        self._dialog_body = None
        self._selected_clients = {}
        self._confirm_btn = None
        self._cred_continue_btn = None
        self._bpm_func_var = None
        self._bpm_senha_var = None
 
    def _clear_dialog_body(self):
        if self._dialog_body is None: return
        for child in self._dialog_body.winfo_children():
            child.destroy()
 
    def _animate_dialog_step(self):
        if (self._mode_dialog is None or not self._mode_dialog.winfo_exists()
                or self._dialog_body is None or not self._dialog_body.winfo_exists()):
            return
        self._mode_dialog.update_idletasks()
        start_y = 10
        steps = 12
        duration = 300
        step_ms = max(1, duration // steps)
        self._dialog_body.pack_forget()
        self._dialog_body.place(x=0, y=start_y, relwidth=1, relheight=1)
 
        def tick(i=0):
            if self._dialog_body is None or not self._dialog_body.winfo_exists(): return
            y = start_y - int((start_y * i) / steps)
            self._dialog_body.place_configure(y=y)
            if i < steps:
                self._dialog_body.after(step_ms, lambda: tick(i + 1))
            else:
                self._dialog_body.place_forget()
                self._dialog_body.pack(fill="both", expand=True)
        tick()
 
    def _render_mode_step(self):
        self._clear_dialog_body()
        card = self._dialog_body
        active_bg = "#1f1f1f"
        active_hover = "#292929"
        active_border = "#343434"
        accent = "#2fb875"
        disabled_bg = "#171717"
        disabled_fg = "#4d4d4d"
 
        tk.Label(card, text="Selecione o modo", bg="#1a1a1a", fg="#f2f2f2", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=24, pady=(28, 0))
        tk.Label(card, text="Escolha como deseja operar.", bg="#1a1a1a", fg="#8c8c8c", font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(2, 14))
 
        list_box = tk.Frame(card, bg="#1a1a1a")
        list_box.pack(fill="x", padx=24, pady=(0, 6))
 
        active_card = tk.Frame(list_box, bg=active_bg, highlightthickness=1, highlightbackground=active_border, cursor="hand2", padx=20, pady=18)
        active_card.pack(fill="x", pady=(0, 10))
        head = tk.Frame(active_card, bg=active_bg)
        head.pack(fill="x")
        tk.Label(head, text="Invertido", bg=active_bg, fg="#ececec", font=("Segoe UI", 10, "normal")).pack(side="left")
        arrow = tk.Label(head, text="->", bg=active_bg, fg=accent, font=("Segoe UI", 11, "bold"))
        arrow.pack(side="right")
        arrow.pack_forget()
        tk.Label(active_card, text="Operacao disponivel agora para execucao.", bg=active_bg, fg="#9a9a9a", font=("Segoe UI", 9), justify="left", wraplength=380).pack(anchor="w", pady=(6, 0))
 
        def on_active_enter(_):
            active_card.configure(bg=active_hover, highlightbackground=accent)
            head.configure(bg=active_hover)
            for w in head.winfo_children(): w.configure(bg=active_hover)
            for w in active_card.winfo_children():
                if isinstance(w, tk.Label): w.configure(bg=active_hover)
            if not arrow.winfo_ismapped(): arrow.pack(side="right")
 
        def on_active_leave(_):
            active_card.configure(bg=active_bg, highlightbackground=active_border)
            head.configure(bg=active_bg)
            for w in head.winfo_children(): w.configure(bg=active_bg)
            for w in active_card.winfo_children():
                if isinstance(w, tk.Label): w.configure(bg=active_bg)
            if arrow.winfo_ismapped(): arrow.pack_forget()
 
        for widget in (active_card, head):
            widget.bind("<Enter>", on_active_enter)
            widget.bind("<Leave>", on_active_leave)
            widget.bind("<Button-1>", lambda _e: self._render_credenciais_step())
        for widget in active_card.winfo_children():
            widget.bind("<Enter>", on_active_enter)
            widget.bind("<Leave>", on_active_leave)
            widget.bind("<Button-1>", lambda _e: self._render_credenciais_step())
 
        def disabled_option(title):
            row = tk.Frame(list_box, bg=disabled_bg, padx=20, pady=18, cursor="no")
            row.pack(fill="x", pady=(0, 10))
            title_row = tk.Frame(row, bg=disabled_bg)
            title_row.pack(fill="x")
            tk.Label(title_row, text=title, bg=disabled_bg, fg=disabled_fg, font=("Segoe UI", 10, "normal")).pack(side="left")
            badge = tk.Frame(title_row, bg="#333333", bd=0, padx=10, pady=2)
            badge.pack(side="left", padx=(8, 0))
            tk.Label(badge, text="EM BREVE", bg="#333333", fg="#737373", font=("Segoe UI", 7, "bold")).pack()
            tk.Label(row, text="Modo indisponivel no momento.", bg=disabled_bg, fg=disabled_fg, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(6, 0))
 
        disabled_option("Via Mesa")
        disabled_option("Nova Plataforma")
        ttk.Button(card, text="Fechar", style="Dark.TButton", command=self.close_bpm_mode_dialog).pack(anchor="e", padx=24, pady=(6, 24))
        self._animate_dialog_step()
 
    @staticmethod
    def _bpm_only_digits(s: str) -> str:
        return "".join(c for c in (s or "") if c.isdigit())

    def _refresh_cred_continue_state(self):
        btn = getattr(self, "_cred_continue_btn", None)
        if btn is None or not btn.winfo_exists():
            return
        f = self._bpm_only_digits((self._bpm_func_var.get() if self._bpm_func_var else "") or "")
        s = self._bpm_only_digits((self._bpm_senha_var.get() if self._bpm_senha_var else "") or "")
        ok = bool(f) and (1 <= len(s) <= 6)
        if ok:
            btn.configure(state="normal", bg="#2fb875", fg="#ffffff", activebackground="#2aa66a", cursor="hand2")
        else:
            btn.configure(state="disabled", bg="#242424", fg="#6f6f6f", activebackground="#242424", cursor="no")
 
    def _credenciais_continue(self):
        f = self._bpm_only_digits((self._bpm_func_var.get() if self._bpm_func_var else "") or "")
        s = self._bpm_only_digits((self._bpm_senha_var.get() if self._bpm_senha_var else "") or "")
        if not f:
            messagebox.showwarning("Identificação", "Funcional deve conter apenas números.")
            return
        if not (1 <= len(s) <= 6):
            messagebox.showwarning("Identificação", "Senha: apenas números, com até 6 dígitos.")
            return
        self.controller.bpm_funcional = f
        self.controller.bpm_password = s
        self._render_clients_step()
 
    def _render_credenciais_step(self):
        self._clear_dialog_body()
        card = self._dialog_body
        fv = self._bpm_only_digits(getattr(self.controller, "bpm_funcional", "") or "")
        sv = self._bpm_only_digits(getattr(self.controller, "bpm_password", "") or "")[:6]
        self._bpm_func_var = tk.StringVar(value=fv)
        self._bpm_senha_var = tk.StringVar(value=sv)
 
        tk.Label(
            card, text="Identificação",
            bg="#1a1a1a", fg="#f2f2f2", font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=24, pady=(28, 0))
        tk.Label(
            card,
            text="Informe funcional e senha do Painel de Serviços (login da rotina BPM).",
            bg="#1a1a1a", fg="#8c8c8c", font=("Segoe UI", 9),
            wraplength=400, justify="left",
        ).pack(anchor="w", padx=24, pady=(4, 0))
 
        form = tk.Frame(card, bg="#1a1a1a")
        form.pack(fill="x", padx=24, pady=(18, 8))
 
        row_f = tk.Frame(form, bg="#1a1a1a")
        row_f.pack(fill="x", pady=(0, 14))
        tk.Label(row_f, text="Funcional", bg="#1a1a1a", fg="#8c8c8c", font=("Segoe UI", 9)).pack(anchor="w")
        ent_f = tk.Entry(
            row_f, textvariable=self._bpm_func_var,
            bg="#121212", fg="#f2f2f2", insertbackground="#f2f2f2",
            relief="flat", highlightthickness=1, highlightbackground="#2e2e2e", highlightcolor="#2fb875",
            font=("Segoe UI", 10),
        )
        ent_f.pack(fill="x", pady=(6, 0))
        tk.Label(
            row_f, text="Apenas números (sem letras).",
            bg="#1a1a1a", fg="#5a5a5a", font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(4, 0))
 
        row_s = tk.Frame(form, bg="#1a1a1a")
        row_s.pack(fill="x", pady=(0, 8))
        tk.Label(row_s, text="Senha", bg="#1a1a1a", fg="#8c8c8c", font=("Segoe UI", 9)).pack(anchor="w")
        ent_s = tk.Entry(
            row_s, textvariable=self._bpm_senha_var, show="•",
            bg="#121212", fg="#f2f2f2", insertbackground="#f2f2f2",
            relief="flat", highlightthickness=1, highlightbackground="#2e2e2e", highlightcolor="#2fb875",
            font=("Segoe UI", 10),
        )
        ent_s.pack(fill="x", pady=(6, 0))
        tk.Label(
            row_s, text="Apenas números, até 6 dígitos.",
            bg="#1a1a1a", fg="#5a5a5a", font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(4, 0))
 
        _SENHA_MAX = 6
 
        def _key_funcional(e):
            if e.keysym in ("BackSpace", "Delete", "Tab", "Return", "Left", "Right", "Home", "End"):
                return
            if e.state & 0x4:
                return
            if e.char and e.char.isdigit():
                return
            if e.char:
                return "break"
 
        def _key_senha(e):
            if e.keysym in ("BackSpace", "Delete", "Tab", "Return", "Left", "Right", "Home", "End"):
                return
            if e.state & 0x4:
                return
            if not e.char:
                return
            if not e.char.isdigit():
                return "break"
            try:
                has_sel = bool(e.widget.selection_present())
            except Exception:
                has_sel = False
            if len(e.widget.get()) >= _SENHA_MAX and not has_sel:
                return "break"
 
        def _paste_funcional(_e=None):
            try:
                clip = self.clipboard_get()
            except Exception:
                return "break"
            d = self._bpm_only_digits(clip)
            self._bpm_func_var.set(d)
            self._refresh_cred_continue_state()
            return "break"
 
        def _paste_senha(_e=None):
            try:
                clip = self.clipboard_get()
            except Exception:
                return "break"
            d = self._bpm_only_digits(clip)[:_SENHA_MAX]
            self._bpm_senha_var.set(d)
            self._refresh_cred_continue_state()
            return "break"
 
        ent_f.bind("<KeyPress>", _key_funcional)
        ent_f.bind("<<Paste>>", _paste_funcional)
        ent_f.bind("<Control-v>", _paste_funcional)
        ent_s.bind("<KeyPress>", _key_senha)
        ent_s.bind("<<Paste>>", _paste_senha)
        ent_s.bind("<Control-v>", _paste_senha)
 
        def _on_cred_change(*_):
            self._refresh_cred_continue_state()
 
        self._bpm_func_var.trace_add("write", _on_cred_change)
        self._bpm_senha_var.trace_add("write", _on_cred_change)
 
        footer = tk.Frame(card, bg="#1a1a1a")
        footer.pack(fill="x", padx=24, pady=(16, 24))
        tk.Button(
            footer, text="Voltar",
            bg="#1a1a1a", fg="#8c8c8c", bd=0, relief="flat",
            activebackground="#1a1a1a", activeforeground="#f2f2f2",
            cursor="hand2", command=self._render_mode_step, padx=16, pady=8,
        ).pack(side="left")
        self._cred_continue_btn = tk.Button(
            footer, text="Continuar",
            bg="#242424", fg="#6f6f6f", bd=0, relief="flat",
            state="disabled", cursor="no", padx=16, pady=8,
            command=self._credenciais_continue,
        )
        self._cred_continue_btn.pack(side="left", fill="x", expand=True, padx=(10, 0))
        self._refresh_cred_continue_state()
        ent_f.focus_set()
        self._animate_dialog_step()
 
    def _render_clients_step(self):
        self._selected_clients = {}
        self._clear_dialog_body()
        card = self._dialog_body
 
        tk.Label(card, text="Selecione os clientes", bg="#1a1a1a", fg="#f2f2f2", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=24, pady=(28, 0))
        tk.Label(card, text="Escolha os clientes e informe o valor de cada operacao.", bg="#1a1a1a", fg="#8c8c8c", font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(4, 0))
 
        list_wrap = tk.Frame(card, bg="#1a1a1a")
        list_wrap.pack(fill="both", expand=True, padx=24, pady=(16, 8))
        canvas = tk.Canvas(list_wrap, bg="#1a1a1a", highlightthickness=0, bd=0, height=340)
        scrollbar = ttk.Scrollbar(list_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas, bg="#1a1a1a")
        window = canvas.create_window((0, 0), window=body, anchor="nw")
 
        def on_body_cfg(_e): canvas.configure(scrollregion=canvas.bbox("all"))
        def on_canvas_cfg(e): canvas.itemconfigure(window, width=e.width)
        body.bind("<Configure>", on_body_cfg)
        canvas.bind("<Configure>", on_canvas_cfg)
 
        def add_client_row(client_name):
            row = tk.Frame(body, bg="#1f1f1f", highlightthickness=1, highlightbackground="#343434", padx=14, pady=10)
            row.pack(fill="x", pady=3)
            row.columnconfigure(1, weight=1)
            row.selected = False
            row.client_name = client_name
            row.value_var = tk.StringVar()
            row.value_digits = ""
            cb = tk.Label(row, text="", bg="#1f1f1f", fg="#ffffff", width=2, relief="flat", highlightthickness=1, highlightbackground="#343434")
            cb.grid(row=0, column=0, sticky="w")
            name = tk.Label(row, text=client_name, bg="#1f1f1f", fg="#9a9a9a", font=("Segoe UI", 9, "normal"))
            name.grid(row=0, column=1, sticky="w", padx=(10, 8))
            entry = tk.Entry(row, textvariable=row.value_var, width=14, justify="right", bg="#121212", fg="#f2f2f2", insertbackground="#f2f2f2", relief="flat", highlightthickness=1, highlightbackground="#2e2e2e", highlightcolor="#2e2e2e", disabledbackground="#121212", disabledforeground="#7a7a7a")
            row.value_var.set("R$ 0,00")
            entry.configure(state="disabled")
            entry.grid(row=0, column=2, sticky="e")
            entry.grid_remove()
 
            def on_focus_in(_): entry.icursor("end")
            def on_focus_out(_):
                if not row.value_digits: row.value_var.set("R$ 0,00")
                self._refresh_confirm_state()
 
            def _set_value_from_digits():
                if not row.value_digits:
                    row.value_var.set("R$ 0,00")
                    return
                d = Decimal(int(row.value_digits)) / Decimal("100")
                row.value_var.set(_format_brl_currency_from_decimal(d))
 
            def toggle(_=None):
                row.selected = not row.selected
                if row.selected:
                    row.configure(highlightbackground="#2fb875")
                    cb.configure(text="✓", bg="#2fb875", highlightbackground="#2fb875")
                    name.configure(fg="#f2f2f2")
                    entry.configure(state="normal")
                    entry.grid()
                    entry.focus_set()
                    row.value_digits = ""
                    row.value_var.set("R$ 0,00")
                    self._selected_clients[row.client_name] = row.value_var
                else:
                    row.configure(highlightbackground="#343434")
                    cb.configure(text="", bg="#1f1f1f", highlightbackground="#343434")
                    name.configure(fg="#9a9a9a")
                    row.value_digits = ""
                    row.value_var.set("R$ 0,00")
                    entry.configure(state="disabled")
                    entry.grid_remove()
                    self._selected_clients.pop(row.client_name, None)
                self._refresh_confirm_state()
 
            def on_hover(_):
                if not row.selected:
                    row.configure(bg="#292929"); cb.configure(bg="#292929"); name.configure(bg="#292929")
            def on_leave(_):
                if not row.selected:
                    row.configure(bg="#1f1f1f"); cb.configure(bg="#1f1f1f"); name.configure(bg="#1f1f1f")
 
            def on_key_press(e):
                if e.keysym == "BackSpace":
                    row.value_digits = row.value_digits[:-1]; _set_value_from_digits(); self._refresh_confirm_state(); return "break"
                if e.char and e.char.isdigit():
                    row.value_digits += e.char; _set_value_from_digits(); self._refresh_confirm_state(); return "break"
                if e.keysym in {"Tab", "Left", "Right", "Home", "End"}: return None
                return "break"
 
            def on_paste(_e):
                try: clip = self.clipboard_get()
                except Exception: return "break"
                d = _parse_brl_to_decimal(clip)
                if d is None: return "break"
                cents = int((d * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                row.value_digits = str(max(cents, 0)); _set_value_from_digits(); self._refresh_confirm_state(); return "break"
 
            for w in (row, cb, name):
                w.bind("<Button-1>", toggle); w.bind("<Enter>", on_hover); w.bind("<Leave>", on_leave)
            entry.bind("<KeyPress>", on_key_press); entry.bind("<<Paste>>", on_paste)
            entry.bind("<Control-v>", on_paste); entry.bind("<FocusIn>", on_focus_in); entry.bind("<FocusOut>", on_focus_out)
 
        for client in self.CLIENTS:
            add_client_row(client)
 
        footer = tk.Frame(card, bg="#1a1a1a")
        footer.pack(fill="x", padx=24, pady=(16, 24))
        tk.Button(footer, text="Voltar", bg="#1a1a1a", fg="#8c8c8c", bd=0, relief="flat", activebackground="#1a1a1a", activeforeground="#f2f2f2", cursor="hand2", command=self._render_credenciais_step, padx=16, pady=8).pack(side="left")
        self._confirm_btn = tk.Button(footer, text="Confirmar (0)", bg="#242424", fg="#6f6f6f", bd=0, relief="flat", state="disabled", cursor="no", padx=16, pady=8, command=self._confirm_clients_selection)
        self._confirm_btn.pack(side="left", fill="x", expand=True, padx=(10, 0))
        self._refresh_confirm_state()
        self._animate_dialog_step()
 
    def _refresh_confirm_state(self):
        if self._confirm_btn is None or not self._confirm_btn.winfo_exists(): return
        selected = len(self._selected_clients)
        all_filled = selected > 0
        for v in self._selected_clients.values():
            raw = (v.get() or "").strip()
            d = _parse_brl_to_decimal(raw)
            if d is None or d == Decimal("0.00"):
                all_filled = False; break
        self._confirm_btn.configure(text=f"Confirmar ({selected})")
        if all_filled:
            self._confirm_btn.configure(state="normal", bg="#2fb875", fg="#ffffff", activebackground="#2aa66a", cursor="hand2")
        else:
            self._confirm_btn.configure(state="disabled", bg="#242424", fg="#6f6f6f", activebackground="#242424", cursor="no")
 
    def _confirm_clients_selection(self):
        selection = []
        for client_name in self.CLIENTS:
            var = self._selected_clients.get(client_name)
            if var is None: continue
            raw = (var.get() or "").strip()
            d = _parse_brl_to_decimal(raw)
            if d is not None and d != Decimal("0.00"):
                selection.append({"cliente": client_name, "valor": _format_brl_currency(raw)})
        self.controller.bpm_run_selection = selection
        self.close_bpm_mode_dialog()
        self.controller.show_frame("BPM")
 
 
class ShareFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self.vars = {k: tk.StringVar() for k in ["razao_social","cnpj","conta","plataforma","regiao","trader_espec","valor_operacao","spread","prazo_min","prazo_max","modalidade"]}
        self.hidden = {"premio": "sem prêmio", "liquidacao": "Débito em CC", "cnpj8": ""}
        self.vars["modalidade"].trace_add("write", lambda *_: self.update_resumo())
        self.vars["regiao"].trace_add("write", lambda *_: self.update_trader_espec_from_regiao())
        self._build_ui()
 
    def _scroll_mousewheel(self, event):
        cvs = self._scroll_canvas
        if not cvs:
            return
        if getattr(event, "delta", 0):
            cvs.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            cvs.yview_scroll(-3, "units")
        elif event.num == 5:
            cvs.yview_scroll(3, "units")

    def _bind_mousewheel_tree(self, widget):
        widget.bind("<MouseWheel>", self._scroll_mousewheel)
        widget.bind("<Button-4>", self._scroll_mousewheel)
        widget.bind("<Button-5>", self._scroll_mousewheel)
        for child in widget.winfo_children():
            self._bind_mousewheel_tree(child)

    def _install_scroll_bindings(self):
        if self._scroll_canvas:
            self._bind_mousewheel_tree(self)

    def _make_scrollable(self, parent):
        self._scroll_canvas = None
        w = ttk.Frame(parent, style="Root.TFrame")
        w.pack(fill="both", expand=True)
        cvs = tk.Canvas(w, bg=self.COLORS["bg"], highlightthickness=0, bd=0)
        self._scroll_canvas = cvs
        sb = ttk.Scrollbar(w, orient="vertical", command=cvs.yview)
        cvs.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cvs.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(cvs, style="Root.TFrame")
        win = cvs.create_window((0, 0), window=inner, anchor="nw")
        def oncfg(_): cvs.configure(scrollregion=cvs.bbox("all"))
        def oncc(e): cvs.itemconfigure(win, width=e.width)
        inner.bind("<Configure>", oncfg)
        cvs.bind("<Configure>", oncc)
        return inner
 
    def _panel(self, parent, **k):
        p = ttk.Frame(parent, style="Dark.TFrame")
        p.pack(**k)
        return p
 
    def _build_ui(self):
        top = self._panel(self, padx=14, pady=12, fill="x")
        ttk.Button(top, text="← Voltar", style="Dark.TButton", command=lambda: self.controller.show_frame("Home")).pack(side="left")
        ttk.Button(top, text="📄 Abrir PDF…", style="Dark.TButton", command=self.on_open_pdf).pack(side="left", padx=(10, 0))
        ttk.Label(top, text="Cadastro Share (extração do PDF + resumo)", style="Sub.TLabel").pack(side="left", padx=10)
        scroll = self._make_scrollable(self)
        main = ttk.Frame(scroll, style="Root.TFrame")
        main.pack(fill="both", expand=True, padx=14, pady=10)
        card = ttk.Frame(main, style="Dark.TFrame")
        card.pack(fill="x")
        ci = ttk.Frame(card, style="Dark.TFrame")
        ci.pack(fill="x", padx=12, pady=12)
        ttk.Label(ci, text="Campos extraídos (edite se necessário)", style="Sub.TLabel").pack(anchor="w", pady=(0, 10))
        grid = ttk.Frame(ci, style="Dark.TFrame")
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
 
        def add(parent, key, label, row, col, width=36, widget="entry", state="normal"):
            b = ttk.Frame(parent, style="Dark.TFrame")
            b.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 10, 0), pady=8)
            ttk.Label(b, text=label, style="Dark.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=(0, 6))
            r = ttk.Frame(b, style="Dark.TFrame")
            r.grid(row=1, column=0, sticky="ew")
            r.columnconfigure(0, weight=1)
            if widget == "combo":
                w = ttk.Combobox(r, textvariable=self.vars[key], values=["Autorizadas CSV", "Autorizadas SISPAG"], state="normal", width=width)
                w.grid(row=0, column=0, sticky="ew", padx=(0, 8))
                ew = w
            else:
                e = ttk.Entry(r, textvariable=self.vars[key], width=width, style="Dark.TEntry")
                if state == "readonly": e.state(["readonly"])
                e.grid(row=0, column=0, sticky="ew", padx=(0, 8))
                ew = e
            ttk.Button(r, text="Copiar", style="Dark.TButton", command=lambda k=key: self.copy_var(k)).grid(row=0, column=1, sticky="e")
            return ew
 
        specs = [
            ("razao_social", "Razão Social:", 52, "entry", "normal"),
            ("cnpj", "CNPJ:", 30, "entry", "normal"),
            ("conta", "Conta Corrente do Cliente:", 30, "entry", "normal"),
            ("valor_operacao", "Valor da Operação:", 30, "entry", "normal"),
            ("plataforma", "Plataforma:", 20, "entry", "normal"),
            ("regiao", "Região da Plataforma:", 20, "entry", "normal"),
            ("trader_espec", "Trader e Espec:", 32, "entry", "readonly"),
            ("spread", "Spread:", 20, "entry", "normal"),
            ("prazo_min", "Prazo Mínimo NF (dias):", 20, "entry", "normal"),
            ("prazo_max", "Prazo Máximo NF (dias):", 20, "entry", "normal"),
            ("modalidade", "Modalidade:", 28, "combo", "normal"),
        ]
        r = 0; c = 0
        for k, l, w, t, st in specs:
            add(grid, k, l, r, c, width=w, widget=t, state=st)
            c += 1
            if c > 1: c = 0; r += 1
 
        act = ttk.Frame(main, style="Dark.TFrame")
        act.pack(fill="x", pady=10)
        ttk.Button(act, text="🔄 Gerar/Atualizar Resumo", style="Dark.TButton", command=self.update_resumo).pack(side="left")
        ttk.Button(act, text="🧽 Limpar", style="Dark.TButton", command=self.clear_all).pack(side="left", padx=8)
        ttk.Separator(main, orient="horizontal", style="Dark.TSeparator").pack(fill="x", pady=(6, 12))
        res = ttk.LabelFrame(main, text="Resumo pronto para copiar", style="Dark.TLabelframe")
        res.pack(fill="both", expand=True)
        tw = ttk.Frame(res, style="Dark.TFrame")
        tw.pack(fill="both", expand=True, padx=10, pady=10)
        tw.columnconfigure(0, weight=1); tw.rowconfigure(0, weight=1)
        self.txt_resumo = tk.Text(tw, height=10, wrap="word", bd=0, relief="flat", background=self.COLORS["muted"], foreground=self.COLORS["text"], insertbackground=self.COLORS["text"], highlightthickness=1, highlightbackground=self.COLORS["border"], highlightcolor=self.COLORS["accent"], padx=10, pady=10)
        self.txt_resumo.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        sc = ttk.Scrollbar(tw, orient="vertical", command=self.txt_resumo.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.txt_resumo.configure(yscrollcommand=sc.set)
        btm = ttk.Frame(main, style="Dark.TFrame")
        btm.pack(fill="x", pady=10)
        ttk.Button(btm, text="📋 Copiar Resumo", style="Dark.TButton", command=self.copy_resumo).pack(side="left")
        ttk.Button(btm, text="💾 Salvar Resumo (.txt)", style="Dark.TButton", command=self.save_resumo).pack(side="left", padx=8)
        self._install_scroll_bindings()

    def on_show(self):
        self._install_scroll_bindings()

    def copy_var(self, k):
        v = self.vars[k].get(); self.clipboard_clear(); self.clipboard_append(v); self._toast(f"Copiado: {k}")
 
    def copy_resumo(self):
        v = self.txt_resumo.get("1.0", "end-1c"); self.clipboard_clear(); self.clipboard_append(v); self._toast("Resumo copiado!")
 
    def save_resumo(self):
        c = self.txt_resumo.get("1.0", "end-1c").strip()
        if not c: messagebox.showinfo("Salvar Resumo", "Nada para salvar."); return
        p = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de Texto", "*.txt")], title="Salvar Resumo")
        if p:
            try: open(p, "w", encoding="utf-8").write(c); self._toast("Resumo salvo.")
            except Exception as e: messagebox.showerror("Erro", f"Não foi possível salvar: {e}")
 
    def update_trader_espec_from_regiao(self):
        reg = (self.vars["regiao"].get() or "").strip()
        tr, sp = trader_espec_from_regiao(reg)
        self.vars["trader_espec"].set(f"{tr} / {sp}".strip(" /") if (tr or sp) else "")
        self.update_resumo()
 
    def on_open_pdf(self):
        p = filedialog.askopenfilename(title="Selecione um PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if not p: return
        try:
            t_layout, t_plain = extract_text_from_pdf(p)
            t = t_layout
            if not (t or "").strip(): messagebox.showwarning("PDF sem texto", "Não foi possível extrair texto do PDF."); return
            lines = t.splitlines()
            tn, tc = normalize_text_variants(t)
            raz = extract_razao_social(t, lines, tn, tc, t_plain=t_plain) or ""
            cn = extract_cnpj(t, tn, tc) or ""
            cnf = format_cnpj(cn)
            cn8 = only_digits(cn)[:8] if len(only_digits(cn)) >= 8 else ""
            conta = extract_conta_corrente(lines, tn, tc) or ""
            plat = extract_plataforma(t, tn, tc) or ""
            reg = extract_regiao_plataforma(t, tn, tc) or ""
            val = extract_valor_operacao(t, tn, tc) or ""
            sp_raw = extract_spread(t, tn, tc) or ""
            sp_fmt = normalize_percent_br(sp_raw) if sp_raw else ""
            pmin = extract_prazo_minimo_nf(t, tn, tc) or ""
            pmax = extract_prazo_maximo_nf(t, tn, tc) or ""
            mod = extract_modalidade(t, tn, tc) or ""
            liq = "Débito em CC" if not RE_LIQ_CRED.search(t) else "Crédito em CC"
            prem = "com prêmio" if RE_PREMIO.search(t) else "sem prêmio"
            self.vars["razao_social"].set(raz); self.vars["cnpj"].set(cnf or cn); self.vars["conta"].set(conta)
            self.vars["plataforma"].set(plat); self.vars["regiao"].set(reg); self.vars["valor_operacao"].set(val)
            self.vars["spread"].set(sp_fmt or sp_raw); self.vars["prazo_min"].set(pmin); self.vars["prazo_max"].set(pmax)
            self.vars["modalidade"].set(mod)
            tr, sp = trader_espec_from_regiao(reg)
            self.vars["trader_espec"].set(f"{tr} / {sp}".strip(" /") if (tr or sp) else "")
            self.hidden["premio"] = prem; self.hidden["liquidacao"] = liq; self.hidden["cnpj8"] = cn8
            self.update_resumo()
        except Exception as e:
            messagebox.showerror("Erro ao ler PDF", str(e))
 
    def update_resumo(self):
        spread = (self.vars["spread"].get() or "").strip()
        pmin = (self.vars["prazo_min"].get() or "").strip()
        pmax = (self.vars["prazo_max"].get() or "").strip()
        plat = (self.vars["plataforma"].get() or "").strip()
        reg = (self.vars["regiao"].get() or "").strip()
        mod = normalize_modalidade(self.vars["modalidade"].get() or "")
        prem = self.hidden.get("premio", "sem prêmio")
        liq = self.hidden.get("liquidacao", "Débito em CC")
        te = (self.vars["trader_espec"].get() or "").strip()
        if spread and "a.a" not in spread.lower(): spread = normalize_percent_br(spread)
        troca = infer_troca_from_modalidade(mod)
        prazo = ""
        if pmin and pmax: prazo = f"Prazo {pmin} a {pmax} dias"
        elif pmin: prazo = f"Prazo mínimo {pmin} dias"
        elif pmax: prazo = f"Prazo máximo {pmax} dias"
        pr = ""
        if plat and reg: pr = f"Plataforma {plat} – Região {reg}"
        elif plat: pr = f"Plataforma {plat}"
        elif reg: pr = f"Região {reg}"
        L = []
        if spread: L.append(f"Spread mínimo {spread} – {premio_str(prem)} – Troca de Arquivo {troca}")
        else: L.append(f"Spread mínimo – {premio_str(prem)} – Troca de Arquivo {troca}")
        if prazo: L.append(prazo)
        if liq: L.append(f"Liquidação {liq}")
        if pr: L.append(pr)
        if te: L.append(f"Trader/Espec {te}")
        r = "\n".join(L).strip()
        self.txt_resumo.delete("1.0", "end"); self.txt_resumo.insert("1.0", r)
 
    def clear_all(self):
        for v in self.vars.values(): v.set("")
        self.hidden["premio"] = "sem prêmio"; self.hidden["liquidacao"] = "Débito em CC"; self.hidden["cnpj8"] = ""
        self.txt_resumo.delete("1.0", "end")
 
    def _toast(self, msg):
        t = tk.Toplevel(self); t.overrideredirect(True); t.attributes("-topmost", True); t.configure(bg=self.COLORS["muted"])
        lbl = tk.Label(t, text=msg, bg=self.COLORS["muted"], fg=self.COLORS["text"], padx=12, pady=8); lbl.pack()
        self.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() - lbl.winfo_reqwidth() - 36
        y = self.winfo_rooty() + 24
        t.geometry(f"+{x}+{y}"); t.after(1200, t.destroy)
 
 
class LimitesInvertidoFrame(ttk.Frame):
    """Frame para consulta de limites do Invertido."""
 
    LIMIT_GREEN = "#2ecc71"
    LIMIT_ORANGE = "#f39c12"
    LIMIT_RED = "#e74c3c"
    CARD_W = 212
    CARD_H = 180
    CARD_R = 18
 
    MAPPED_CLIENTS = {
        "Transdourada", "RPB", "Posto Arinos", "Mirian Varzea",
        "Petrocal", "Posto Sapucaia",
    }
    MIRROR_CLIENTS = {
        "Brasnorte":                "Posto Arinos",
        "Mirian Cuiaba":            "Mirian Varzea",
        "PetroMix":                 "Petrocal",
        "PetroVel":                 "Petrocal",
        "Auto Posto M Timbozao":    "Posto Sapucaia",
        "Posto Gasol Timbo III":    "Posto Sapucaia",
        "Posto Timbozao Itaperuna": "Posto Sapucaia",
        "Posto Pioneiro":           "Posto Sapucaia",
    }
 
    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self._cards = []
        self._worker_running = False
        self._cancel_requested = False
        self._browser = None
        self._started = False
        self._restart_pending = False
        self._build_ui()
 
    def _ui(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass
 
    def _build_ui(self):
        page_bg = self.COLORS["bg"]
 
        header_bar = tk.Frame(self, bg=page_bg)
        header_bar.pack(fill="x", padx=20, pady=(18, 0))
        tk.Button(
            header_bar, text="← Voltar", command=self._go_back,
            bg=page_bg, fg="#8c8c8c", bd=0, relief="flat",
            activebackground=page_bg, activeforeground="#f2f2f2",
            cursor="hand2", font=("Segoe UI", 9), padx=0, pady=0,
        ).pack(side="left")
 
        title_bar = tk.Frame(self, bg=page_bg)
        title_bar.pack(fill="x", pady=(20, 0))
        self._title_lbl = tk.Label(
            title_bar, text="Limites Invertido",
            bg=page_bg, fg="#f2f2f2", font=("Segoe UI", 15, "bold"),
        )
        self._title_lbl.pack()
        self._subtitle_lbl = tk.Label(
            title_bar, text="Consulte o LTC e o limite disponível de cada cliente.",
            bg=page_bg, fg="#8c8c8c", font=("Segoe UI", 10, "normal"),
        )
        self._subtitle_lbl.pack(pady=(6, 0))
        self._btn_restart = tk.Button(
            title_bar,
            text="↻",
            command=self._restart_limites,
            bg=page_bg, fg="#6f6f6f", bd=0, relief="flat",
            activebackground=page_bg, activeforeground="#f2f2f2",
            cursor="hand2", font=("Segoe UI", 11), padx=2, pady=2,
            takefocus=0,
        )
        self._btn_restart.pack(pady=(8, 0))
 
        scroll_wrap = tk.Frame(self, bg=page_bg)
        scroll_wrap.pack(fill="both", expand=True, padx=0, pady=(18, 0))
 
        self._canvas = tk.Canvas(scroll_wrap, bg=page_bg, highlightthickness=0, bd=0)
        self._vscroll = ttk.Scrollbar(scroll_wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vscroll.set)
        self._vscroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
 
        self._scroll_inner = tk.Frame(self._canvas, bg=page_bg)
        self._scroll_win = self._canvas.create_window((0, 0), window=self._scroll_inner, anchor="nw")
 
        def _on_inner_cfg(_e):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
 
        def _on_canvas_cfg(e):
            # Mantém o container interno com a largura do viewport do canvas.
            # Assim o frame de cards pode ser centralizado corretamente.
            self._canvas.itemconfigure(self._scroll_win, width=e.width)
 
        self._scroll_inner.bind("<Configure>", _on_inner_cfg)
        self._canvas.bind("<Configure>", _on_canvas_cfg)
 
        def _mw(e):
            if getattr(e, "delta", 0):
                self._canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
            elif e.num == 4:
                self._canvas.yview_scroll(-3, "units")
            elif e.num == 5:
                self._canvas.yview_scroll(3, "units")
 
        self._canvas.bind_all("<MouseWheel>", _mw)
        self._canvas.bind_all("<Button-4>", _mw)
        self._canvas.bind_all("<Button-5>", _mw)
 
        self._grid = tk.Frame(self._scroll_inner, bg=page_bg)
        self._grid.pack(padx=20, pady=(10, 20), anchor="n")
        self._grid.columnconfigure(0, weight=1, uniform="lcards")
        self._grid.columnconfigure(1, weight=1, uniform="lcards")
        self._grid.columnconfigure(2, weight=1, uniform="lcards")
 
        self._footer = tk.Frame(self._scroll_inner, bg=page_bg)
        self._footer.pack(fill="x", padx=20, pady=(0, 20))
 
    def _go_back(self):
        self.controller.show_frame("Home")

    def _restart_limites(self):
        """Limpa os cards atuais e refaz a rotina sob demanda."""
        self._restart_pending = True
        self._cancel_requested = True
        br = getattr(self, "_browser", None)
        if br is not None:
            try: br.close()
            except Exception: pass
        self.after(200, self._start_restart_if_possible)

    def _start_restart_if_possible(self):
        if not self._restart_pending:
            return
        if self._worker_running:
            self.after(220, self._start_restart_if_possible)
            return
        self._restart_pending = False
        self._cards = []
        for w in self._grid.winfo_children():
            w.destroy()
        self._started = False
        self._cancel_requested = False
        self._setup_cards()
        self._started = True
        self._start_worker()
 
    def on_show(self):
        if self._started:
            return
        self._started = True
        self._cancel_requested = False
        self._setup_cards()
        self._start_worker()
 
    @staticmethod
    def _hex_to_rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
 
    @staticmethod
    def _rgb_to_hex(r, g, b):
        return f"#{int(r):02x}{int(g):02x}{int(b):02x}"
 
    def _lerp_hex(self, a, b, t):
        t = max(0.0, min(1.0, t))
        ar, ag, ab = self._hex_to_rgb(a)
        br, bg_, bb = self._hex_to_rgb(b)
        return self._rgb_to_hex(ar+(br-ar)*t, ag+(bg_-ag)*t, ab+(bb-ab)*t)
 
    def _stop_spinner(self, idx):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        aid = c.get("anim_after_id")
        if aid is not None:
            try: self.after_cancel(aid)
            except Exception: pass
        c["anim_after_id"] = None
 
    def _tick_spinner(self, idx):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        if c.get("state") not in ("processing", "waiting_mapped"): return
        c["angle"][0] = (c["angle"][0] + 20) % 360
        ang = c["angle"][0]
        if c.get("icon_arc_id") is not None:
            c["icon_canvas"].itemconfigure(c["icon_arc_id"], start=ang)
        if c.get("status_arc_id") is not None:
            c["status_spin"].itemconfigure(c["status_arc_id"], start=ang)
        c["anim_after_id"] = self.after(70, lambda i=idx: self._tick_spinner(i))
 
    def _apply_shell(self, c, inner_bg, border_color):
        c["inner"].configure(bg=inner_bg)
        c["icon_canvas"].configure(bg=inner_bg)
        c["status_spin"].configure(bg=inner_bg)
        c["status_row"].configure(bg=inner_bg)
        c["card_canvas"].itemconfigure(c["round_rect_id"], fill=inner_bg, outline=border_color, width=2)
 
    def _render_icon_waiting(self, c, bg):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline="#333333", width=1)
        ic.create_text(22, 22, text="–", fill="#555555", font=("Segoe UI", 16, "bold"))
 
    def _render_icon_processing(self, c, bg, color=None):
        color = color or self.LIMIT_GREEN
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline="#333333", width=1)
        c["icon_arc_id"] = ic.create_arc(10, 10, 34, 34, start=0, extent=270, style=tk.ARC, width=2, outline=color)
 
    def _render_icon_ok(self, c, bg):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline=self.LIMIT_GREEN, width=1)
        ic.create_text(22, 22, text="✓", fill=self.LIMIT_GREEN, font=("Segoe UI", 17, "bold"))
 
    def _render_icon_warn(self, c, bg):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline=self.LIMIT_ORANGE, width=1)
        ic.create_text(22, 22, text="!", fill=self.LIMIT_ORANGE, font=("Segoe UI", 17, "bold"))
 
    def _render_icon_error(self, c, bg):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline=self.LIMIT_RED, width=1)
        ic.create_text(22, 22, text="✗", fill=self.LIMIT_RED, font=("Segoe UI", 17, "bold"))
 
    def _render_icon_pending(self, c, bg):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline="#444444", width=1)
        ic.create_text(22, 22, text="⧗", fill="#555555", font=("Segoe UI", 14))
 
    def _render_spinner_status(self, c, bg, color=None):
        color = color or self.LIMIT_GREEN
        sc = c["status_spin"]; sc.delete("all"); sc.configure(bg=bg)
        c["status_arc_id"] = sc.create_arc(2, 2, 20, 20, start=0, extent=270, style=tk.ARC, width=2, outline=color)
 
    def _setup_cards(self):
        for w in self._grid.winfo_children(): w.destroy()
        self._cards = []
 
        all_clients = list(BPM_CLIENT_DATA.keys())
        page_bg = self.COLORS["bg"]
 
        for idx, client_name in enumerate(all_clients):
            is_mapped = client_name in self.MAPPED_CLIENTS
            is_mirror = client_name in self.MIRROR_CLIENTS
            row = idx // 3
            col = idx % 3
            bg_init = "#171717"
            bord_init = "#2e2e2e"
 
            outer = tk.Frame(self._grid, bg=page_bg)
            outer.grid(row=row, column=col, sticky="nsew", padx=6, pady=6)
 
            card_canvas = tk.Canvas(outer, width=self.CARD_W, height=self.CARD_H, highlightthickness=0, bg=page_bg)
            card_canvas.pack()
 
            rr_id = _bpm_round_rect(card_canvas, 0, 0, self.CARD_W, self.CARD_H, self.CARD_R, fill=bg_init, outline=bord_init, width=2)
 
            inner = tk.Frame(card_canvas, bg=bg_init)
            card_canvas.create_window(self.CARD_W // 2, self.CARD_H // 2, window=inner, width=self.CARD_W - 28, height=self.CARD_H - 28)
 
            icon_canvas = tk.Canvas(inner, width=44, height=44, highlightthickness=0, bg=bg_init)
            icon_canvas.pack(pady=(2, 4))
 
            name_lbl = tk.Label(inner, text=client_name, bg=bg_init,
                                fg="#c0c0c0" if (is_mapped or is_mirror) else "#8c8c8c",
                                font=("Segoe UI", 9, "bold"))
            name_lbl.pack()
 
            if is_mirror:
                source = self.MIRROR_CLIENTS[client_name]
                badge_lbl = tk.Label(inner, text=f"via {source}", bg=bg_init, fg="#555555",
                                     font=("Segoe UI", 7, "italic"))
                badge_lbl.pack()
            else:
                badge_lbl = None
 
            info_lbl = tk.Label(inner, text="", bg=bg_init, fg="#8c8c8c",
                                font=("Segoe UI", 8), wraplength=170, justify="center")
            info_lbl.pack(pady=(3, 0))
 
            status_row = tk.Frame(inner, bg=bg_init)
            status_row.pack(pady=(8, 2))
 
            status_spin = tk.Canvas(status_row, width=22, height=22, highlightthickness=0, bg=bg_init)
            if is_mapped:
                status_text = "EM ESPERA"
            elif is_mirror:
                status_text = "AGUARDANDO"
            else:
                status_text = "A MAPEAR"
            status_txt = tk.Label(status_row, bg=bg_init, fg="#8c8c8c",
                                  font=("Segoe UI", 8, "bold"), text=status_text)
            status_txt.pack(side=tk.LEFT)
 
            card_data = {
                "outer": outer,
                "card_canvas": card_canvas,
                "round_rect_id": rr_id,
                "inner": inner,
                "icon_canvas": icon_canvas,
                "badge_lbl": badge_lbl,
                "status_spin": status_spin,
                "status_row": status_row,
                "status_txt": status_txt,
                "name_lbl": name_lbl,
                "info_lbl": info_lbl,
                "angle": [0],
                "state": "init",
                "anim_after_id": None,
                "icon_arc_id": None,
                "status_arc_id": None,
                "is_mapped": is_mapped,
                "is_mirror": is_mirror,
                "client_name": client_name,
            }
            self._cards.append(card_data)
 
            if is_mapped:
                self._render_icon_waiting(card_data, bg_init)
            elif is_mirror:
                self._render_icon_waiting(card_data, bg_init)
            else:
                self._render_icon_pending(card_data, bg_init)
 
        for idx in range(len(self._cards)):
            self.after(60 * idx, lambda i=idx: self._reveal_card(i))
 
    def _reveal_card(self, idx):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        if c["is_mapped"]:
            self._set_card_state(idx, "waiting")
        elif c["is_mirror"]:
            self._set_card_state(idx, "mirror_waiting")
        else:
            self._set_card_state(idx, "pending")
 
    def _set_card_state(self, idx, state, info_text=""):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        self._stop_spinner(idx)
        c["state"] = state
        G = self.LIMIT_GREEN
        O = self.LIMIT_ORANGE
        R = self.LIMIT_RED
 
        if state == "waiting":
            bg = "#171717"; bord = "#2e2e2e"
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#c0c0c0")
            c["info_lbl"].configure(bg=bg, text="")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg="#8c8c8c", text="EM ESPERA")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_waiting(c, bg)
 
        elif state == "processing":
            bg = self.COLORS["muted"]; bord = G
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2")
            c["info_lbl"].configure(bg=bg, text="")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_spin"].pack(side=tk.LEFT, padx=(0, 6))
            c["status_txt"].configure(bg=bg, fg=G, text="CONSULTANDO")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_processing(c, bg)
            self._render_spinner_status(c, bg)
            self._tick_spinner(idx)
 
        elif state == "ok":
            bg = self.COLORS["muted"]; bord = G
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2")
            c["info_lbl"].configure(bg=bg, fg="#a8e6c4", text=info_text)
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg=G, text="LIMITE OK")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_ok(c, bg)
 
        elif state == "warn":
            bg = self.COLORS["muted"]; bord = O
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2")
            c["info_lbl"].configure(bg=bg, fg="#f0c070", text=info_text)
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg=O, text="ATENÇÃO")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_warn(c, bg)
 
        elif state == "error":
            bg = self.COLORS["muted"]; bord = R
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2")
            c["info_lbl"].configure(bg=bg, fg="#f08080", text=info_text)
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg=R, text="ERRO")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_error(c, bg)
 
        elif state == "ltc_expired":
            bg = self.COLORS["muted"]; bord = R
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2")
            c["info_lbl"].configure(bg=bg, fg="#f08080", text=info_text)
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg=R, text="LTC VENCIDO")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_error(c, bg)
 
        elif state == "pending":
            bg = "#151515"; bord = "#2a2a2a"
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#555555")
            c["info_lbl"].configure(bg=bg, text="")
            if c.get("badge_lbl"): c["badge_lbl"].configure(bg=bg, fg="#444444")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg="#444444", text="A MAPEAR")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_pending(c, bg)
 
        elif state == "mirror_waiting":
            bg = "#171717"; bord = "#2a2a2a"
            self._apply_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#888888")
            c["info_lbl"].configure(bg=bg, text="")
            if c.get("badge_lbl"): c["badge_lbl"].configure(bg=bg, fg="#555555")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg="#666666", text="AGUARDANDO")
            c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_waiting(c, bg)
 
        if state in ("ok", "warn", "error", "ltc_expired", "processing"):
            bg = self.COLORS["muted"]
            if c.get("badge_lbl"): c["badge_lbl"].configure(bg=bg, fg="#555555")
 
    def _update_status_text(self, idx, text):
        if idx < 0 or idx >= len(self._cards): return
        self._cards[idx]["status_txt"].configure(text=text)
 
    def _start_worker(self):
        if self._worker_running: return
        self._worker_running = True
        threading.Thread(target=self._worker, daemon=True).start()
 
    def _find_card_idx(self, client_name):
        for i, c in enumerate(self._cards):
            if c["client_name"] == client_name:
                return i
        return -1
 
    def _parse_br_date(self, text: str):
        text = text.strip()
        try:
            return datetime.strptime(text, "%d/%m/%Y").date()
        except ValueError:
            return None
 
    def _parse_br_number(self, text: str):
        text = text.strip().replace(".", "").replace(",", "")
        try:
            return int(text)
        except ValueError:
            return None
 
    def _worker(self):
        if not PLAYWRIGHT_OK or sync_playwright is None:
            self._ui(lambda: messagebox.showerror("Erro", "Playwright não disponível neste ambiente."))
            self._worker_running = False
            return
 
        today = date.today()
 
        with sync_playwright() as p:
            browser = None
            try:
                last_err = None
                for channel in ("chrome", "msedge"):
                    try:
                        browser = p.chromium.launch(channel=channel, headless=False)
                        break
                    except Exception as e:
                        last_err = e
                if browser is None:
                    try:
                        # Fallback para Chromium embutido no Playwright (útil no executável em PC sem Chrome/Edge).
                        browser = p.chromium.launch(headless=False)
                    except Exception as e:
                        last_err = e
                if browser is None:
                    raise RuntimeError(f"Não foi possível iniciar Chrome/Edge. Detalhe: {last_err}")
 
                self._browser = browser
                page = browser.new_page()
                page.set_default_timeout(60_000)
 
                for client_name in list(BPM_CLIENT_DATA.keys()):
                    if self._cancel_requested:
                        break
 
                    if client_name not in self.MAPPED_CLIENTS:
                        continue
 
                    idx = self._find_card_idx(client_name)
                    if idx == -1:
                        continue
 
                    self._ui(lambda i=idx: self._set_card_state(i, "processing"))
 
                    url = LIMITE_CLIENT_URLS.get(client_name)
                    if not url:
                        self._ui(lambda i=idx: self._set_card_state(i, "error", "URL não mapeada"))
                        continue
 
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
 
                        self._ui(lambda i=idx: self._update_status_text(i, "AGUARDANDO PÁGINA"))
                        time.sleep(7)
 
                        if self._cancel_requested:
                            break
 
                        self._ui(lambda i=idx: self._update_status_text(i, "LENDO LTC"))
 
                        ltc_date_str = None
                        RE_DATE = re.compile(r"\d{2}/\d{2}/\d{4}")
 
                        def _read_ltc_date():
                            try:
                                card_venc = page.locator("#vencimentoSubgrupo")
                                card_venc.wait_for(state="visible", timeout=10_000)
                                spans = card_venc.locator("span.u-font-size--14.u-ml--8.u-mt--8.u-block")
                                for si in range(spans.count()):
                                    txt = (spans.nth(si).inner_text() or "").strip()
                                    if RE_DATE.match(txt):
                                        return txt
                            except Exception:
                                pass
 
                            try:
                                all_spans = page.locator("span.u-font-size--14.u-ml--8.u-mt--8.u-block")
                                for si in range(all_spans.count()):
                                    txt = (all_spans.nth(si).inner_text() or "").strip()
                                    if RE_DATE.match(txt):
                                        return txt
                            except Exception:
                                pass
 
                            try:
                                body_text = page.locator("body").inner_text()
                                m = RE_DATE.search(body_text)
                                if m:
                                    return m.group(0)
                            except Exception:
                                pass
 
                            return None
 
                        ltc_date_str = _read_ltc_date()
 
                        ltc_date = self._parse_br_date(ltc_date_str) if ltc_date_str else None
 
                        if ltc_date is None:
                            self._ui(lambda i=idx: self._set_card_state(i, "error", "Não foi possível ler\ndata do LTC"))
                            for mirror_name in LIMITE_SHARED_RESULTS.get(client_name, []):
                                midx = self._find_card_idx(mirror_name)
                                if midx != -1:
                                    self._ui(lambda mi=midx: self._set_card_state(mi, "error", "Não foi possível ler\ndata do LTC"))
                            continue
 
                        if ltc_date <= today:
                            info = f"LTC vencido em {ltc_date_str}"
                            self._ui(lambda i=idx, inf=info: self._set_card_state(i, "ltc_expired", inf))
                            for mirror_name in LIMITE_SHARED_RESULTS.get(client_name, []):
                                midx = self._find_card_idx(mirror_name)
                                if midx != -1:
                                    self._ui(lambda mi=midx, inf=info: self._set_card_state(mi, "ltc_expired", inf))
                            continue
 
                        self._ui(lambda i=idx: self._update_status_text(i, "LENDO LIMITES"))
 
                        limite_disp = None
                        LIMITE_BAIXO_REAIS = 1_000_000
 
                        try:
                            
                            limite_fornecedor = page.evaluate("""
                                (function() {
                                    function norm(s) {
                                        return String(s || '')
                                          .replace(/\\u00a0/g, ' ')
                                          .replace(/\\s+/g, ' ')
                                          .trim()
                                          .toLowerCase();
                                    }
                                    var rows = Array.prototype.slice.call(document.querySelectorAll('table.atual tbody tr'));
                                    for (var i = 0; i < rows.length; i++) {
                                        var r = rows[i];
                                        var tdName = r.querySelector('td.tdNomeFinalidade');
                                        if (!tdName) continue;
                                        if (norm(tdName.textContent) !== 'fornecedor') continue;
                                        var tdDisp = r.querySelector('td[id^=\"valorDisponibilidade_\"]') || r.querySelector('td.tdValorDisp[id^=\"valorDisponibilidade_\"]') || r.querySelector('td.tdValorDisp');
                                        if (!tdDisp) return null;
                                        return norm(tdDisp.textContent);
                                    }
                                    return null;
                                })()
                            """)

                            if limite_fornecedor:
                                limite_disp = self._parse_br_number(limite_fornecedor)
                            else:
                                
                                tds_disp = page.locator("td.tdValorDisp")
                                tds_disp.first.wait_for(state="visible", timeout=10_000)
                                for ti in range(tds_disp.count()):
                                    val = self._parse_br_number(tds_disp.nth(ti).inner_text() or "")
                                    if val is not None and (limite_disp is None or val > limite_disp):
                                        limite_disp = val
                        except Exception:
                            pass
 
                        ltc_fmt = ltc_date_str
                        if limite_disp is not None:
                            disp_fmt = f"R$ {limite_disp:,}".replace(",", ".")
                        else:
                            disp_fmt = "N/D"
 
                        limite_baixo = (
                            limite_disp is not None
                            and limite_disp < LIMITE_BAIXO_REAIS
                        )
                        limite_warn_msg = ""
                        if limite_baixo:
                            limite_warn_msg = (
                                f"Disponibilidade fornecedor abaixo de "
                                f"R$ {LIMITE_BAIXO_REAIS:,}".replace(",", ".")
                            )
 
                        info_lines = [
                            f"LTC ativo · vence {ltc_fmt}",
                            f"Limite Disp. (fornecedor): {disp_fmt}",
                        ]
                        info_text = "\n".join(info_lines)
 
                        if limite_baixo:
                            final_info = info_text + f"\n⚠ {limite_warn_msg}"
                            final_state = "warn"
                            self._ui(lambda i=idx, inf=final_info: self._set_card_state(i, "warn", inf))
                        else:
                            final_state = "ok"
                            self._ui(lambda i=idx, inf=info_text: self._set_card_state(i, "ok", inf))
 
                        mirrors = LIMITE_SHARED_RESULTS.get(client_name, [])
                        for mirror_name in mirrors:
                            midx = self._find_card_idx(mirror_name)
                            if midx == -1:
                                continue
                            mirror_info = info_text + f"\n(via {client_name})"
                            if final_state == "warn":
                                mirror_info = info_text + f"\n⚠ {limite_warn_msg}\n(via {client_name})"
                            self._ui(lambda mi=midx, st=final_state, inf=mirror_info: self._set_card_state(mi, st, inf))
 
                    except BPMUserCancelled:
                        break
                    except Exception as e:
                        if self._cancel_requested:
                            break
                        err_str = str(e)[:80]
                        self._ui(lambda i=idx, err=err_str: self._set_card_state(i, "error", err))
 
            except Exception as e:
                if not self._cancel_requested:
                    err_msg = str(e)
                    self._ui(lambda: messagebox.showerror("Erro na consulta", err_msg))
            finally:
                self._browser = None
                if browser is not None:
                    try: browser.close()
                    except Exception: pass
                self._worker_running = False
 
 
class BPMFrame(ttk.Frame):
    BPM_GREEN = "#2ecc71"
    CARD_W = 212
    CARD_H = 172
    CARD_R = 18
    CANCEL_FG = "#5b5b5b"
    CANCEL_HOVER_FG = "#f25c5c"
 
    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self._started = False
        self._cards = []
        self._selected = []
        self._worker_running = False
        self._cancel_requested = False
        self._browser = None
        self._cancel_anim_after = None
        self._cancel_hover_after = None
        self._cancel_btn = None
        self._build_ui()
 
    def _ui(self, fn):
        try: self.after(0, fn)
        except Exception: pass
 
    def _build_ui(self):
        page_bg = self.COLORS["bg"]
        outer = tk.Frame(self, bg=page_bg)
        outer.pack(fill="both", expand=True)
        container = tk.Frame(outer, bg=page_bg, width=672)
        container.pack(expand=True)
        header = tk.Frame(container, bg=page_bg)
        header.pack(fill="x", pady=(40, 28))
        self._title_lbl = tk.Label(header, text="Operações em andamento", bg=page_bg, fg="#f2f2f2", font=("Segoe UI", 15, "bold"))
        self._title_lbl.pack()
        self._subtitle_lbl = tk.Label(header, text="Acompanhe o progresso de cada cliente.", bg=page_bg, fg="#8c8c8c", font=("Segoe UI", 10, "normal"))
        self._subtitle_lbl.pack(pady=(8, 0))
        self._grid = tk.Frame(container, bg=page_bg)
        self._grid.pack(fill="both", expand=True)
        self._grid.columnconfigure(0, weight=1, uniform="cards")
        self._grid.columnconfigure(1, weight=1, uniform="cards")
        self._grid.columnconfigure(2, weight=1, uniform="cards")
        self._cancel_footer = tk.Frame(container, bg=page_bg)
        self._cancel_footer.pack(fill="x", pady=(32, 0))
 
    @staticmethod
    def _hex_to_rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
 
    @staticmethod
    def _rgb_to_hex(r, g, b):
        return f"#{int(r):02x}{int(g):02x}{int(b):02x}"
 
    def _lerp_hex(self, a: str, b: str, t: float):
        t = max(0.0, min(1.0, t))
        ar, ag, ab = self._hex_to_rgb(a); br, bg, bb = self._hex_to_rgb(b)
        return self._rgb_to_hex(ar+(br-ar)*t, ag+(bg-ag)*t, ab+(bb-ab)*t)
 
    def _clear_cancel_footer(self):
        if self._cancel_anim_after is not None:
            try: self.after_cancel(self._cancel_anim_after)
            except Exception: pass
            self._cancel_anim_after = None
        if self._cancel_hover_after is not None:
            try: self.after_cancel(self._cancel_hover_after)
            except Exception: pass
            self._cancel_hover_after = None
        self._cancel_btn = None
        for w in self._cancel_footer.winfo_children(): w.destroy()
 
    def _setup_cancel_button(self):
        self._clear_cancel_footer()
        page_bg = self.COLORS["bg"]
        n = len(self._selected)
        delay_ms = n * 80 + 200
        row = tk.Frame(self._cancel_footer, bg=page_bg)
        row.pack(fill="x")
        holder = tk.Frame(row, bg=page_bg, height=28)
        holder.pack(fill="x")
        cancel_text = "\u200a".join("CANCELAR OPERAÇÕES")
        btn = tk.Button(holder, text=cancel_text, command=self._on_cancel_operations, font=("Segoe UI", 12), fg=page_bg, activeforeground=self.CANCEL_HOVER_FG, bg=page_bg, activebackground=page_bg, bd=0, relief="flat", highlightthickness=0, padx=0, pady=0, cursor="hand2", borderwidth=0, takefocus=0)
        self._cancel_btn = btn
        btn.place(relx=0.5, y=10, anchor="n")
        target_fg = self.CANCEL_FG
        steps = 10; duration_ms = 300; step_ms = max(1, duration_ms // steps)
 
        def step(i):
            if not btn.winfo_exists(): return
            t = (i + 1) / steps; fg = self._lerp_hex(page_bg, target_fg, t); y = 10 * (1.0 - t)
            btn.configure(fg=fg); btn.place_configure(y=y)
            if i + 1 < steps: self._cancel_anim_after = self.after(step_ms, lambda: step(i + 1))
            else: self._cancel_anim_after = None; btn.place_configure(y=0)
 
        def start_anim():
            if not btn.winfo_exists(): return
            self._cancel_anim_after = None; step(0)
 
        self._cancel_anim_after = self.after(delay_ms, start_anim)
        trans_steps = 4; trans_ms = 200; step_h = max(1, trans_ms // trans_steps)
        normal = self.CANCEL_FG; hover = self.CANCEL_HOVER_FG
 
        def cancel_hover_anim():
            if self._cancel_hover_after is not None:
                try: self.after_cancel(self._cancel_hover_after)
                except Exception: pass
                self._cancel_hover_after = None
 
        def animate_color(to_hex, from_hex, i):
            if not btn.winfo_exists(): return
            t = (i + 1) / trans_steps; btn.configure(fg=self._lerp_hex(from_hex, to_hex, t))
            if i + 1 < trans_steps: self._cancel_hover_after = self.after(step_h, lambda j=i+1: animate_color(to_hex, from_hex, j))
            else: self._cancel_hover_after = None
 
        def on_enter(_): cancel_hover_anim(); cur = btn.cget("fg") if btn.winfo_exists() else normal; animate_color(hover, cur, 0)
        def on_leave(_): cancel_hover_anim(); cur = btn.cget("fg") if btn.winfo_exists() else normal; animate_color(normal, cur, 0)
        btn.bind("<Enter>", on_enter); btn.bind("<Leave>", on_leave)
 
    def _reset_bpm_view(self):
        for i in range(len(self._cards)): self._stop_card_spinner(i)
        self._cards = []
        for w in self._grid.winfo_children(): w.destroy()
        self._clear_cancel_footer()
 
    def _on_cancel_operations(self):
        self._cancel_requested = True
        self.controller.bpm_run_selection = []
        br = getattr(self, "_browser", None)
        if br is not None:
            try: br.close()
            except Exception: pass
        self._reset_bpm_view()
        self._started = False
        self.controller.show_frame("Home")
        home = self.controller.frames.get("Home")
        if home is not None and hasattr(home, "open_bpm_mode_dialog"):
            self.after(0, home.open_bpm_mode_dialog)
 
    def on_show(self):
        selection = getattr(self.controller, "bpm_run_selection", None) or []
        if not selection: return
        self._cancel_requested = False
        if self._started: return
        self._started = True
        self._selected = selection
        self._setup_cards()
        self._start_worker()
 
    def _stop_card_spinner(self, idx: int):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        aid = c.get("anim_after_id")
        if aid is not None:
            try: self.after_cancel(aid)
            except Exception: pass
        c["anim_after_id"] = None; c["icon_arc_id"] = None; c["status_arc_id"] = None
 
    def _tick_card_spinner(self, idx: int):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]
        if c.get("state") != "processing": return
        c["angle"][0] = (c["angle"][0] + 20) % 360; ang = c["angle"][0]
        if c.get("icon_arc_id") is not None: c["icon_canvas"].itemconfigure(c["icon_arc_id"], start=ang)
        if c.get("status_arc_id") is not None: c["status_spin"].itemconfigure(c["status_arc_id"], start=ang)
        c["anim_after_id"] = self.after(70, lambda i=idx: self._tick_card_spinner(i))
 
    def _apply_card_shell(self, c, inner_bg: str, border_color: str):
        c["inner"].configure(bg=inner_bg); c["icon_canvas"].configure(bg=inner_bg)
        c["status_spin"].configure(bg=inner_bg); c["status_row"].configure(bg=inner_bg)
        c["card_canvas"].itemconfigure(c["round_rect_id"], fill=inner_bg, outline=border_color, width=2)
 
    def _render_icon_waiting(self, c, bg: str):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline="#333333", width=1)
 
    def _render_icon_processing(self, c, bg: str):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline="#333333", width=1)
        c["icon_arc_id"] = ic.create_arc(10, 10, 34, 34, start=0, extent=270, style=tk.ARC, width=2, outline=self.BPM_GREEN)
 
    def _render_icon_done(self, c, bg: str):
        ic = c["icon_canvas"]; ic.delete("all"); ic.configure(bg=bg)
        _bpm_round_rect(ic, 1, 1, 43, 43, 8, fill="#252525", outline=self.BPM_GREEN, width=1)
        ic.create_text(22, 22, text="✓", fill=self.BPM_GREEN, font=("Segoe UI", 17, "bold"))
 
    def _render_status_spinner(self, c, bg: str):
        sc = c["status_spin"]; sc.delete("all"); sc.configure(bg=bg)
        c["status_arc_id"] = sc.create_arc(2, 2, 20, 20, start=0, extent=270, style=tk.ARC, width=2, outline=self.BPM_GREEN)
 
    def _setup_cards(self):
        for c in self._grid.winfo_children(): c.destroy()
        self._cards = []
        border_waiting = "#2e2e2e"; bg_waiting = "#171717"
        for idx, item in enumerate(self._selected):
            client = item.get("cliente") or ""
            raw_val = (item.get("valor") or "").strip()
            display_val = raw_val
            if display_val.lower().startswith("r$"): display_val = display_val[2:].strip()
            row = idx // 3; col = idx % 3
            page_bg = self.COLORS["bg"]
            outer = tk.Frame(self._grid, bg=page_bg)
            outer.grid(row=row, column=col, sticky="nsew", padx=6, pady=6)
            outer.grid_remove()
            card_canvas = tk.Canvas(outer, width=self.CARD_W, height=self.CARD_H, highlightthickness=0, bg=page_bg)
            card_canvas.pack()
            round_rect_id = _bpm_round_rect(card_canvas, 0, 0, self.CARD_W, self.CARD_H, self.CARD_R, fill=bg_waiting, outline=border_waiting, width=2)
            inner = tk.Frame(card_canvas, bg=bg_waiting)
            card_canvas.create_window(self.CARD_W // 2, self.CARD_H // 2, window=inner, width=self.CARD_W - 28, height=self.CARD_H - 28)
            icon_canvas = tk.Canvas(inner, width=44, height=44, highlightthickness=0, bg=bg_waiting)
            icon_canvas.pack(pady=(2, 4))
            name_lbl = tk.Label(inner, text=client, bg=bg_waiting, fg="#f2f2f2", font=("Segoe UI", 10, "bold"))
            name_lbl.pack()
            val_lbl = tk.Label(inner, text=f"R$ {display_val}", bg=bg_waiting, fg="#8c8c8c", font=("Segoe UI", 9))
            val_lbl.pack(pady=(3, 0))
            status_row = tk.Frame(inner, bg=bg_waiting)
            status_row.pack(pady=(10, 2))
            status_spin = tk.Canvas(status_row, width=22, height=22, highlightthickness=0, bg=bg_waiting)
            status_txt = tk.Label(status_row, text="EM ESPERA", bg=bg_waiting, fg="#8c8c8c", font=("Segoe UI", 8, "bold"))
            status_txt.pack(side=tk.LEFT)
            self._cards.append({
                "outer": outer, "card": outer, "card_canvas": card_canvas, "round_rect_id": round_rect_id,
                "inner": inner, "icon_canvas": icon_canvas, "status_spin": status_spin, "status_row": status_row,
                "status_txt": status_txt, "name_lbl": name_lbl, "val_lbl": val_lbl,
                "angle": [0], "state": "init", "anim_after_id": None, "icon_arc_id": None, "status_arc_id": None,
            })
            self._render_icon_waiting(self._cards[-1], bg_waiting)
 
        for idx in range(1, len(self._cards)):
            delay = 80 * idx
            self.after(delay, lambda i=idx: self._set_card_state(i, "waiting", show=True))
        if self._cards:
            self.after(0, lambda: self._set_card_state(0, "processing", show=True))
        self._setup_cancel_button()
 
    def _set_card_state(self, idx, state, show=False):
        if idx < 0 or idx >= len(self._cards): return
        c = self._cards[idx]; self._stop_card_spinner(idx); c["state"] = state; g = self.BPM_GREEN
        if state == "processing":
            bg = self.COLORS["muted"]
            self._apply_card_shell(c, bg, g)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2"); c["val_lbl"].configure(bg=bg, fg="#8c8c8c")
            c["status_txt"].configure(bg=bg, fg=g, text="ABRINDO BPM"); c["status_txt"].pack_forget()
            c["status_spin"].pack(side=tk.LEFT, padx=(0, 8)); c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_processing(c, bg); self._render_status_spinner(c, bg); self._tick_card_spinner(idx)
            if show: c["outer"].grid()
        elif state == "waiting":
            bg = "#171717"; bord = "#2e2e2e"
            self._apply_card_shell(c, bg, bord)
            c["name_lbl"].configure(bg=bg, fg="#8c8c8c"); c["val_lbl"].configure(bg=bg, fg="#8c8c8c")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg="#8c8c8c", text="EM ESPERA"); c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_waiting(c, bg)
            if show: c["outer"].grid()
        elif state == "done":
            bg = self.COLORS["muted"]
            self._apply_card_shell(c, bg, g)
            c["name_lbl"].configure(bg=bg, fg="#f2f2f2"); c["val_lbl"].configure(bg=bg, fg="#8c8c8c")
            c["status_spin"].pack_forget(); c["status_txt"].pack_forget()
            c["status_txt"].configure(bg=bg, fg=g, text="CONCLUÍDO"); c["status_txt"].pack(side=tk.LEFT)
            self._render_icon_done(c, bg)
            if show: c["outer"].grid()
 
    def _parse_amount_for_web(self, raw: str) -> str:
        return _format_brl_plain_for_web(raw)
 
    def _start_worker(self):
        if self._worker_running: return
        self._worker_running = True
        threading.Thread(target=self._worker, daemon=True).start()
 
    def _worker(self):
        if not PLAYWRIGHT_OK or sync_playwright is None:
            self._ui(lambda: messagebox.showerror("Erro", "Playwright não disponível neste ambiente."))
            self._worker_running = False
            return
 
        funcional_login = (getattr(self.controller, "bpm_funcional", None) or "").strip()
        senha_login = (getattr(self.controller, "bpm_password", None) or "").strip()
        if not funcional_login or not senha_login:
            self._ui(lambda: messagebox.showerror("BPM", "Funcional ou senha não informados. Refaça o fluxo pelo botão BPM."))
            self._worker_running = False
            return
 
        PAINEL_MIDDLE_URL = "https://painelservicos.cloud.ihf/AplicAutSolicitacoesMiddle"
        PACE_MUL = 0.85
        STEP_MS = 15_000
 
        def pick_browser(p):
            last_err = None
            for channel in ("chrome", "msedge"):
                try: return p.chromium.launch(channel=channel, headless=False)
                except Exception as e: last_err = e
            try:
                return p.chromium.launch(headless=False)
            except Exception as e:
                last_err = e
            raise RuntimeError(f"Não foi possível iniciar Chrome/Edge/Chromium embutido. Detalhe: {last_err}")
 
        def wait_until_enabled(locator, timeout_ms=120_000):
            end = time.time() + timeout_ms / 1000.0
            locator.wait_for(state="visible", timeout=timeout_ms)
            while time.time() < end:
                if self._cancel_requested: raise BPMUserCancelled()
                try:
                    if not locator.is_disabled(): return
                except Exception: pass
                time.sleep(0.2)
            raise TimeoutError("Timeout aguardando campo habilitar.")

        def pace(min_s=1.15, max_s=1.75):
            time.sleep(random.uniform(min_s * PACE_MUL, max_s * PACE_MUL))

        def _to_cent_value(raw_text: str):
            s = (raw_text or "").strip()
            s = s.replace("\xa0", " ")
            s = re.sub(r"[^\d,.\-]", "", s)
            if not s:
                return None
            last_dot = s.rfind(".")
            last_comma = s.rfind(",")
            sep_idx = max(last_dot, last_comma)
            try:
                if sep_idx == -1:
                    d = Decimal(re.sub(r"[^\d\-]", "", s))
                else:
                    int_part = re.sub(r"[^\d\-]", "", s[:sep_idx])
                    dec_part = re.sub(r"[^\d]", "", s[sep_idx + 1 :])
                    if not dec_part:
                        return None
                    d = Decimal((int_part if int_part not in {"", "-"} else "0") + "." + dec_part)
            except Exception:
                return None
            return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        def fill_currency_masked(ctx, locator, amount_value: str, timeout_ms=120_000):
            """
            Campo usa mascaramoeda (Angular). Se usar fill() direto, pode concatenar em zeros existentes.
            Faz limpeza real + digitação humana + validação final do valor renderizado.
            """
            wait_until_enabled(locator, timeout_ms=timeout_ms)
            locator.scroll_into_view_if_needed()
            locator.click()
            pace(0.18, 0.35)
            locator.press("Control+A")
            locator.press("Backspace")
            pace(0.12, 0.25)
            locator.press("Delete")
            pace(0.15, 0.30)

            as_digits = re.sub(r"\D", "", amount_value or "")
            if not as_digits:
                raise RuntimeError(f"Valor inválido para digitação no campo monetário: {amount_value}")

            for ch in as_digits:
                locator.type(
                    ch,
                    delay=random.randint(int(85 * PACE_MUL), int(150 * PACE_MUL)),
                )

            locator.press("Tab")
            pace(0.45, 0.85)

            expected = _to_cent_value(amount_value)
            got = _to_cent_value(locator.input_value())
            if expected is None or got is None or got != expected:
                raw_now = locator.input_value()
                raise RuntimeError(
                    f"Valor mascarado divergente. Esperado {amount_value}, obtido '{raw_now}'."
                )

        def wait_until_continuar_filtro_ready(ctx, timeout_ms=120_000):
            """#continuar usa ng-show='exibirAplicAutFiltro'; só fica realmente exibido após validação (blur/change no funcional)."""
            deadline = time.time() + timeout_ms / 1000.0
            while time.time() < deadline:
                if self._cancel_requested:
                    raise BPMUserCancelled()
                ok = ctx.evaluate("""
                    (function() {
                        var el = document.querySelector('#continuar');
                        if (!el) return false;
                        if (el.disabled === true) return false;
                        if (String(el.getAttribute('aria-disabled')).toLowerCase() === 'true') return false;
                        var st = window.getComputedStyle(el);
                        if (st.display === 'none' || st.visibility === 'hidden') return false;
                        if (parseFloat(st.opacity || '1') === 0) return false;
                        var r = el.getBoundingClientRect();
                        if (r.width < 2 || r.height < 2) return false;
                        return true;
                    })()
                """)
                if ok:
                    return
                time.sleep(0.25)
            raise TimeoutError(
                "Timeout: #continuar não ficou clicável (ng-show / validação do funcional). "
                "Confirme blur/validação no campo #funcionalac."
            )

        def click_continuar_aplic_aut_filtro(ctx, timeout_ms=120_000):
            """Clica em input#continuar (ng-click ContinuarAplicAutFiltro); tenta Playwright, DOM e escopo Angular."""
            loc = ctx.locator('input#continuar[type="submit"]').first

            def _center_mouse():
                bb = loc.bounding_box()
                if not bb:
                    raise RuntimeError("bounding_box vazio para #continuar")
                page.mouse.click(bb["x"] + bb["width"] / 2, bb["y"] + bb["height"] / 2)

            strategies = (
                lambda: (loc.scroll_into_view_if_needed(timeout=10_000), loc.click(timeout=12_000)),
                lambda: loc.click(force=True, timeout=12_000),
                lambda: ctx.evaluate("""
                    (function() {
                        var el = document.querySelector('input#continuar[type="submit"]') || document.querySelector('#continuar');
                        if (!el) throw new Error('continuar nao encontrado');
                        el.click();
                    })()
                """),
                lambda: ctx.evaluate("""
                    (function() {
                        var el = document.querySelector('#continuar');
                        if (!el) throw new Error('continuar');
                        var ang = window.angular;
                        if (!ang) { el.click(); return; }
                        var inj = ang.element(el);
                        var scope = inj.scope() || inj.isolateScope();
                        while (scope) {
                            if (typeof scope.ContinuarAplicAutFiltro === 'function') {
                                scope.$apply(function() { scope.ContinuarAplicAutFiltro(); });
                                return;
                            }
                            scope = scope.$parent;
                        }
                        el.click();
                    })()
                """),
                _center_mouse,
            )

            end = time.time() + timeout_ms / 1000.0
            last_err = None
            while time.time() < end:
                if self._cancel_requested:
                    raise BPMUserCancelled()
                for action in strategies:
                    try:
                        action()
                        return
                    except Exception as e:
                        last_err = e
                time.sleep(0.35 * PACE_MUL)
            raise RuntimeError(f"Falha ao clicar em #continuar (AplicAut filtro). Último erro: {last_err}")

        def resolve_bpm_frame(timeout_ms=120_000):
            deadline = time.time() + timeout_ms / 1000.0
            while time.time() < deadline:
                if self._cancel_requested:
                    raise BPMUserCancelled()
                for frame in list(page.frames):
                    try:
                        if frame.is_detached():
                            continue
                    except Exception:
                        continue
                    try:
                        if frame.locator("#nova-solicitacao").count() > 0:
                            return frame
                    except Exception:
                        pass
                time.sleep(0.35 * PACE_MUL)
            raise TimeoutError("Timeout: #nova-solicitacao não encontrado em nenhum frame (verifique iframes).")
 
        def wait_for_nova_solicitacao(ctx, timeout_ms=120_000):
            deadline = time.time() + timeout_ms / 1000.0
            while time.time() < deadline:
                if self._cancel_requested:
                    raise BPMUserCancelled()
                ready = ctx.evaluate("""
                    (function() {
                        var el = document.querySelector('#nova-solicitacao');
                        if (!el) return false;
                        if (el.disabled === true) return false;
                        if (String(el.getAttribute('aria-disabled')).toLowerCase() === 'true') return false;
                        var st = window.getComputedStyle(el);
                        if (st.display === 'none' || st.visibility === 'hidden') return false;
                        return true;
                    })()
                """)
                if ready:
                    return
                time.sleep(0.4)
            raise TimeoutError("Timeout aguardando #nova-solicitacao ficar pronto.")
 
        def click_nova_solicitacao(ctx, timeout_ms=120_000):
            wait_for_nova_solicitacao(ctx, timeout_ms)

            def click_worked():
                try: ctx.locator("#tipooper").wait_for(state="visible", timeout=5_000); return True
                except Exception: return False

            def _el_js():
                return """(function(){
                    var el = document.querySelector('#nova-solicitacao');
                    if (!el) el = document.evaluate(
                        '//*[@id="nova-solicitacao"]',
                        document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
                    ).singleNodeValue;
                    return el;
                })()"""

            def click_center_css():
                loc = ctx.locator("#nova-solicitacao").first
                loc.scroll_into_view_if_needed()
                box = loc.bounding_box()
                if not box:
                    raise RuntimeError("bounding_box vazio para #nova-solicitacao")
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)

            def click_center_xpath():
                loc = ctx.locator('xpath=//*[@id="nova-solicitacao"]').first
                loc.scroll_into_view_if_needed()
                box = loc.bounding_box()
                if not box:
                    raise RuntimeError("bounding_box vazio para xpath nova-solicitacao")
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)

            strategies = [
                lambda: (
                    ctx.locator("#nova-solicitacao").first.scroll_into_view_if_needed(),
                    time.sleep(0.15),
                    ctx.locator("#nova-solicitacao").first.click(timeout=8_000),
                ),
                lambda: (
                    ctx.locator('[id="nova-solicitacao"]').first.scroll_into_view_if_needed(),
                    time.sleep(0.15),
                    ctx.locator('[id="nova-solicitacao"]').first.click(timeout=8_000),
                ),
                lambda: (
                    ctx.locator('xpath=//*[@id="nova-solicitacao"]').first.scroll_into_view_if_needed(),
                    time.sleep(0.15),
                    ctx.locator('xpath=//*[@id="nova-solicitacao"]').first.click(timeout=8_000),
                ),
                lambda: (ctx.evaluate("(function(){ var el = " + _el_js() + "; if(el) el.scrollIntoView({block:'center',inline:'center'}); })()"),
                         time.sleep(0.2),
                         ctx.evaluate("(function(){ var el = " + _el_js() + "; if(el) el.click(); })()")),
                lambda: ctx.evaluate("(function(){ var el = " + _el_js() + "; if(!el) throw new Error('nova-solicitacao nao encontrado'); el.click(); })()"),
                lambda: ctx.evaluate("""(function(){
                    var el = """ + _el_js() + """;
                    if (!el) throw new Error('nova-solicitacao nao encontrado');
                    el.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true,view:window}));
                    el.dispatchEvent(new MouseEvent('mouseup',{bubbles:true,cancelable:true,view:window}));
                    el.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}));
                })()"""),
                lambda: ctx.evaluate("""(function(){
                    var el = """ + _el_js() + """;
                    if (!el) throw new Error('nova-solicitacao nao encontrado');
                    var scope = angular.element(el).scope();
                    if (scope && typeof scope.NovaSolicitacao === 'function')
                        scope.$apply(function(){ scope.NovaSolicitacao(); });
                    else throw new Error('Angular NovaSolicitacao indisponivel');
                })()"""),
                lambda: click_center_css(),
                lambda: click_center_xpath(),
                lambda: (
                    ctx.locator("#nova-solicitacao").first.scroll_into_view_if_needed(),
                    ctx.locator("#nova-solicitacao").first.click(force=True, timeout=8_000),
                ),
                lambda: ctx.locator('xpath=//*[@id="nova-solicitacao"]').first.click(force=True, timeout=8_000),
                lambda: (
                    ctx.locator('xpath=//*[@id="nova-solicitacao"]').first.scroll_into_view_if_needed(),
                    ctx.locator('xpath=//*[@id="nova-solicitacao"]').first.focus(),
                    page.keyboard.press("Enter"),
                ),
            ]
 
            for i, strategy in enumerate(strategies, start=1):
                if self._cancel_requested:
                    raise BPMUserCancelled()
                try:
                    strategy()
                except Exception as exc:
                    print(f"[nova-solicitacao] estratégia {i} erro: {exc}")
                    time.sleep(0.5)
                    continue
 
                for _ in range(12):
                    if click_worked():
                        print(f"[nova-solicitacao] estratégia {i} funcionou.")
                        return
                    time.sleep(0.5)
 
            raise RuntimeError("Falha ao clicar em nova-solicitacao após todas as estratégias.")
 
        def ensure_painel():
            if self._cancel_requested: raise BPMUserCancelled()
            page.goto(PAINEL_MIDDLE_URL, wait_until="domcontentloaded")
            if "/Home" in page.url:
                page.goto(PAINEL_MIDDLE_URL, wait_until="domcontentloaded")
            if page.locator("#username").count() > 0 or "login.itau/oauth" in page.url:
                page.locator("#username").fill(funcional_login)
                page.locator("#password").fill(senha_login)
                page.locator("#btLogin").click()
                page.wait_for_load_state("domcontentloaded")
                page.goto(PAINEL_MIDDLE_URL, wait_until="domcontentloaded")
            ctx = resolve_bpm_frame(timeout_ms=120_000)
            wait_for_nova_solicitacao(ctx, timeout_ms=120_000)
            return ctx
 
        with sync_playwright() as p:
            browser = pick_browser(p)
            self._browser = browser
           
            page = browser.new_page()
            page.set_default_timeout(120_000)
            try:
                for idx, item in enumerate(self._selected):
                    if self._cancel_requested: raise BPMUserCancelled()
                    self._ui(lambda i=idx: self._set_card_state(i, "processing", show=True))
                    client_name = item.get("cliente") or ""; raw_amount = item.get("valor") or ""
                    info = BPM_CLIENT_DATA.get(client_name)
                    if not info: raise RuntimeError(f"Cliente sem mapeamento interno: {client_name}")
                    amount_web = self._parse_amount_for_web(raw_amount)
                    if not amount_web: raise RuntimeError(f"Valor inválido para o cliente {client_name}: {raw_amount}")
                    tentativa = 0
                    while True:
                        if self._cancel_requested: raise BPMUserCancelled()
                        tentativa += 1
                        try:
                            ctx = ensure_painel()
                            click_nova_solicitacao(ctx, timeout_ms=120_000)
                            pace()
                            ctx.locator("#tipooper").wait_for(state="visible", timeout=STEP_MS)
                            ctx.locator("#tipooper").select_option(value="AT")
                            pace()
                            ctx.locator("#CPNJ").fill(info["CNPJ"])
                            pace()
                            ctx.locator('input[name="agenciafiltro"]').fill(info["AG"])
                            pace()
                            ctx.locator('input[name="contadacfiltro"]').fill(info["CONTA"])
                            pace()
                            self._ui(lambda i=idx: self._cards[i]["status_txt"].configure(text="ALOCANDO INFORMAÇÕES", fg="#2ecc71"))
                            loc_processar_nova = ctx.locator("#processar-nova")
                            wait_until_enabled(loc_processar_nova, timeout_ms=STEP_MS); loc_processar_nova.click()
                            pace()
                            loc_plat = ctx.locator('input[name="plataforma"]')
                            wait_until_enabled(loc_plat, timeout_ms=STEP_MS); loc_plat.fill(info["PLATAFORMA"])
                            pace()
                            loc_fn = ctx.locator("#funcionalac")
                            wait_until_enabled(loc_fn, timeout_ms=STEP_MS)
                            loc_fn.fill(funcional_login)
                            loc_fn.press("Tab")
                            try:
                                loc_fn.evaluate(
                                    "el => { el.dispatchEvent(new Event('input',{bubbles:true}));"
                                    " el.dispatchEvent(new Event('change',{bubbles:true})); el.blur(); }"
                                )
                            except Exception:
                                pass
                            time.sleep(0.35 * PACE_MUL)
                            wait_until_continuar_filtro_ready(ctx, timeout_ms=STEP_MS)
                            click_continuar_aplic_aut_filtro(ctx, timeout_ms=STEP_MS)
                            pace()
                            ctx.locator('select[name="produto"]').wait_for(state="visible", timeout=STEP_MS)
                            ctx.locator('select[name="produto"]').select_option(value="34")
                            pace()
                            ctx.locator('select[name="tpOperacao"]').wait_for(state="visible", timeout=STEP_MS)
                            ctx.locator('select[name="tpOperacao"]').select_option(value="1")
                            pace()
                            self._ui(lambda i=idx: self._cards[i]["status_txt"].configure(text="ALTERANDO TIPO DE OPERAÇÃO", fg="#2ecc71"))
                            loc_buscar = ctx.locator('input[type="submit"][data-ng-click="BuscarListaPN()"]')
                            if loc_buscar.count() == 0: loc_buscar = ctx.locator('input[type="submit"][value="Processar"]')
                            wait_until_enabled(loc_buscar, timeout_ms=STEP_MS); loc_buscar.click()
                            pace()
                            checkbox = ctx.locator('input[type="checkbox"]')
                            checkbox.first.wait_for(state="visible", timeout=STEP_MS)
                            if checkbox.first.is_disabled():
                                raise TimeoutError("Checkbox visível porém desabilitado.")
                            checkbox.first.click()
                            pace()
                            loc_val = ctx.locator('input[name="ValordaOperacao0"]')
                            wait_until_enabled(loc_val, timeout_ms=STEP_MS)
                            self._ui(lambda i=idx: self._cards[i]["status_txt"].configure(text="INFORMANDO VALOR DA OPERAÇÃO", fg="#2ecc71"))
                            fill_currency_masked(ctx, loc_val, amount_web, timeout_ms=STEP_MS)
                            pace()
                            loc_final = ctx.locator('input[type="submit"][ng-click="vm.finalizarSol()"]')
                            if loc_final.count() == 0: loc_final = ctx.locator('input[type="submit"][value="Continuar"]')
                            wait_until_enabled(loc_final, timeout_ms=STEP_MS); loc_final.click()
                            pace()
                            loc_incluir = ctx.locator('input[type="submit"][value="Incluir"]')
                            if loc_incluir.count() == 0: loc_incluir = ctx.locator('input[type="submit"][ng-click="vm.IncluirConta()"]')
                            wait_until_enabled(loc_incluir, timeout_ms=STEP_MS); loc_incluir.click()
                            pace(1.4, 2.1)

                            def click_grey_continuar():
                                loc_g = ctx.locator('input[type="button"].grey[ng-click="vm.continuar()"]')
                                if loc_g.count() == 0: loc_g = ctx.locator('input.grey[ng-click="vm.continuar()"]')
                                if loc_g.count() == 0: return False
                                wait_until_enabled(loc_g.first, timeout_ms=STEP_MS); loc_g.first.click(); return True

                            click_grey_continuar()
                            pace()
                            try: click_grey_continuar()
                            except Exception: pass
                            loc_voltar = ctx.locator('input[type="button"].grey[ng-click="vm.voltarAplicAut()"]')
                            if loc_voltar.count() == 0: loc_voltar = ctx.locator('input.grey[ng-click="vm.voltarAplicAut()"]')
                            if loc_voltar.count() > 0:
                                wait_until_enabled(loc_voltar.first, timeout_ms=max(STEP_MS, 25_000)); loc_voltar.first.click()
                            ctx_after = resolve_bpm_frame(timeout_ms=120_000)
                            wait_for_nova_solicitacao(ctx_after, timeout_ms=120_000)
                            self._ui(lambda i=idx: self._set_card_state(i, "done", show=True))
                            break
                        except BPMUserCancelled:
                            raise
                        except Exception as e:
                            if tentativa >= 20:
                                raise RuntimeError(
                                    f"Falha recorrente no fluxo BPM para {client_name} após {tentativa} tentativas: {e}"
                                ) from e
                            try:
                                page.goto(PAINEL_MIDDLE_URL, wait_until="domcontentloaded")
                            except Exception:
                                pass
                            continue
 
            except BPMUserCancelled: pass
            except Exception as e:
                if self._cancel_requested: pass
                else: err_msg = str(e); self._ui(lambda: messagebox.showerror("Erro na rotina", err_msg))
            finally:
                self._browser = None
                try: browser.close()
                except Exception: pass
                self._worker_running = False
 
 
class App(tk.Tk):
    COLORS = {
        "bg": "#121212",
        "panel": "#1a1a1a",
        "text": "#f2f2f2",
        "muted": "#1f1f1f",
        "border": "#2e2e2e",
        "accent": "#2fb875",
    }
 
    def __init__(self):
        super().__init__()
        self.title("Mesa - Itaú")
        self.geometry("960x700")
        self.configure(bg=self.COLORS["bg"])
        self._setup_styles()
        self.bpm_run_selection = []
        self.bpm_funcional = ""
        self.bpm_password = ""
        container = ttk.Frame(self, style="Root.TFrame")
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        self.frames = {}
        for Cls, name in (
            (HomeFrame, "Home"),
            (ShareFrame, "Share"),
            (BPMFrame, "BPM"),
            (LimitesInvertidoFrame, "LimitesInvertido"),
        ):
            f = Cls(container, self)
            self.frames[name] = f
            f.grid(row=0, column=0, sticky="nsew")
        self.show_frame("Home")
 
    def _setup_styles(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except Exception: pass
        c = self.COLORS
        s.configure("Root.TFrame", background=c["bg"])
        s.configure("Dark.TFrame", background=c["panel"])
        s.configure("Dark.TLabel", background=c["panel"], foreground=c["text"])
        s.configure("Sub.TLabel", background=c["bg"], foreground=c["text"])
        s.configure("Dark.TButton", background=c["panel"], foreground=c["text"])
        s.map("Dark.TButton", background=[("active", c["border"])])
        s.configure("Primary.TButton", background=c["accent"], foreground="#ffffff")
        s.map("Primary.TButton", background=[("active", "#279865")])
        s.configure("Limit.TButton", background="#1e3a5f", foreground="#7ec8e3")
        s.map("Limit.TButton", background=[("active", "#254d7a")])
        s.configure("Dark.TEntry", fieldbackground=c["muted"], foreground=c["text"])
        s.configure("Dark.TSeparator", background=c["border"])
        s.configure("Dark.TLabelframe", background=c["panel"], foreground=c["text"])
        s.configure("Dark.TLabelframe.Label", background=c["panel"], foreground=c["text"])
 
    def show_frame(self, name):
        f = self.frames.get(name)
        if f is not None:
            f.tkraise()
            if hasattr(f, "on_show"):
                try: f.on_show()
                except Exception: pass
 
 
if __name__ == "__main__":
    App().mainloop()
