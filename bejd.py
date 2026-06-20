import os, sys, re, random, struct, tkinter as tk, threading, time, tempfile, shutil, webbrowser, ctypes, json as _json_mod, uuid as _uuid_mod, queue
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from tkinter import ttk, filedialog, messagebox, font as tkfont
from datetime import datetime, date
from PyPDF2 import PdfReader

try:
    import win32com.client as win32
    WIN32_OK = True
except Exception:
    WIN32_OK = False

try:
    import pythoncom
    PYTHONCOM_OK = True
except Exception:
    PYTHONCOM_OK = False

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_OK = True
except Exception:
    sync_playwright = None
    PLAYWRIGHT_OK = False

try:
    import openpyxl
    OPENPYXL_OK = True
except Exception:
    openpyxl = None
    OPENPYXL_OK = False

C = {
    "bg":          "#191919",
    "surface":     "#212121",
    "surface2":    "#2a2a2a",
    "surface3":    "#333333",
    "ink":         "#e6e6e6",
    "ink_muted":   "#999999",
    "ink_faint":   "#4d4d4d",
    "accent":      "#EC7000",
    "accent_dim":  "#3d2a14",
    "accent_soft": "#2a1f12",
    "ok":          "#4ea87a",
    "ok_dim":      "#1a3a2a",
    "warn":        "#d49b45",
    "err":         "#c95f5f",
    "err_dim":     "#3d1515",
    "hair":        "#2a2a2a",
    "log_step":    "#606060",
    "log_ok":      "#4ea87a",
    "log_warn":    "#d49b45",
    "log_err":     "#c95f5f",
}

DOT_COLORS = [
    "#9b9b9b",
    "#a07450",
    "#c87941",
    "#c4a832",
    "#5a9e72",
    "#EC7000",
    "#8b72c9",
    "#c97a9e",
    "#c96060",
]
DOT_LABELS = ["Cinza","Marrom","Laranja","Amarelo","Verde","Azul","Roxo","Rosa","Vermelho"]

ICON_FILENAME = "itaulogo.png"
LOGO_FILENAME = ICON_FILENAME
APP_USER_MODEL_ID = "MesaItau.RiscoSacado"

MESES_PT = (
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
)
DIAS_PT = (
    "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
    "sexta-feira", "sábado", "domingo",
)


def format_data_pt_br(dt: datetime) -> str:
    return f"{DIAS_PT[dt.weekday()]}, {dt.day} de {MESES_PT[dt.month]} de {dt.year}"


def _parse_brl(raw: str):
    s = (raw or "").strip().replace("R$","").replace("r$","").replace(" ","")
    s = re.sub(r"[^\d,.\-]","",s)
    if not s or s in {".",",","-","-.","-.","-.","-."}:
        return None
    ld, lc = s.rfind("."), s.rfind(",")
    si = max(ld, lc)
    try:
        if si == -1:
            d = Decimal(re.sub(r"[^\d\-]","",s))
        else:
            ip = re.sub(r"[^\d\-]","", s[:si])
            dp = re.sub(r"[^\d]","", s[si+1:])
            if not dp: return None
            d = Decimal((ip if ip not in {"","-"} else "0")+"."+dp)
    except (InvalidOperation, ValueError):
        return None
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

def _fmt_brl(d: Decimal) -> str:
    d2 = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sign = "-" if d2 < 0 else ""
    s = f"{abs(d2):.2f}"; i, f = s.split(".")
    return f"R$ {sign}{'{:,}'.format(int(i)).replace(',','.')},{f}"

def _fmt_brl_from_raw(raw: str) -> str:
    d = _parse_brl(raw)
    return _fmt_brl(d) if d is not None else (raw or "").strip()

def _fmt_brl_plain_web(raw: str) -> str:
    d = _parse_brl(raw)
    if d is None: return ""
    s = f"{abs(d):.2f}"; i, f = s.split(".")
    sign = "-" if d < 0 else ""
    return f"{sign}{'{:,}'.format(int(i)).replace(',','.')},{f}"


def _fmt_date_short(val) -> str:
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, date):
        return val.strftime("%d/%m/%Y")
    s = str(val or "").strip()
    return s


def _valor_to_decimal(val):
    if val is None:
        return Decimal("0")
    if isinstance(val, Decimal):
        return val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if isinstance(val, (int, float)):
        return Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    parsed = _parse_brl(str(val))
    return parsed if parsed is not None else Decimal("0")


def _fmt_valor_cell(val) -> str:
    if val is None:
        return "—"
    if isinstance(val, (int, float, Decimal)):
        try:
            return _fmt_brl(Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
        except Exception:
            pass
    return _fmt_brl_from_raw(str(val))


def _normalize_sacado_key(name: str) -> str:
    return RE_SPACES.sub(" ", (name or "").strip()).upper()


def _find_invertido_header(ws) -> tuple:
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i > 30:
            break
        cells = [str(c or "").strip().lower() for c in (row or [])]
        joined = " ".join(cells)
        if "nome sacado" in joined or "nome_sacado" in joined:
            col_map = {}
            for j, cell in enumerate(cells):
                if "doc sacado" in cell or cell == "doc sacado":
                    col_map["doc_sacado"] = j
                elif "nome sacado" in cell or cell == "nome sacado":
                    col_map["nome"] = j
                elif cell in ("número", "numero", "nº", "nf"):
                    col_map["nf"] = j
                elif cell == "valor":
                    col_map["valor"] = j
                elif "inclus" in cell:
                    col_map["inclusao"] = j
                elif "vencimento" in cell:
                    col_map["vencimento"] = j
                elif cell == "prazo":
                    col_map["prazo"] = j
            if "nome" not in col_map:
                col_map["nome"] = 1
            col_map.setdefault("doc_sacado", 0)
            col_map.setdefault("nf", 4)
            col_map.setdefault("valor", 5)
            col_map.setdefault("inclusao", 6)
            col_map.setdefault("vencimento", 7)
            col_map.setdefault("prazo", 8)
            return i, col_map
    return 2, {"doc_sacado": 0, "nome": 1, "nf": 4, "valor": 5, "inclusao": 6, "vencimento": 7, "prazo": 8}


def _parse_invertido_xlsx(path: str) -> list:
    if not OPENPYXL_OK:
        raise RuntimeError("Biblioteca openpyxl não disponível. Instale com: pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        header_row, cols = _find_invertido_header(ws)
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i <= header_row:
                continue
            if not row:
                continue
            nome_idx = cols["nome"]
            if nome_idx >= len(row) or row[nome_idx] is None:
                continue
            nome = str(row[nome_idx] or "").strip()
            if not nome:
                continue

            def _cell(key):
                idx = cols.get(key)
                if idx is None or idx >= len(row):
                    return None
                return row[idx]

            rows.append({
                "doc_sacado":  only_digits(str(_cell("doc_sacado") or "")),
                "nome_sacado": nome,
                "nf":          str(_cell("nf") or "").strip(),
                "valor_raw":   _valor_to_decimal(_cell("valor")),
                "valor":       _fmt_valor_cell(_cell("valor")),
                "data_inclusao": _fmt_date_short(_cell("inclusao")),
                "data_vencimento": _fmt_date_short(_cell("vencimento")),
                "prazo":       str(_cell("prazo") if _cell("prazo") is not None else "").strip(),
            })
        return rows
    finally:
        wb.close()


def _group_invertido_ops(ops: list) -> list:
    groups = {}
    for op in ops:
        key = _normalize_sacado_key(op["nome_sacado"])
        if key not in groups:
            groups[key] = {
                "nome_sacado": op["nome_sacado"].strip(),
                "doc_sacado":  only_digits(op.get("doc_sacado") or ""),
                "notas": [],
                "total": Decimal("0"),
            }
        if not groups[key]["doc_sacado"] and op.get("doc_sacado"):
            groups[key]["doc_sacado"] = only_digits(op["doc_sacado"])
        groups[key]["notas"].append(op)
        groups[key]["total"] += op.get("valor_raw", Decimal("0"))
    result = []
    for g in groups.values():
        g["count"] = len(g["notas"])
        g["valor_total"] = _fmt_brl(g["total"])
        g["notas"].sort(key=lambda n: (n.get("data_vencimento") or "", n.get("nf") or ""))
        result.append(g)
    result.sort(key=lambda g: g["nome_sacado"].upper())
    return result


LIMITE_SOBRA_MIN = 100_000


def _evaluate_limite_operacao(montante: Decimal, limite_data: dict | None):
    if not limite_data:
        return "nao_validado", "Limites não validados"
    state = limite_data.get("state")
    if state == "processing":
        return "validando", "Validando…"
    if state in ("error", "ltc_expired"):
        return "insuficiente", "Limite insuficiente"
    limite = limite_data.get("limite_disp")
    if limite is None or state not in ("ok", "warn"):
        return "nao_validado", "Limites não validados"
    mont = int(montante.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if mont > limite:
        return "insuficiente", "Limite insuficiente"
    if (limite - mont) < LIMITE_SOBRA_MIN:
        return "quase", "Limite quase lá"
    return "ok", "Limite OK"


def _fmt_limite_int(val) -> str:
    if val is None:
        return "N/D"
    return f"R$ {int(val):,}".replace(",", ".")

BPM_CLIENT_DATA = {
    "Transdourada":             {"CNPJ":"01259730000174","PLATAFORMA":"2939","AG":"1643","CONTA":"99451-8"},
    "RPB":                      {"CNPJ":"07075892000139","PLATAFORMA":"8973","AG":"6627","CONTA":"06471-7"},
    "Posto Arinos":             {"CNPJ":"05798923000154","PLATAFORMA":"8250","AG":"1364","CONTA":"98355-9"},
    "Brasnorte":                {"CNPJ":"00514301000133","PLATAFORMA":"8250","AG":"1364","CONTA":"98654-5"},
    "Mirian Varzea":            {"CNPJ":"16519674000137","PLATAFORMA":"8250","AG":"1689","CONTA":"05136-3"},
    "Mirian Cuiaba":            {"CNPJ":"41240105000103","PLATAFORMA":"8250","AG":"1689","CONTA":"58145-0"},
    "Petrocal":                 {"CNPJ":"12781233000158","PLATAFORMA":"7948","AG":"8251","CONTA":"44190-6"},
    "Posto Sapucaia":           {"CNPJ":"22787055000126","PLATAFORMA":"0352","AG":"1334","CONTA":"57853-9"},
    "PetroMix":                 {"CNPJ":"05684913000198","PLATAFORMA":"7948","AG":"8251","CONTA":"99886-3"},
    "Auto Posto M Timbozao":    {"CNPJ":"04632746000179","PLATAFORMA":"0352","AG":"8296","CONTA":"18100-4"},
    "PetroVel":                 {"CNPJ":"01294927000144","PLATAFORMA":"7948","AG":"8251","CONTA":"99887-1"},
    "Posto Gasol Timbo III":    {"CNPJ":"32179707000101","PLATAFORMA":"0352","AG":"8296","CONTA":"06655-1"},
    "Posto Timbozao Itaperuna": {"CNPJ":"25032853000136","PLATAFORMA":"0352","AG":"1334","CONTA":"49413-3"},
    "Posto Pioneiro":           {"CNPJ":"23184831000166","PLATAFORMA":"0352","AG":"5255","CONTA":"12888-5"},
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
    "Petrocal":       ["PetroMix","PetroVel"],
    "Posto Sapucaia": ["Auto Posto M Timbozao","Posto Gasol Timbo III","Posto Timbozao Itaperuna","Posto Pioneiro"],
}

# ─── PRAZO CONFIG ────────────────────────────────────────────────────────────
PRAZO_CONFIG = {
    "RPB":                      {"max": 11, "min_alert": None},
    "Transdourada":             {"max": 90, "min_alert": 50},
    "Posto Sapucaia":           {"max": 15, "min_alert": 7},
    "Auto Posto M Timbozao":    {"max": 15, "min_alert": 7},
    "Posto Gasol Timbo III":    {"max": 15, "min_alert": 7},
    "Posto Timbozao Itaperuna": {"max": 15, "min_alert": 7},
    "Posto Pioneiro":           {"max": 15, "min_alert": 7},
    "Mirian Cuiaba":            {"max": 10, "min_alert": None},
    "Mirian Varzea":            {"max": 10, "min_alert": None},
    "Petrocal":                 {"max": 10, "min_alert": None},
    "PetroMix":                 {"max": 10, "min_alert": None},
    "PetroVel":                 {"max": 10, "min_alert": None},
}

MAPPED_CLIENTS = {"Transdourada","RPB","Posto Arinos","Mirian Varzea","Petrocal","Posto Sapucaia"}
MIRROR_CLIENTS = {
    "Brasnorte":"Posto Arinos","Mirian Cuiaba":"Mirian Varzea",
    "PetroMix":"Petrocal","PetroVel":"Petrocal",
    "Auto Posto M Timbozao":"Posto Sapucaia","Posto Gasol Timbo III":"Posto Sapucaia",
    "Posto Timbozao Itaperuna":"Posto Sapucaia","Posto Pioneiro":"Posto Sapucaia",
}

REGIAO_TRADER_ESPEC = {
    "21":("Debora","Vinicios Luz"),"22":("Thiago","Paula Costa"),
    "23":("Thiago","Paula Costa"),"24":("Gabriel","Luiz Gustavo Sarmento"),
    "25":("Giovanna","Lucas Capeli"),"26":("Gabriel","Vinicios Luz"),
    "28":("Thiago","Paula Costa"),"29":("Debora","Renata Leviski"),
    "30":("Adriana","Luiz Gustavo Sarmento"),"32":("Debora","Renata Leviski"),
    "33":("Giovanna","Lucas Capeli"),
}

RE_SPACES        = re.compile(r"\s+")
RE_CNPJ          = re.compile(r"(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})")
RE_CNPJ_LABEL    = re.compile(r"CNPJ[^\d]*([\d./-]{14,20})", re.IGNORECASE)
RE_VALOR_1       = re.compile(r"Valor\s+da\s+Opera[cç][aã]o[^\d]*(R\$\s*[\d\.,]+)", re.IGNORECASE)
RE_VALOR_2       = re.compile(r"R\$\s*[\d\.]{1,},\d{2}")
RE_VALOR_3       = re.compile(r"(R\$\s*[\d\.,]{3,})")
RE_PLATAFORMA    = re.compile(r"Plataforma\D{0,20}(\d{4})", re.IGNORECASE)
RE_REGIAO_PLAT   = re.compile(r"Regi[aã]o(?:\s+da)?\s+Plataforma\D{0,20}(\d{2})", re.IGNORECASE)
RE_REGIAO        = re.compile(r"Regi[aã]o\D{0,20}(\d{2})", re.IGNORECASE)
RE_SPREAD_1      = re.compile(r"(?:Taxa\/Spread|Spread)\D{0,25}([\d\.,]+)\s*%?", re.IGNORECASE)
RE_SPREAD_2      = re.compile(r"Spread\s+m[íi]nimo\D{0,25}([\d\.,]+)\s*%?", re.IGNORECASE)
RE_PRAZO_MIN_1   = re.compile(r"Prazo\s+M[ií]n[ií]mo\s+NF\D{0,25}(\d{1,4})", re.IGNORECASE)
RE_PRAZO_MAX_1   = re.compile(r"Prazo\s+M[áa]ximo\s+NF\D{0,25}(\d{1,4})", re.IGNORECASE)
RE_MODALIDADE_1  = re.compile(r"Modalidade[:\s]+([^\n\.;]+)", re.IGNORECASE)
RE_AUTORIZADAS_CSV   = re.compile(r"\bautorizadas\s+csv\b", re.IGNORECASE)
RE_AUTORIZADAS_SIS   = re.compile(r"\bautorizadas\s+sispag\b", re.IGNORECASE)
RE_RAZAO_1       = re.compile(r"Raz[aã]o\s+Social[:\s]*([^\n]+)", re.IGNORECASE)
RE_RAZAO_ATE     = re.compile(r"Raz[aã]o\s+Social\s*[:\s]*\s*(.+?)(?=\n\s*(?:CNPJ|Plataforma|Modalidade|Regi[aã]o|Conta\s+Corrente|Valor\s+da|Spread|Prazo)\b)", re.IGNORECASE|re.DOTALL)
RE_CONTA_AG_CC   = re.compile(r"\b(\d{4})\s*[/\s]\s*(\d{3,10}(?:[-‑–—−]\d)?)\b")
RE_CONTA_LABEL   = re.compile(r"conta\s+corrente(?:\s+do\s+cliente)?\s*[:\s-]*", re.IGNORECASE)
RE_LIQ_CRED      = re.compile(r"Cr[ée]dito\s+em\s+CC", re.IGNORECASE)
RE_PREMIO        = re.compile(r"com\s+pr[êe]mio", re.IGNORECASE)

def app_base_dir():
    return os.path.dirname(sys.executable) if getattr(sys,"frozen",False) else os.path.dirname(os.path.abspath(__file__))

def resource_path(p):
    return os.path.join(getattr(sys,"_MEIPASS",app_base_dir()), p)

def only_digits(s):
    return re.sub(r"\D","",s or "")

LIMITE_INVERTIDO_CNPJS = frozenset(
    only_digits(v["CNPJ"]) for v in BPM_CLIENT_DATA.values() if v.get("CNPJ")
)

def normalize_text_variants(t):
    return RE_SPACES.sub(" ",t).strip(), re.sub(r"\s+","",t or "")

def extract_text_from_pdf(p):
    r = PdfReader(p, strict=False)
    lo, pl = [], []
    for pg in r.pages:
        try:
            try: lo.append(pg.extract_text(extraction_mode="layout") or "")
            except TypeError: lo.append(pg.extract_text() or "")
        except: lo.append("")
        try: pl.append(pg.extract_text() or "")
        except: pl.append(lo[-1])
    return "\n".join(lo), "\n".join(pl)

def normalize_modalidade(m):
    s = RE_SPACES.sub(" ", (m or "").strip()); sl = s.lower().replace(" ","")
    if "sispag" in sl: return "Autorizadas SISPAG"
    if RE_AUTORIZADAS_SIS.search(s): return "Autorizadas SISPAG"
    if RE_AUTORIZADAS_CSV.search(s): return "Autorizadas CSV"
    if re.search(r"\bsispag\b",s,re.I): return "Autorizadas SISPAG"
    if re.search(r"\bcsv\b",s,re.I): return "Autorizadas CSV"
    return s

def infer_troca(m):
    ml = (m or "").lower().replace(" ","")
    if "sispag" in ml: return "Sispag"
    return "CSV"

def normalize_percent_br(p):
    if not p: return ""
    v = p.strip().replace(" ","").replace(",",".")
    try:
        n = float(re.sub(r"[^\d\.]","",v))
        return f"{int(n)},00% a.a." if n.is_integer() else f"{n:.2f}".replace(".",",")+"%  a.a."
    except:
        v = p.strip().replace(".",",")
        if not v.endswith("%"): v += "%"
        return v+("" if "a.a" in v.lower() else " a.a.")

def trader_espec_from_regiao(reg):
    d = only_digits(reg)
    if len(d) >= 2:
        return REGIAO_TRADER_ESPEC.get(d[-2:], ("",""))
    m = RE_REGIAO_PLAT.search(reg) or RE_REGIAO.search(reg)
    if m: return REGIAO_TRADER_ESPEC.get(m.group(1),("",""))
    return ("","")

def sanitize_razao(s):
    if not s: return s
    s = RE_SPACES.sub(" ", s).strip()
    s = re.sub(r"\s*\(\s*/[^)]+\)","",s)
    s = re.split(r"\bCNPJ\b",s,maxsplit=1,flags=re.I)[0]
    s = re.sub(r"\s*\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\s*$","",s)
    return RE_SPACES.sub(" ",s).strip(" :-–—.,;")

def _razao_stop(ln):
    s = (ln or "").strip().lower()
    if not s: return False
    return bool(re.match(r"^\s*(?:cnpj|plataforma|modalidade|regi[aã]o|conta\s+corrente|valor\s+da\s+opera|spread|prazo|liquida)\b",s)) \
           or (len(only_digits(s))>=14 and len(s)<40)

def extract_razao_social(t, lines, tn, tc, t_plain=None):
    for blk in ([t]+([t_plain] if t_plain else [])):
        if not blk: continue
        m = re.search(r"raz[aã]o\s+social\s*[:\s]*",blk,re.I)
        if m:
            rest = blk[m.end():]; end = len(rest)
            for pat in [re.compile(r"\bCNPJ\b",re.I), RE_CNPJ,
                        re.compile(r"\bPlataforma\b",re.I), re.compile(r"\bModalidade\b",re.I)]:
                rm = pat.search(rest)
                if rm and rm.start() < end: end = rm.start()
            chunk = rest[:end].strip()
            if "\n\n" in chunk: chunk = chunk.split("\n\n")[0].strip()
            if len(chunk) >= 2: return sanitize_razao(chunk)
    return None

def extract_cnpj(t, tn, tc):
    m = RE_CNPJ_LABEL.search(t)
    if m: return m.group(1).strip()
    m = RE_CNPJ.search(t)
    if m: return m.group(1).strip()
    m = re.search(r"cnpj.*?(\d{14})",tc,re.I)
    if m: return m.group(1)
    return None

def extract_conta_corrente(lines, tn, tc):
    for hay in (tn, tc):
        if not hay: continue
        m = RE_CONTA_AG_CC.search(hay)
        if m: return f"{m.group(1)} / {m.group(2)}"
    for i,line in enumerate(lines):
        if not re.search(r"conta\s+corrente",line,re.I): continue
        mm = RE_CONTA_LABEL.search(line)
        if not mm: continue
        rest = line[mm.end():].strip()
        if rest and not _razao_stop(rest): return rest
        if i+1 < len(lines):
            nxt = lines[i+1].strip()
            if nxt and not _razao_stop(nxt): return nxt
    return None

def extract_plataforma(t, tn, tc):
    m = RE_PLATAFORMA.search(tn) or RE_PLATAFORMA.search(t)
    return m.group(1) if m else None

def extract_regiao(t, tn, tc):
    m = RE_REGIAO_PLAT.search(tn) or RE_REGIAO_PLAT.search(t)
    if m: return m.group(1)
    m = RE_REGIAO.search(tn) or RE_REGIAO.search(t)
    return m.group(1) if m else None

def extract_valor(t, tn, tc):
    for pat in [RE_VALOR_1, RE_VALOR_2, RE_VALOR_3]:
        m = pat.search(tn)
        if m: return (m.group(1) if pat is RE_VALOR_1 else m.group(0)).strip()
    return None

def extract_spread(t, tn, tc):
    m = RE_SPREAD_1.search(tn) or RE_SPREAD_2.search(tn)
    return m.group(1).strip() if m else None

def extract_prazo_min(t, tn, tc):
    m = RE_PRAZO_MIN_1.search(tn)
    return m.group(1) if m else None

def extract_prazo_max(t, tn, tc):
    m = RE_PRAZO_MAX_1.search(tn)
    return m.group(1) if m else None

def extract_modalidade(t, tn, tc):
    for hay in (tn, t, tc):
        if not hay: continue
        m = RE_MODALIDADE_1.search(hay)
        if m: return normalize_modalidade(m.group(1).strip())
    return None

# ─── Validação de Alertas ────────────────────────────────────────────────────

def _find_client_for_prazo(doc_sacado, nome_sacado):
    """Match sacado to a client in PRAZO_CONFIG via CNPJ then name."""
    cnpj = only_digits(doc_sacado or "")
    if cnpj:
        for client_name, data in BPM_CLIENT_DATA.items():
            if only_digits(data.get("CNPJ", "")) == cnpj:
                return client_name
    key = _normalize_sacado_key(nome_sacado or "")
    for client in PRAZO_CONFIG:
        if _normalize_sacado_key(client) == key:
            return client
    for client in PRAZO_CONFIG:
        ck = _normalize_sacado_key(client)
        if ck in key or key in ck:
            return client
    return None


def _validate_nf(nf_str):
    """Alert if NF has leading zeros or is suspiciously low."""
    nf = (nf_str or "").strip()
    if not nf:
        return "NF vazia"
    if len(nf) > 1 and nf.startswith("0"):
        return f"NF com zeros à esquerda ({nf})"
    try:
        val = int(re.sub(r'\D', '', nf))
        if val <= 0:
            return f"NF inválida ({nf})"
    except ValueError:
        return f"NF não numérica ({nf})"
    return None


def _validate_valor(valor_raw):
    """Alert if value is below R$ 10.000,00."""
    if valor_raw is None or valor_raw < Decimal("10000"):
        return "Valor abaixo de R$ 10.000,00"
    return None


def _validate_prazo(nome_sacado, doc_sacado, prazo_str):
    """Alert if prazo exceeds max or is significantly below for known clients."""
    client = _find_client_for_prazo(doc_sacado, nome_sacado)
    if not client:
        return None
    config = PRAZO_CONFIG.get(client)
    if not config:
        return None
    try:
        prazo = int(re.sub(r'\D', '', prazo_str or ""))
    except (ValueError, TypeError):
        return f"Prazo inválido ({prazo_str})"
    if prazo > config["max"]:
        return f"Prazo ({prazo}d) acima do máximo ({config['max']}d) para {client}"
    min_alert = config.get("min_alert")
    if min_alert is not None and prazo < min_alert:
        return f"Prazo ({prazo}d) muito abaixo do esperado ({config['max']}d) para {client}"
    return None


def _validate_operation(op):
    """Validate a single operation and return list of alert reasons."""
    alerts = []
    nf_alert = _validate_nf(op.get("nf", ""))
    if nf_alert:
        alerts.append(nf_alert)
    valor_alert = _validate_valor(op.get("valor_raw"))
    if valor_alert:
        alerts.append(valor_alert)
    prazo_alert = _validate_prazo(
        op.get("nome_sacado", ""),
        op.get("doc_sacado", ""),
        op.get("prazo", ""))
    if prazo_alert:
        alerts.append(prazo_alert)
    return alerts


class BPMUserCancelled(Exception): pass


class ThreadSafeUIMixin:
    def _init_ui_queue(self):
        self._ui_q = queue.Queue()
        self._poll_ui_queue()

    def _ui(self, fn):
        try:
            self._ui_q.put_nowait(fn)
        except Exception:
            pass

    def _poll_ui_queue(self):
        try:
            while True:
                fn = self._ui_q.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass
        try:
            if self.winfo_exists():
                self.after(50, self._poll_ui_queue)
        except Exception:
            pass

def make_hairline(parent, orient="h", **kwargs):
    kw = {"bg": C["hair"]}
    kw.update(kwargs)
    if orient == "h":
        return tk.Frame(parent, height=1, **kw)
    else:
        return tk.Frame(parent, width=1, **kw)


def _make_dot(parent, color, size=10, bg=None):
    bg = bg or parent.cget("bg")
    c = tk.Canvas(parent, width=size + 4, height=size + 4,
                  bg=bg, highlightthickness=0, bd=0)
    c.create_oval(2, 2, size + 2, size + 2, fill=color, outline="")
    return c


def _canvas_round_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    r = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
    points = [
        x1 + r, y1, x2 - r, y1,
        x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r,
        x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=36, **kwargs)


FRAME_LABELS = {
    "Home":              "Início",
    "Rotinas":           "Rotinas",
    "Share":             "Cadastro Share",
    "BPM_CONFIG":        "Configurar BPM",
    "BPM":               "BPM — Operações",
    "OperacoesInvertido":"Operações Invertido",
    "LimitesInvertido":  "Limites Invertido",
    "AnalisarOperacoes": "Analisar Operações",
}


def _icon_png_path():
    for path in (resource_path(ICON_FILENAME), os.path.join(app_base_dir(), ICON_FILENAME)):
        if os.path.isfile(path):
            return path
    return None


def _png_to_ico_bytes(png_path):
    with open(png_path, "rb") as f:
        png_data = f.read()
    if len(png_data) < 24 or png_data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    width = int.from_bytes(png_data[16:20], "big")
    height = int.from_bytes(png_data[20:24], "big")
    w_byte = 0 if width >= 256 else width
    h_byte = 0 if height >= 256 else height
    image_offset = 6 + 16
    header = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", w_byte, h_byte, 0, 0, 1, 32, len(png_data), image_offset)
    return header + entry + png_data


def _build_ico_from_png(png_path, ico_path):
    try:
        from PIL import Image
        img = Image.open(png_path).convert("RGBA")
        img.save(ico_path, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
        return True
    except Exception:
        return False


def _ensure_ico_path():
    cached = os.path.join(app_base_dir(), "itaulogo.ico")
    png_path = _icon_png_path()
    if png_path:
        needs_build = not os.path.isfile(cached)
        if needs_build and _build_ico_from_png(png_path, cached):
            return cached
    if os.path.isfile(cached):
        return cached
    if not png_path:
        return None
    ico_bytes = _png_to_ico_bytes(png_path)
    if not ico_bytes:
        return None
    try:
        with open(cached, "wb") as f:
            f.write(ico_bytes)
        return cached
    except OSError:
        try:
            fd, tmp = tempfile.mkstemp(suffix=".ico", prefix="mesa_itau_")
            os.close(fd)
            with open(tmp, "wb") as f:
                f.write(ico_bytes)
            return tmp
        except OSError:
            return None


def apply_taskbar_presence(root):
    if sys.platform != "win32":
        return
    try:
        GWL_EXSTYLE = -20
        WS_EX_APPWINDOW = 0x00040000
        WS_EX_TOOLWINDOW = 0x00000080
        SW_HIDE = 0
        SW_SHOW = 5
        root.update_idletasks()
        hwnd = _window_hwnd(root)
        if not hwnd:
            return
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        new_style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
        if new_style != style:
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_style)
        ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
        ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
        root.lift()
    except Exception:
        pass


def apply_window_icon(root):
    ico_path = _ensure_ico_path()
    if not ico_path:
        return
    try:
        root.iconbitmap(default=ico_path)
    except Exception:
        pass
    if sys.platform == "win32":
        try:
            root.update_idletasks()
            hwnd = _window_hwnd(root)
            if not hwnd:
                return
            IMAGE_ICON = 1
            LR_LOADFROMFILE = 0x10
            WM_SETICON = 0x0080
            for size in (16, 32):
                hicon = ctypes.windll.user32.LoadImageW(
                    None, ico_path, IMAGE_ICON, size, size, LR_LOADFROMFILE,
                )
                if hicon:
                    which = 0 if size <= 16 else 1
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, which, hicon)
        except Exception:
            pass


def apply_windows_shell(root):
    apply_taskbar_presence(root)
    apply_window_icon(root)


def apply_modern_window_chrome(root):
    if sys.platform != "win32":
        return
    try:
        hwnd = _window_hwnd(root)
        dwm = ctypes.windll.dwmapi
        dark = ctypes.c_int(1)
        round_pref = ctypes.c_int(2)
        dwm.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(dark), ctypes.sizeof(dark))
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(round_pref), ctypes.sizeof(round_pref))
    except Exception:
        pass


def _window_hwnd(root):
    root.update_idletasks()
    return ctypes.windll.user32.GetParent(root.winfo_id())


def apply_frameless_resize(root):
    if sys.platform != "win32":
        return
    try:
        hwnd = _window_hwnd(root)
        gwl_style = -16
        style = ctypes.windll.user32.GetWindowLongW(hwnd, gwl_style)
        style |= 0x00040000
        style |= 0x00020000
        style |= 0x00010000
        ctypes.windll.user32.SetWindowLongW(hwnd, gwl_style, style)
    except Exception:
        pass


def start_native_window_drag(root):
    if sys.platform != "win32":
        return False
    try:
        hwnd = _window_hwnd(root)
        if not hwnd:
            return False
        WM_NCLBUTTONDOWN = 0x00A1
        HTCAPTION = 2
        user32 = ctypes.windll.user32
        user32.ReleaseCapture()
        user32.PostMessageW(ctypes.c_void_p(hwnd), WM_NCLBUTTONDOWN,
                            ctypes.c_void_p(HTCAPTION), ctypes.c_void_p(0))
        return True
    except Exception:
        return False


class AppTitleBar(tk.Frame):
    BG = "#1c1c1c"
    HEIGHT = 36

    def __init__(self, parent, root):
        super().__init__(parent, bg=self.BG, height=self.HEIGHT)
        self.pack_propagate(False)
        self.root = root
        self._drag_offset = None
        self._maximized = False

        row = tk.Frame(self, bg=self.BG)
        row.pack(fill="both", expand=True)

        left = tk.Frame(row, bg=self.BG)
        left.pack(side="left", fill="y", padx=(14, 0))

        tk.Label(left, text="Mesa", bg=self.BG, fg=C["ink"],
                 font=("Segoe UI", 10, "bold")).pack(side="left", pady=8)
        tk.Label(left, text="Itaú", bg=self.BG, fg=C["accent"],
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=(3, 0), pady=8)
        tk.Label(left, text="·", bg=self.BG, fg="#404040",
                 font=("Segoe UI", 9)).pack(side="left", padx=8, pady=8)
        self._module_lbl = tk.Label(left, text="Início", bg=self.BG, fg="#8a8a8a",
                                    font=("Segoe UI", 9))
        self._module_lbl.pack(side="left", pady=8)

        controls = tk.Frame(row, bg=self.BG)
        controls.pack(side="right", fill="y")

        self._btn_min = self._win_btn(controls, "─", self._minimize)
        self._btn_max = self._win_btn(controls, "□", self._toggle_maximize)
        self._btn_close = self._win_btn(controls, "✕", self.root.destroy, close=True)

        tk.Frame(self, bg="#2e2e2e", height=1).pack(fill="x", side="bottom")

        self._bind_drag(self)
        self._bind_drag(row)
        self._bind_drag(left)
        for w in left.winfo_children():
            if w is not self._module_lbl:
                self._bind_drag(w)
        self._bind_drag(self._module_lbl)
        self.bind("<Double-Button-1>", lambda _e: self._toggle_maximize())

    def _win_btn(self, parent, text, command, close=False):
        hover = "#c95f5f" if close else "#333333"
        lbl = tk.Label(parent, text=text, bg=self.BG, fg="#9a9a9a",
                       font=("Segoe UI", 9), width=4, cursor="hand2")
        lbl.pack(side="left", fill="y")
        lbl.bind("<Button-1>", lambda _e: command())
        lbl.bind("<Enter>", lambda _e, l=lbl, h=hover: l.configure(bg=h, fg="#f2f2f2"))
        lbl.bind("<Leave>", lambda _e, l=lbl: l.configure(bg=self.BG, fg="#9a9a9a"))
        return lbl

    def _bind_drag(self, widget):
        widget.bind("<ButtonPress-1>", self._start_drag, add="+")
        if sys.platform != "win32":
            widget.bind("<B1-Motion>", self._on_drag_fallback, add="+")

    def _start_drag(self, event):
        if start_native_window_drag(self.root):
            return
        if self._maximized:
            return
        self._drag_offset = (event.x_root - self.root.winfo_x(),
                             event.y_root - self.root.winfo_y())

    def _on_drag_fallback(self, event):
        if not self._drag_offset or self._maximized:
            return
        ox, oy = self._drag_offset
        self.root.geometry(f"+{event.x_root - ox}+{event.y_root - oy}")

    def _minimize(self):
        if sys.platform == "win32":
            try:
                ctypes.windll.user32.ShowWindow(_window_hwnd(self.root), 6)
                return
            except Exception:
                pass
        self.root.iconify()

    def _toggle_maximize(self):
        if sys.platform == "win32":
            try:
                hwnd = _window_hwnd(self.root)
                if self._maximized:
                    ctypes.windll.user32.ShowWindow(hwnd, 9)
                    self._maximized = False
                    self._btn_max.configure(text="□")
                else:
                    ctypes.windll.user32.ShowWindow(hwnd, 3)
                    self._maximized = True
                    self._btn_max.configure(text="❐")
                return
            except Exception:
                pass
        if self._maximized:
            self.root.state("normal")
            self._maximized = False
            self._btn_max.configure(text="□")
        else:
            self.root.state("zoomed")
            self._maximized = True
            self._btn_max.configure(text="❐")

    def set_module(self, frame_name):
        self._module_lbl.configure(text=FRAME_LABELS.get(frame_name, frame_name))


class AppStatusBar(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"], height=30)
        self.pack_propagate(False)
        self.controller = controller

        tk.Frame(self, bg="#303030", height=1).pack(fill="x")

        row = tk.Frame(self, bg=C["bg"])
        row.pack(fill="both", expand=True, padx=16)

        left = tk.Frame(row, bg=C["bg"])
        left.pack(side="left", fill="y")

        tk.Label(left, text="Mesa Itaú", bg=C["bg"], fg=C["ink_faint"],
                 font=("Segoe UI", 8)).pack(side="left", pady=4)
        tk.Label(left, text="·", bg=C["bg"], fg="#3a3a3a",
                 font=("Segoe UI", 8)).pack(side="left", padx=6, pady=4)
        tk.Label(left, text="Risco Sacado", bg=C["bg"], fg=C["ink_faint"],
                 font=("Segoe UI", 8)).pack(side="left", pady=4)

        right = tk.Frame(row, bg=C["bg"])
        right.pack(side="right", fill="y")

        self._clock_lbl = tk.Label(right, text="", bg=C["bg"], fg=C["ink_faint"],
                                   font=("Segoe UI", 8))
        self._clock_lbl.pack(side="right", padx=(12, 0), pady=4)

        self._module_lbl = tk.Label(right, text="Início", bg=C["bg"], fg="#6b6b6b",
                                    font=("Segoe UI", 8))
        self._module_lbl.pack(side="right", pady=4)

        self._tick_clock()

    def set_module(self, frame_name):
        label = FRAME_LABELS.get(frame_name, frame_name)
        self._module_lbl.configure(text=label)

    def _tick_clock(self):
        now = datetime.now()
        self._clock_lbl.configure(text=now.strftime("%d/%m/%Y  %H:%M"))
        self.after(30_000, self._tick_clock)


def styled_label(parent, text, size=10, weight="normal", color=None, **kwargs):
    return tk.Label(parent, text=text,
                    font=("Segoe UI", size, weight),
                    fg=color or C["ink"],
                    bg=kwargs.pop("bg", C["surface"]),
                    **kwargs)


def styled_button(parent, text, command, accent=False, danger=False, small=False, **kwargs):
    if accent:
        bg  = C["accent_dim"];  fg  = C["accent"]
        abg = C["accent"];      afg = C["bg"]
    elif danger:
        bg  = C["err_dim"];     fg  = C["err"]
        abg = C["err"];         afg = C["bg"]
    else:
        bg  = C["surface2"];    fg  = C["ink_muted"]
        abg = C["surface3"];    afg = C["ink"]
    pad = (7, 3) if small else (13, 6)
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, activebackground=abg, activeforeground=afg,
                    font=("Segoe UI", 8 if small else 9),
                    relief="flat", bd=0, padx=pad[0], pady=pad[1],
                    cursor="hand2", **kwargs)
    btn.bind("<Enter>", lambda _: btn.configure(bg=abg, fg=afg))
    btn.bind("<Leave>", lambda _: btn.configure(bg=bg,  fg=fg))
    return btn


def styled_button_limite(parent, text, command, variant="warn", small=False, **kwargs):
    styles = {
        "ok":  (C["ok_dim"],  C["ok"],  C["ok"],  C["bg"]),
        "warn":("#3d3520",    C["warn"], C["warn"], C["bg"]),
        "err": (C["err_dim"],  C["err"], C["err"], C["bg"]),
        "idle":(C["surface2"], C["ink_muted"], C["surface3"], C["ink"]),
    }
    bg, fg, abg, afg = styles.get(variant, styles["warn"])
    pad = (7, 3) if small else (13, 6)
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, activebackground=abg, activeforeground=afg,
                    font=("Segoe UI", 8 if small else 9),
                    relief="flat", bd=0, padx=pad[0], pady=pad[1],
                    cursor="hand2", **kwargs)
    btn._limite_variant = variant
    btn._limite_bg = bg
    btn._limite_fg = fg
    btn._limite_abg = abg
    btn._limite_afg = afg
    btn.bind("<Enter>", lambda _: btn.configure(bg=abg, fg=afg))
    btn.bind("<Leave>", lambda _: btn.configure(bg=bg, fg=fg))
    return btn


def _set_limite_button(btn, text, variant):
    styles = {
        "ok":  (C["ok_dim"],  C["ok"],  C["ok"],  C["bg"]),
        "warn":("#3d3520",    C["warn"], C["warn"], C["bg"]),
        "err": (C["err_dim"],  C["err"], C["err"], C["bg"]),
        "idle":(C["surface2"], C["ink_muted"], C["surface3"], C["ink"]),
    }
    bg, fg, abg, afg = styles.get(variant, styles["warn"])
    btn.configure(text=text, bg=bg, fg=fg, activebackground=abg, activeforeground=afg)
    btn._limite_variant = variant
    btn._limite_bg = bg
    btn._limite_fg = fg
    btn._limite_abg = abg
    btn._limite_afg = afg
    btn.bind("<Enter>", lambda _: btn.configure(bg=abg, fg=afg))
    btn.bind("<Leave>", lambda _: btn.configure(bg=bg, fg=fg))


def styled_entry(parent, textvariable=None, width=20, show=None, **kwargs):
    return tk.Entry(parent, textvariable=textvariable, width=width,
                    show=show or "",
                    bg=C["surface2"], fg=C["ink"],
                    insertbackground=C["accent"],
                    relief="flat", highlightthickness=1,
                    highlightbackground=C["hair"],
                    highlightcolor=C["accent"],
                    font=("Segoe UI", 10), **kwargs)


def card_frame(parent, **kwargs):
    kw = {"bg": C["surface"], "highlightthickness": 1,
          "highlightbackground": C["hair"], "bd": 0}
    kw.update(kwargs)
    return tk.Frame(parent, **kw)


def eyebrow_label(parent, text, bg=None):
    bg = bg or C["bg"]
    return tk.Label(parent, text=text, bg=bg, fg=C["ink_faint"],
                    font=("Segoe UI", 7, "bold"))


def section_divider(parent, text="", bg=None):
    bg = bg or C["bg"]
    row = tk.Frame(parent, bg=bg)
    if text:
        tk.Label(row, text=text, bg=bg, fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(side="left")
        spacer = tk.Frame(row, bg=C["hair"], height=1)
        spacer.pack(side="left", fill="x", expand=True, padx=(10, 0), pady=(5, 0))
    else:
        tk.Frame(row, bg=C["hair"], height=1).pack(fill="x")
    return row

class MinimalScrollbar(tk.Canvas):
    THUMB_MIN = 28

    def __init__(self, parent, command=None, bg=None, width=6, **kwargs):
        self._track_bg = bg or C["bg"]
        super().__init__(
            parent, width=width, highlightthickness=0, bd=0,
            bg=self._track_bg, cursor="arrow", **kwargs,
        )
        self._command = command
        self._first = 0.0
        self._last = 1.0
        self._thumb_fill = "#5c5c5c"
        self._thumb_hover = "#787878"
        self._drag_y = 0
        self._thumb_rect = None

        self.bind("<Configure>", self._redraw, add="+")
        self.bind("<Button-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", lambda _e: self._paint_thumb(self._thumb_hover))
        self.bind("<Leave>", lambda _e: self._paint_thumb(self._thumb_fill))

    def set(self, first, last):
        f, l = float(first), float(last)
        if f == self._first and l == self._last:
            return
        self._first, self._last = f, l
        self._redraw()

    def _visible(self):
        return self._first > 0.001 or self._last < 0.999

    def _thumb_geometry(self):
        h = max(self.winfo_height(), 1)
        w = max(self.winfo_width(), 1)
        span = max(self._last - self._first, 0.001)
        thumb_h = max(self.THUMB_MIN, int(h * span))
        thumb_y = int(h * self._first)
        if thumb_y + thumb_h > h:
            thumb_y = max(0, h - thumb_h)
        return w, h, thumb_y, thumb_h

    def _redraw(self, _event=None):
        self.delete("all")
        self._thumb_rect = None
        if not self._visible():
            return
        w, _h, thumb_y, thumb_h = self._thumb_geometry()
        margin = 1
        x0, x1 = margin, max(margin + 2, w - margin)
        radius = (x1 - x0) / 2
        if thumb_h <= (x1 - x0):
            self.create_oval(x0, thumb_y, x1, thumb_y + (x1 - x0),
                             fill=self._thumb_fill, outline="", tags="thumb")
            self.create_oval(x0, thumb_y + thumb_h - (x1 - x0), x1, thumb_y + thumb_h,
                             fill=self._thumb_fill, outline="", tags="thumb")
            self._thumb_rect = (x0, thumb_y, x1, thumb_y + thumb_h)
        else:
            mid_top = thumb_y + radius
            mid_bot = thumb_y + thumb_h - radius
            self.create_oval(x0, thumb_y, x1, mid_top + radius,
                             fill=self._thumb_fill, outline="", tags="thumb")
            self.create_rectangle(x0, mid_top, x1, mid_bot,
                                  fill=self._thumb_fill, outline="", tags="thumb")
            self.create_oval(x0, mid_bot - radius, x1, thumb_y + thumb_h,
                             fill=self._thumb_fill, outline="", tags="thumb")
            self._thumb_rect = (x0, thumb_y, x1, thumb_y + thumb_h)

    def _paint_thumb(self, color):
        for item in self.find_withtag("thumb"):
            self.itemconfig(item, fill=color)

    def _on_press(self, event):
        if not self._command or not self._visible():
            return
        w, h, thumb_y, thumb_h = self._thumb_geometry()
        if self._thumb_rect and self._thumb_rect[1] <= event.y <= self._thumb_rect[3]:
            self._drag_y = event.y - thumb_y
            return
        if event.y > thumb_y + thumb_h:
            self._command("scroll", 1, "pages")
        elif event.y < thumb_y:
            self._command("scroll", -1, "pages")

    def _on_drag(self, event):
        if not self._command or not self._visible():
            return
        w, h, _thumb_y, thumb_h = self._thumb_geometry()
        span = max(h - thumb_h, 1)
        frac = (event.y - self._drag_y) / span
        self._command("moveto", max(0.0, min(1.0, frac)))

    def _on_release(self, _event):
        self._drag_y = 0


def bind_text_mousewheel(text_widget):
    def _mw(event):
        try:
            if not text_widget.winfo_exists():
                return
        except tk.TclError:
            return
        if getattr(event, "delta", 0):
            text_widget.yview_scroll(int(-event.delta / 120), "units")
        elif event.num == 4:
            text_widget.yview_scroll(-3, "units")
        elif event.num == 5:
            text_widget.yview_scroll(3, "units")
    text_widget.bind("<MouseWheel>", _mw)
    text_widget.bind("<Button-4>", _mw)
    text_widget.bind("<Button-5>", _mw)


class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg=None, **kwargs):
        bg = bg or C["bg"]
        super().__init__(parent, bg=bg, **kwargs)
        self._bg = bg
        self._wheel_roots = []
        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self._vbar = MinimalScrollbar(self, command=self._canvas.yview, bg=bg, width=6)
        self._canvas.configure(yscrollcommand=self._vbar.set)
        self._vbar.pack(side="right", fill="y", padx=(0, 2), pady=4)
        self._canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._canvas, bg=bg)
        self._win = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner)
        self._canvas.bind("<Configure>", self._on_canvas)
        self.bind("<Destroy>", self._on_destroy)
        self._wheel_roots = [self]
        self._scroll_job = None
        self.refresh_bindings()

    def _canvas_alive(self):
        try:
            return bool(self._canvas.winfo_exists())
        except tk.TclError:
            return False

    def _scroll_mousewheel(self, event):
        if not self._canvas_alive():
            return
        try:
            if getattr(event, "delta", 0):
                self._canvas.yview_scroll(int(-event.delta / 120), "units")
            elif event.num == 4:
                self._canvas.yview_scroll(-3, "units")
            elif event.num == 5:
                self._canvas.yview_scroll(3, "units")
        except tk.TclError:
            pass

    def _bind_mousewheel_tree(self, widget):
        if isinstance(widget, (tk.Text, tk.Listbox)):
            return
        widget.bind("<MouseWheel>", self._scroll_mousewheel)
        widget.bind("<Button-4>", self._scroll_mousewheel)
        widget.bind("<Button-5>", self._scroll_mousewheel)
        for child in widget.winfo_children():
            self._bind_mousewheel_tree(child)

    def refresh_bindings(self):
        if not self._canvas_alive():
            return
        for root in self._wheel_roots:
            try:
                if root.winfo_exists():
                    self._bind_mousewheel_tree(root)
            except tk.TclError:
                pass

    def link_wheel(self, container):
        if container not in self._wheel_roots:
            self._wheel_roots.append(container)
        self.refresh_bindings()

    def _on_inner(self, _event):
        if not self._canvas_alive():
            return
        if getattr(self, "_scroll_job", None):
            try:
                self.after_cancel(self._scroll_job)
            except Exception:
                pass
        self._scroll_job = self.after(16, self._sync_scrollregion)

    def _sync_scrollregion(self):
        self._scroll_job = None
        if not self._canvas_alive():
            return
        try:
            self._canvas.update_idletasks()
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        except tk.TclError:
            pass

    def _on_canvas(self, event):
        if self._canvas_alive():
            self._canvas.itemconfigure(self._win, width=event.width)

    def _on_destroy(self, event):
        pass


class Sidebar(tk.Frame):
    NAV = [
        ("Home",              "⌂",  "Início"),
        ("Rotinas",           "◈",  "Rotinas"),
        ("Share",             "⊕",  "Cadastro Share"),
        ("BPM",               "⚡",  "BPM"),
        ("OperacoesInvertido","⬡",  "Operações Invertido"),
    ]

    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=C["surface"], width=210, **kwargs)
        self.pack_propagate(False)
        self.controller = controller
        self._btns = {}
        self._build()

    def _build(self):
        top = tk.Frame(self, bg=C["surface"])
        top.pack(fill="x", padx=18, pady=(14, 0))

        logo_row = tk.Frame(top, bg=C["surface"])
        logo_row.pack(fill="x")
        tk.Label(logo_row, text="Mesa", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Label(logo_row, text=" Itaú", bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Label(top, text="Risco Sacado", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(3, 0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(18, 12))

        nav_outer = tk.Frame(self, bg=C["surface"])
        nav_outer.pack(fill="both", expand=True, anchor="n")

        for name, icon, label in self.NAV:
            row = tk.Frame(nav_outer, bg=C["surface"], cursor="hand2")
            row.pack(fill="x", pady=1, padx=6)

            bar = tk.Frame(row, bg=C["surface"], width=3)
            bar.pack(side="left", fill="y")

            inner = tk.Frame(row, bg=C["surface"], padx=8, pady=7)
            inner.pack(side="left", fill="x", expand=True)

            icon_lbl = tk.Label(inner, text=icon, bg=C["surface"],
                                fg=C["ink_faint"], font=("Segoe UI", 12), width=2, anchor="w")
            icon_lbl.pack(side="left")

            text_lbl = tk.Label(inner, text=label, bg=C["surface"],
                                fg=C["ink_muted"], font=("Segoe UI", 9), anchor="w")
            text_lbl.pack(side="left", padx=(6, 0))

            def _click(n=name): self.controller.show_frame(n)

            def _enter(e, r=row, inn=inner, il=icon_lbl, tl=text_lbl):
                active = getattr(self.controller, "_active_frame", None)
                n_     = ""
                for _n, _b in self._btns.items():
                    if _b["row"] is r: n_ = _n; break
                if n_ != active:
                    for w in (r, inn, il, tl):
                        try: w.configure(bg=C["surface2"])
                        except: pass

            def _leave(e, r=row, inn=inner, il=icon_lbl, tl=text_lbl):
                active = getattr(self.controller, "_active_frame", None)
                n_     = ""
                for _n, _b in self._btns.items():
                    if _b["row"] is r: n_ = _n; break
                if n_ != active:
                    for w in (r, inn, il, tl):
                        try: w.configure(bg=C["surface"])
                        except: pass

            for w in (row, inner, icon_lbl, text_lbl):
                w.bind("<Button-1>", lambda _, n=name: _click(n))
                w.bind("<Enter>",    _enter)
                w.bind("<Leave>",    _leave)

            self._btns[name] = {
                "row": row, "inner": inner,
                "icon": icon_lbl, "text": text_lbl, "bar": bar
            }

    def set_active(self, name):
        for n, w in self._btns.items():
            is_active = (n == name)
            row_bg  = C["surface2"] if is_active else C["surface"]
            icon_fg = C["accent"]   if is_active else C["ink_faint"]
            text_fg = C["ink"]      if is_active else C["ink_muted"]
            bar_bg  = C["accent"]   if is_active else C["surface"]

            w["row"].configure(bg=row_bg)
            w["inner"].configure(bg=row_bg)
            w["icon"].configure(bg=row_bg, fg=icon_fg)
            w["text"].configure(bg=row_bg, fg=text_fg)
            w["bar"].configure(bg=bar_bg)

class HomeFrame(tk.Frame):
    MODULES = [
        {"name": "Cadastro Share",     "sub": "Extração e análise de PDF",          "icon": "⊕", "frame": "Share",            "color": "#5a9e72"},
        {"name": "BPM",                "sub": "Abertura de solicitações",            "icon": "⚡", "frame": "BPM_CONFIG",       "color": "#EC7000"},
        {"name": "Operações Invertido","sub": "Limites, LTC e análise de planilhas", "icon": "⬡", "frame": "OperacoesInvertido","color": "#c87941"},
        {"name": "Rotinas",            "sub": "Sequências configuráveis",            "icon": "◈", "frame": "Rotinas",          "color": "#8b72c9"},
    ]

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._build()

    def _build(self):
        self._sf = ScrollableFrame(self)
        self._sf.pack(fill="both", expand=True)
        inner = self._sf.inner
        inner.configure(bg=C["bg"])
        inner.columnconfigure(0, weight=1)

        greet = tk.Frame(inner, bg=C["bg"])
        greet.pack(fill="x", padx=44, pady=(40, 0))

        now  = datetime.now()
        hour = now.hour
        saudacao = "Bom dia" if hour < 12 else ("Boa tarde" if hour < 18 else "Boa noite")

        eyebrow_label(greet, "MESA DE OPERAÇÕES").pack(anchor="w")
        tk.Label(greet, text=f"{saudacao}.", bg=C["bg"], fg=C["ink"],
                 font=("Segoe UI", 26, "bold"), anchor="w").pack(anchor="w", pady=(6, 0))
        tk.Label(greet,
                 text=format_data_pt_br(now),
                 bg=C["bg"], fg=C["ink_muted"],
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 0))

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=44, pady=(28, 26))

        eyebrow_label(inner, "MÓDULOS").pack(anchor="w", padx=44, pady=(0, 14))

        grid = tk.Frame(inner, bg=C["bg"])
        grid.pack(fill="x", padx=44)
        grid.columnconfigure(0, weight=1, uniform="m")
        grid.columnconfigure(1, weight=1, uniform="m")

        for i, mod in enumerate(self.MODULES):
            r, c = divmod(i, 2)
            self._make_module_card(grid, mod, r, c)

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=44, pady=(30, 24))

        eyebrow_label(inner, "ATALHOS RÁPIDOS").pack(anchor="w", padx=44, pady=(0, 10))

        quick = tk.Frame(inner, bg=C["bg"])
        quick.pack(fill="x", padx=44, pady=(0, 40))

        links = [
            ("Nova solicitação BPM",    "BPM_CONFIG"),
            ("Operações Invertido",      "OperacoesInvertido"),
            ("Extrair dados de PDF",     "Share"),
            ("Gerenciar rotinas",        "Rotinas"),
        ]
        for label, frame in links:
            row_f = tk.Frame(quick, bg=C["bg"])
            row_f.pack(fill="x", pady=1)

            arrow = tk.Label(row_f, text="→", bg=C["bg"], fg=C["ink_faint"],
                             font=("Segoe UI", 9))
            arrow.pack(side="left", padx=(0, 6))

            btn = tk.Button(row_f, text=label,
                            command=lambda f=frame: self.controller.show_frame(f),
                            bg=C["bg"], fg=C["ink_muted"],
                            activebackground=C["bg"], activeforeground=C["ink"],
                            font=("Segoe UI", 9),
                            relief="flat", bd=0, anchor="w", padx=0, pady=4,
                            cursor="hand2")
            btn.pack(side="left")
            btn.bind("<Enter>", lambda e, b=btn, a=arrow: (
                b.configure(fg=C["accent"]), a.configure(fg=C["accent"])))
            btn.bind("<Leave>", lambda e, b=btn, a=arrow: (
                b.configure(fg=C["ink_muted"]), a.configure(fg=C["ink_faint"])))

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=44, pady=(28, 22))
        eyebrow_label(inner, "ROTINAS DE HOJE").pack(anchor="w", padx=44, pady=(0, 8))

        self._rot_container = tk.Frame(inner, bg=C["bg"])
        self._rot_container.pack(fill="x", padx=44, pady=(0, 40))
        self._build_rotinas_hoje()

    def on_show(self):
        self.refresh_rotinas()

    def refresh_rotinas(self):
        if not hasattr(self, "_rot_container"):
            return
        for w in self._rot_container.winfo_children():
            w.destroy()
        self._build_rotinas_hoje()

    def _build_rotinas_hoje(self):
        parent = self._rot_container
        data   = RotinasData.get()
        rots   = data.today_rotinas()
        done_n, total = data.today_stats()

        if not rots:
            tk.Label(parent, text="Nenhuma rotina para hoje — configure em Rotinas.",
                     bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 9)).pack(anchor="w")
            return

        prog_row = tk.Frame(parent, bg=C["bg"])
        prog_row.pack(fill="x", pady=(0, 10))
        prog_color = C["ok"] if done_n == total else C["accent"]
        tk.Label(prog_row, text=f"{done_n}/{total} concluída{'s' if total > 1 else ''}",
                 bg=C["bg"], fg=prog_color,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        lnk = tk.Button(prog_row, text="ver todas →",
                        command=lambda: self.controller.show_frame("Rotinas"),
                        bg=C["bg"], fg=C["ink_faint"],
                        activebackground=C["bg"], activeforeground=C["accent"],
                        font=("Segoe UI", 8), relief="flat", bd=0, padx=0,
                        cursor="hand2")
        lnk.pack(side="right")
        lnk.bind("<Enter>", lambda e: lnk.configure(fg=C["accent"]))
        lnk.bind("<Leave>", lambda e: lnk.configure(fg=C["ink_faint"]))

        sorted_rots = sorted(rots, key=RotinasData._rot_sort_key)
        for rot in sorted_rots[:6]:
            self._make_home_rot_row(parent, rot, data)

        if len(rots) > 6:
            tk.Label(parent,
                     text=f"+ {len(rots)-6} rotinas — veja todas em Rotinas",
                     bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 0))

    def _make_home_rot_row(self, parent, rot, data):
        from tkinter import font as tkfont
        rid   = rot["id"]
        color = rot.get("cor", C["accent"])
        nome  = rot.get("nome", "Rotina")
        _als  = rot.get("alertas") or []
        hora  = _als[0] if _als else (rot.get("hora_alerta") or "")

        row = tk.Frame(parent, bg=C["bg"], pady=4)
        row.pack(fill="x")

        chk = tk.Canvas(row, width=16, height=16,
                        bg=C["bg"], highlightthickness=0, bd=0, cursor="hand2")
        chk.pack(side="left", padx=(0, 8))

        if hora:
            tk.Label(row, text=hora, bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 7), width=5, anchor="w").pack(side="left")

        _make_dot(row, color, 6, bg=C["bg"]).pack(side="left", padx=(0, 6))

        name_lbl = tk.Label(row, text=nome, bg=C["bg"],
                            fg=C["ink"], font=("Segoe UI", 9), anchor="w")
        name_lbl.pack(side="left")

        def _refresh_chk(done):
            chk.delete("all")
            if done:
                chk.create_oval(0, 0, 15, 15, fill=color, outline="")
                chk.create_text(8, 8, text="✓", fill=C["bg"],
                                font=("Segoe UI", 7, "bold"))
                name_lbl.configure(fg=C["ink_faint"],
                                   font=("Segoe UI", 9, "overstrike"))
            else:
                chk.create_oval(0, 0, 15, 15, fill="",
                                outline=C["ink_faint"], width=1.5)
                name_lbl.configure(fg=C["ink"], font=("Segoe UI", 9))

        _refresh_chk(data.is_done(rid))

        def toggle(e=None):
            new = not data.is_done(rid)
            data.set_done(rid, new)
            _refresh_chk(new)
            self.refresh_rotinas()

        for w in [row, chk, name_lbl]:
            try: w.bind("<Button-1>", toggle)
            except Exception: pass

    def _make_module_card(self, parent, mod, row, col):
        dot_color = mod.get("color", C["accent"])
        pad = (0, 6) if col == 0 else (6, 0)

        outer = tk.Frame(parent, bg=C["surface"],
                         highlightthickness=1, highlightbackground=C["hair"],
                         cursor="hand2")
        outer.grid(row=row, column=col, sticky="nsew", padx=pad, pady=6)

        top_line = tk.Frame(outer, bg=C["hair"], height=2)
        top_line.pack(fill="x")

        body = tk.Frame(outer, bg=C["surface"], padx=18, pady=16)
        body.pack(fill="both", expand=True)

        icon_row = tk.Frame(body, bg=C["surface"])
        icon_row.pack(fill="x")
        dot = _make_dot(icon_row, dot_color, size=8, bg=C["surface"])
        dot.pack(side="left", pady=(3, 0))
        tk.Label(icon_row, text=mod["icon"], bg=C["surface"], fg=dot_color,
                 font=("Segoe UI", 16)).pack(side="left", padx=(6, 0))

        name_lbl = tk.Label(body, text=mod["name"], bg=C["surface"], fg=C["ink"],
                             font=("Segoe UI", 11, "bold"), anchor="w")
        name_lbl.pack(anchor="w", pady=(8, 2))

        sub_lbl = tk.Label(body, text=mod["sub"], bg=C["surface"], fg=C["ink_muted"],
                           font=("Segoe UI", 8), anchor="w",
                           wraplength=160, justify="left")
        sub_lbl.pack(anchor="w")

        def cmd(f=mod["frame"]): self.controller.show_frame(f)

        def _enter(e):
            outer.configure(bg=C["surface2"], highlightbackground=dot_color)
            top_line.configure(bg=dot_color)
            body.configure(bg=C["surface2"])
            icon_row.configure(bg=C["surface2"])
            name_lbl.configure(bg=C["surface2"])
            sub_lbl.configure(bg=C["surface2"])
            dot.configure(bg=C["surface2"])
            for w in icon_row.winfo_children():
                try: w.configure(bg=C["surface2"])
                except: pass

        def _leave(e):
            outer.configure(bg=C["surface"], highlightbackground=C["hair"])
            top_line.configure(bg=C["hair"])
            body.configure(bg=C["surface"])
            icon_row.configure(bg=C["surface"])
            name_lbl.configure(bg=C["surface"])
            sub_lbl.configure(bg=C["surface"])
            dot.configure(bg=C["surface"])
            for w in icon_row.winfo_children():
                try: w.configure(bg=C["surface"])
                except: pass

        for w in [outer, body, icon_row, name_lbl, sub_lbl, dot, top_line]:
            try:
                w.bind("<Button-1>", lambda _: cmd())
                w.bind("<Enter>",    _enter)
                w.bind("<Leave>",    _leave)
            except: pass


# ─── Rotinas Data Layer ─────────────────────────────────────────────────────

def _rotinas_data_path():
    return os.path.join(app_base_dir(), "rotinas_data.json")


class RotinasData:
    _instance = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self._rotinas    = []
        self._conclusoes = {}
        self._load()

    def _load(self):
        path = _rotinas_data_path()
        if not os.path.isfile(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = _json_mod.load(f)
            self._rotinas    = data.get("rotinas",    [])
            self._conclusoes = data.get("conclusoes", {})
            for r in self._rotinas:
                if "dias" not in r:
                    freq = r.get("frequencia", "diaria")
                    if freq == "diaria":
                        r["dias"] = [0, 1, 2, 3, 4, 5, 6]
                    elif freq == "dias_uteis":
                        r["dias"] = [0, 1, 2, 3, 4]
                    elif freq.startswith("semanal_"):
                        try:
                            r["dias"] = [int(freq.split("_")[1])]
                        except Exception:
                            r["dias"] = [0, 1, 2, 3, 4, 5, 6]
                    else:
                        r["dias"] = [0, 1, 2, 3, 4, 5, 6]
                if "alertas" not in r:
                    h = (r.get("hora_alerta") or "").strip()
                    r["alertas"] = [h] if h else []
        except Exception:
            pass

    def save(self):
        path = _rotinas_data_path()
        try:
            with open(path, "w", encoding="utf-8") as f:
                _json_mod.dump(
                    {"rotinas": self._rotinas, "conclusoes": self._conclusoes},
                    f, ensure_ascii=False, indent=2
                )
        except Exception:
            pass

    def rotinas(self):
        return list(self._rotinas)

    def add_rotina(self, nome, dias, alertas, cor, notas=""):
        r = {
            "id":      str(_uuid_mod.uuid4())[:8],
            "nome":    nome,
            "dias":    dias,
            "alertas": alertas,
            "cor":     cor,
            "notas":   notas,
            "ativa":   True,
        }
        self._rotinas.append(r)
        self.save()
        return r

    def update_rotina(self, rid, **kwargs):
        for r in self._rotinas:
            if r["id"] == rid:
                r.update(kwargs)
                self.save()
                return True
        return False

    def delete_rotina(self, rid):
        self._rotinas = [r for r in self._rotinas if r["id"] != rid]
        self.save()

    @staticmethod
    def _today_key():
        return date.today().isoformat()

    def is_done(self, rid, day=None):
        key = day or self._today_key()
        return self._conclusoes.get(key, {}).get(rid, False)

    def set_done(self, rid, done: bool, day=None):
        key = day or self._today_key()
        if key not in self._conclusoes:
            self._conclusoes[key] = {}
        self._conclusoes[key][rid] = done
        today = date.today()
        old_keys = [k for k in list(self._conclusoes)
                    if (today - date.fromisoformat(k)).days > 90]
        for k in old_keys:
            del self._conclusoes[k]
        self.save()

    def today_rotinas(self):
        today = date.today()
        wd    = today.weekday()
        result = []
        for r in self._rotinas:
            if not r.get("ativa", True):
                continue
            dias = r.get("dias")
            if dias is not None:
                if wd in dias:
                    result.append(r)
            else:
                freq = r.get("frequencia", "diaria")
                if freq == "diaria":
                    result.append(r)
                elif freq == "dias_uteis" and wd < 5:
                    result.append(r)
                elif freq.startswith("semanal_"):
                    try:
                        if wd == int(freq.split("_")[1]):
                            result.append(r)
                    except Exception:
                        pass
        return result

    @staticmethod
    def _rot_sort_key(rot):
        alertas = rot.get("alertas") or []
        if alertas:
            return alertas[0]
        return (rot.get("hora_alerta") or "99:99")

    def today_stats(self):
        rots  = self.today_rotinas()
        total = len(rots)
        done  = sum(1 for r in rots if self.is_done(r["id"]))
        return done, total


# ─── RotinasFrame ────────────────────────────────────────────────────────────

class RotinasFrame(tk.Frame):
    FREQ_LABELS = {
        "diaria":     "Diária",
        "dias_uteis": "Dias úteis (Seg–Sex)",
        "semanal_0":  "Semanal — Segunda",
        "semanal_1":  "Semanal — Terça",
        "semanal_2":  "Semanal — Quarta",
        "semanal_3":  "Semanal — Quinta",
        "semanal_4":  "Semanal — Sexta",
        "semanal_5":  "Semanal — Sábado",
        "semanal_6":  "Semanal — Domingo",
    }
    FREQ_OPTIONS = list(FREQ_LABELS.values())
    FREQ_KEYS    = list(FREQ_LABELS.keys())

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._data      = RotinasData.get()
        self._tab       = "hoje"
        self._build()
        self.after(500, self._start_alert_checker)

    def _build(self):
        hdr_wrap = tk.Frame(self, bg=C["bg"])
        hdr_wrap.pack(fill="x", padx=44, pady=(36, 0))
        eyebrow_label(hdr_wrap, "PLANEJAMENTO DO DIA").pack(anchor="w")
        tk.Label(hdr_wrap, text="Rotinas", bg=C["bg"], fg=C["ink"],
                 font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(6, 0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", pady=(20, 0))

        tab_row = tk.Frame(self, bg=C["bg"])
        tab_row.pack(fill="x", padx=44)
        self._tab_btns = {}
        for key, label in [("hoje", "Hoje"), ("gerenciar", "Gerenciar")]:
            btn = tk.Button(tab_row, text=label,
                            command=lambda k=key: self._switch_tab(k),
                            bg=C["bg"], fg=C["ink_muted"],
                            activebackground=C["bg"], activeforeground=C["ink"],
                            font=("Segoe UI", 9), relief="flat", bd=0,
                            padx=0, pady=10, cursor="hand2")
            btn.pack(side="left", padx=(0, 22))
            self._tab_btns[key] = btn

        make_hairline(self, bg=C["hair"]).pack(fill="x")

        self._content_area = tk.Frame(self, bg=C["bg"])
        self._content_area.pack(fill="both", expand=True)
        self._switch_tab("hoje")

    def _switch_tab(self, tab):
        self._tab = tab
        for k, btn in self._tab_btns.items():
            if k == tab:
                btn.configure(fg=C["accent"], font=("Segoe UI", 9, "bold"))
            else:
                btn.configure(fg=C["ink_muted"], font=("Segoe UI", 9))
        for w in self._content_area.winfo_children():
            w.destroy()
        if tab == "hoje":
            self._build_hoje(self._content_area)
        else:
            self._build_gerenciar(self._content_area)

    def on_show(self):
        self._switch_tab(self._tab)

    def _build_hoje(self, parent):
        sf    = ScrollableFrame(parent, bg=C["bg"])
        sf.pack(fill="both", expand=True)
        inner = sf.inner
        inner.configure(bg=C["bg"])

        now        = datetime.now()
        today_rots = self._data.today_rotinas()
        done_n, total = self._data.today_stats()

        hdr = tk.Frame(inner, bg=C["bg"])
        hdr.pack(fill="x", padx=44, pady=(28, 0))
        tk.Label(hdr, text=format_data_pt_br(now), bg=C["bg"],
                 fg=C["ink_muted"], font=("Segoe UI", 10)).pack(anchor="w")

        if total > 0:
            prog_text = f"{done_n} de {total} concluída{'s' if total > 1 else ''}"
            prog_color = C["ok"] if done_n == total else C["accent"]
        else:
            prog_text  = "Nenhuma rotina para hoje"
            prog_color = C["ink_faint"]

        tk.Label(hdr, text=prog_text, bg=C["bg"], fg=prog_color,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(6, 0))

        if total > 0:
            bar_bg = tk.Frame(hdr, bg=C["surface2"], height=4)
            bar_bg.pack(fill="x", pady=(8, 0))
            bar_bg.pack_propagate(False)
            bar_fg = tk.Frame(bar_bg, bg=prog_color, height=4)
            bar_fg.place(x=0, y=0, relheight=1.0, relwidth=min(done_n / total, 1.0))

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=44, pady=(20, 4))

        if not today_rots:
            empty = tk.Frame(inner, bg=C["bg"])
            empty.pack(fill="x", padx=44, pady=40)
            tk.Label(empty, text="◈", bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 24)).pack()
            tk.Label(empty, text="Nenhuma rotina configurada para hoje",
                     bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 11)).pack(pady=(8, 0))
            tk.Label(empty, text="Vá em Gerenciar para criar suas rotinas.",
                     bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 9)).pack(pady=(4, 0))
            styled_button(empty, "→ Gerenciar rotinas",
                          lambda: self._switch_tab("gerenciar"),
                          accent=True).pack(pady=(16, 0))
            return

        sorted_rots = sorted(today_rots, key=RotinasData._rot_sort_key)
        for rot in sorted_rots:
            self._make_hoje_row(inner, rot)
        tk.Frame(inner, bg=C["bg"], height=40).pack()
        sf.refresh_bindings()

    def _make_hoje_row(self, parent, rot):
        rid   = rot["id"]
        color = rot.get("cor", C["accent"])
        _als  = rot.get("alertas") or []
        hora  = _als[0] if _als else (rot.get("hora_alerta") or "")
        nome  = rot.get("nome", "Rotina")
        notas = (rot.get("notas") or "").strip()

        row = tk.Frame(parent, bg=C["bg"])
        row.pack(fill="x", padx=44, pady=2)
        inner_row = tk.Frame(row, bg=C["bg"], pady=9)
        inner_row.pack(fill="x")

        chk = tk.Canvas(inner_row, width=20, height=20,
                        bg=C["bg"], highlightthickness=0, bd=0, cursor="hand2")
        chk.pack(side="left", padx=(0, 12))

        if hora:
            tk.Label(inner_row, text=hora, bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 8), width=5, anchor="w").pack(side="left")

        dot = _make_dot(inner_row, color, 8, bg=C["bg"])
        dot.pack(side="left", padx=(0, 8))

        name_lbl = tk.Label(inner_row, text=nome, bg=C["bg"],
                            fg=C["ink"], font=("Segoe UI", 10), anchor="w")
        name_lbl.pack(side="left", fill="x", expand=True)

        note_lbl = None
        if notas:
            note_lbl = tk.Label(parent, text=f"      {notas}", bg=C["bg"],
                                fg=C["ink_faint"], font=("Segoe UI", 8),
                                anchor="w", wraplength=520, justify="left")
            note_lbl.pack(fill="x", padx=44, pady=(0, 2))

        make_hairline(parent, bg=C["surface2"]).pack(fill="x", padx=44)

        def _refresh_visual(done):
            self._draw_checkbox(chk, done, color)
            if done:
                name_lbl.configure(fg=C["ink_faint"],
                                   font=("Segoe UI", 10, "overstrike"))
            else:
                name_lbl.configure(fg=C["ink"],
                                   font=("Segoe UI", 10))

        _refresh_visual(self._data.is_done(rid))

        def toggle(e=None):
            new = not self._data.is_done(rid)
            self._data.set_done(rid, new)
            _refresh_visual(new)
            home = self.controller.frames.get("Home")
            if home and hasattr(home, "refresh_rotinas"):
                home.refresh_rotinas()

        for w in [row, inner_row, chk, name_lbl, dot]:
            try: w.bind("<Button-1>", toggle)
            except Exception: pass

    @staticmethod
    def _draw_checkbox(canvas, checked, color):
        canvas.delete("all")
        if checked:
            canvas.create_oval(1, 1, 19, 19, fill=color, outline="")
            canvas.create_text(10, 10, text="✓", fill=C["bg"],
                               font=("Segoe UI", 9, "bold"))
        else:
            canvas.create_oval(1, 1, 19, 19, fill="", outline=C["ink_faint"],
                               width=1.5)

    def _build_gerenciar(self, parent):
        act = tk.Frame(parent, bg=C["bg"])
        act.pack(fill="x", padx=44, pady=(18, 2))
        styled_button(act, "+ Nova rotina", self._open_nova_dialog,
                      accent=True, small=True).pack(side="right")

        sf = ScrollableFrame(parent, bg=C["bg"])
        sf.pack(fill="both", expand=True)
        self._ger_inner = sf.inner
        self._ger_sf    = sf
        self._ger_inner.configure(bg=C["bg"])
        self._refresh_gerenciar()

    def _refresh_gerenciar(self):
        inner = getattr(self, "_ger_inner", None)
        if inner is None:
            return
        for w in inner.winfo_children():
            w.destroy()
        rots = self._data.rotinas()
        if not rots:
            e = tk.Frame(inner, bg=C["bg"])
            e.pack(fill="x", padx=44, pady=60)
            tk.Label(e, text="◈", bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 24)).pack()
            tk.Label(e, text="Nenhuma rotina cadastrada", bg=C["bg"],
                     fg=C["ink_faint"], font=("Segoe UI", 11)).pack(pady=(8, 0))
            styled_button(e, "+ Nova rotina", self._open_nova_dialog,
                          accent=True).pack(pady=(16, 0))
        else:
            for rot in rots:
                self._make_ger_row(inner, rot)
            tk.Frame(inner, bg=C["bg"], height=40).pack()
        sf = getattr(self, "_ger_sf", None)
        if sf:
            sf.refresh_bindings()

    @staticmethod
    def _dias_label(dias):
        if dias is None:
            return "Diária"
        s = sorted(dias)
        if s == list(range(7)):
            return "Diária"
        if s == list(range(5)):
            return "Dias úteis"
        abbr = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
        return "  ".join(abbr[d] for d in s)

    def _make_ger_row(self, parent, rot):
        color    = rot.get("cor", C["accent"])
        nome     = rot.get("nome", "Rotina")
        dias     = rot.get("dias")
        alertas  = rot.get("alertas") or []
        if not alertas:
            h = (rot.get("hora_alerta") or "").strip()
            if h: alertas = [h]
        notas    = (rot.get("notas") or "").strip()
        dias_lbl = self._dias_label(dias)
        horas_lbl = "  ·  " + ", ".join(alertas) if alertas else ""

        row_o = tk.Frame(parent, bg=C["bg"])
        row_o.pack(fill="x")
        row_i = tk.Frame(row_o, bg=C["bg"], padx=44, pady=10)
        row_i.pack(fill="x")

        dot_col = tk.Frame(row_i, bg=C["bg"], width=20)
        dot_col.pack(side="left", fill="y")
        dot_col.pack_propagate(False)
        _make_dot(dot_col, color, 10, bg=C["bg"]).pack(anchor="n", pady=(5, 0))

        center = tk.Frame(row_i, bg=C["bg"])
        center.pack(side="left", fill="both", expand=True, padx=(10, 16))
        tk.Label(center, text=nome, bg=C["bg"], fg=C["ink"],
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w")
        meta = dias_lbl + horas_lbl
        tk.Label(center, text=meta, bg=C["bg"], fg=C["ink_muted"],
                 font=("Segoe UI", 8), anchor="w").pack(anchor="w", pady=(2, 0))
        if notas:
            tk.Label(center, text=notas, bg=C["bg"], fg=C["ink_faint"],
                     font=("Segoe UI", 8), anchor="w",
                     wraplength=380, justify="left").pack(anchor="w", pady=(1, 0))

        right = tk.Frame(row_i, bg=C["bg"])
        right.pack(side="right", fill="y", pady=(2, 0))
        styled_button(right, "Editar",
                      lambda r=rot: self._open_edit_dialog(r),
                      small=True).pack(side="left")
        styled_button(right, "✕",
                      lambda r=rot: self._del_rotina(r),
                      danger=True, small=True).pack(side="left", padx=(4, 0))

        make_hairline(parent, bg=C["hair"]).pack(fill="x", padx=44)

        def _bg_all(bg):
            for w in [row_o, row_i, dot_col, center, right]:
                try: w.configure(bg=bg)
                except Exception: pass
            for w in center.winfo_children() + right.winfo_children() + \
                     dot_col.winfo_children():
                try: w.configure(bg=bg)
                except Exception: pass
        row_o.bind("<Enter>", lambda _: _bg_all(C["surface"]))
        row_o.bind("<Leave>", lambda _: _bg_all(C["bg"]))
        row_i.bind("<Enter>", lambda _: _bg_all(C["surface"]))
        row_i.bind("<Leave>", lambda _: _bg_all(C["bg"]))

    def _del_rotina(self, rot):
        if messagebox.askyesno("Remover", f"Remover '{rot['nome']}'?", parent=self):
            self._data.delete_rotina(rot["id"])
            self._refresh_gerenciar()
            home = self.controller.frames.get("Home")
            if home and hasattr(home, "refresh_rotinas"):
                home.refresh_rotinas()

    def _open_nova_dialog(self):
        self._open_rotina_dialog(None)

    def _open_edit_dialog(self, rot):
        self._open_rotina_dialog(rot)

    def _open_rotina_dialog(self, rot):
        editing = rot is not None
        dlg = tk.Toplevel(self)
        dlg.title("Editar Rotina" if editing else "Nova Rotina")
        dlg.configure(bg=C["surface"])
        dlg.geometry("480x580")
        dlg.resizable(False, True)
        dlg.grab_set()

        hdr = tk.Frame(dlg, bg=C["surface"], padx=24)
        hdr.pack(fill="x", pady=(22, 0))
        color_var = tk.StringVar(value=(rot or {}).get("cor", C["accent"]))

        title_row = tk.Frame(hdr, bg=C["surface"])
        title_row.pack(fill="x")
        title_dot = _make_dot(title_row, color_var.get(), 12, bg=C["surface"])
        title_dot.pack(side="left", pady=(3, 0))
        tk.Label(title_row,
                 text="Editar Rotina" if editing else "Nova Rotina",
                 bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=(8, 0))
        make_hairline(dlg, bg=C["hair"]).pack(fill="x", pady=(18, 0))

        sf   = ScrollableFrame(dlg, bg=C["surface"])
        sf.pack(fill="both", expand=True)
        sf.link_wheel(dlg)
        form = sf.inner
        form.configure(bg=C["surface"])
        pad  = tk.Frame(form, bg=C["surface"])
        pad.pack(fill="both", expand=True, padx=24)

        tk.Label(pad, text="NOME", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(16, 0))
        nome_var = tk.StringVar(value=(rot or {}).get("nome", ""))
        styled_entry(pad, textvariable=nome_var).pack(fill="x", pady=(4, 0))

        tk.Label(pad, text="DIAS DA SEMANA", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(16, 0))

        rot_dias = (rot or {}).get("dias")
        if rot_dias is None:
            freq0 = (rot or {}).get("frequencia", "diaria")
            if freq0 == "dias_uteis":
                rot_dias = [0, 1, 2, 3, 4]
            elif freq0.startswith("semanal_"):
                try:
                    rot_dias = [int(freq0.split("_")[1])]
                except Exception:
                    rot_dias = [0, 1, 2, 3, 4, 5, 6]
            else:
                rot_dias = [0, 1, 2, 3, 4, 5, 6]

        dias_sel  = set(rot_dias)
        DIA_ABBR  = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
        dias_btns = {}

        dias_row = tk.Frame(pad, bg=C["surface"])
        dias_row.pack(anchor="w", pady=(6, 0))

        def _render_dias_btns():
            for d, btn in dias_btns.items():
                sel = d in dias_sel
                btn.configure(
                    bg=C["accent"]   if sel else C["surface3"],
                    fg=C["bg"]       if sel else C["ink_muted"],
                    font=("Segoe UI", 8, "bold") if sel else ("Segoe UI", 8),
                )

        def _toggle_dia(d):
            if d in dias_sel:
                if len(dias_sel) > 1:
                    dias_sel.discard(d)
            else:
                dias_sel.add(d)
            _render_dias_btns()

        for d, lbl in enumerate(DIA_ABBR):
            sel = d in dias_sel
            btn = tk.Button(dias_row, text=lbl,
                            command=lambda dd=d: _toggle_dia(dd),
                            bg=C["accent"] if sel else C["surface3"],
                            fg=C["bg"] if sel else C["ink_muted"],
                            font=("Segoe UI", 8, "bold") if sel else ("Segoe UI", 8),
                            relief="flat", bd=0, padx=9, pady=5,
                            cursor="hand2")
            btn.pack(side="left", padx=(0, 3))
            dias_btns[d] = btn

        preset_row = tk.Frame(pad, bg=C["surface"])
        preset_row.pack(anchor="w", pady=(5, 0))

        def _preset(days):
            dias_sel.clear()
            dias_sel.update(days)
            _render_dias_btns()

        for ptxt, pdays in [("Todos os dias", range(7)), ("Dias úteis", range(5))]:
            tk.Button(preset_row, text=ptxt,
                      command=lambda d=pdays: _preset(d),
                      bg=C["surface3"], fg=C["ink_faint"],
                      font=("Segoe UI", 7), relief="flat", bd=0,
                      padx=6, pady=3, cursor="hand2").pack(side="left", padx=(0, 5))

        tk.Label(pad, text="ALERTAS", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(16, 0))

        rot_als = list((rot or {}).get("alertas") or [])
        if not rot_als:
            h0 = (rot or {}).get("hora_alerta") or ""
            if h0:
                rot_als = [h0]
        alertas_list = list(rot_als)

        alertas_frame = tk.Frame(pad, bg=C["surface"])
        alertas_frame.pack(fill="x", pady=(4, 0))

        def _render_alertas():
            for w in alertas_frame.winfo_children():
                w.destroy()
            if not alertas_list:
                tk.Label(alertas_frame, text="Nenhum alerta configurado",
                         bg=C["surface"], fg=C["ink_faint"],
                         font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
                return
            for i, hora in enumerate(alertas_list):
                r = tk.Frame(alertas_frame, bg=C["surface"])
                r.pack(anchor="w", pady=2)
                tk.Label(r, text=hora,
                         bg=C["surface2"], fg=C["ink"],
                         font=("Segoe UI", 9, "bold"),
                         padx=10, pady=3).pack(side="left")
                tk.Button(r, text="✕",
                          command=lambda idx=i: (_del_alerta(idx)),
                          bg=C["surface"], fg=C["ink_faint"],
                          activebackground=C["surface"], activeforeground=C["err"],
                          font=("Segoe UI", 8), relief="flat", bd=0,
                          padx=4, cursor="hand2").pack(side="left", padx=(3, 0))

        def _del_alerta(idx):
            if 0 <= idx < len(alertas_list):
                alertas_list.pop(idx)
                _render_alertas()

        _render_alertas()

        make_hairline(pad, bg=C["surface3"]).pack(fill="x", pady=(10, 0))
        add_lbl = tk.Frame(pad, bg=C["surface"])
        add_lbl.pack(fill="x", pady=(8, 0))
        tk.Label(add_lbl, text="Nova hora", bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI", 8)).pack(side="left")

        add_row = tk.Frame(pad, bg=C["surface"])
        add_row.pack(fill="x", pady=(4, 0))
        add_hora_var = tk.StringVar()
        styled_entry(add_row, textvariable=add_hora_var, width=8).pack(side="left")
        tk.Label(add_row, text="HH:MM", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7)).pack(side="left", padx=(5, 12))

        def _add_hora():
            h = add_hora_var.get().strip()
            if not re.match(r"^\d{2}:\d{2}$", h):
                return
            if h not in alertas_list:
                alertas_list.append(h)
                alertas_list.sort()
            add_hora_var.set("")
            _render_alertas()

        tk.Button(add_row, text="+ Adicionar", command=_add_hora,
                  bg=C["surface3"], fg=C["ink_muted"],
                  activebackground=C["accent_dim"], activeforeground=C["ink"],
                  font=("Segoe UI", 8), relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2").pack(side="left")

        antes_row = tk.Frame(pad, bg=C["surface"])
        antes_row.pack(fill="x", pady=(8, 0))
        antes_var = tk.StringVar(value="15")
        tk.Label(antes_row, text="ou", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(0, 6))
        styled_entry(antes_row, textvariable=antes_var, width=4).pack(side="left")
        tk.Label(antes_row, text="min antes de", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(5, 5))
        ref_var = tk.StringVar()
        styled_entry(antes_row, textvariable=ref_var, width=8).pack(side="left")
        tk.Label(antes_row, text="HH:MM", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7)).pack(side="left", padx=(4, 8))

        def _add_antes():
            try:
                mins = int(antes_var.get().strip())
                ref  = ref_var.get().strip()
                if not re.match(r"^\d{2}:\d{2}$", ref):
                    return
                h, m  = int(ref[:2]), int(ref[3:])
                total = (h * 60 + m - mins) % (24 * 60)
                nova  = f"{total // 60:02d}:{total % 60:02d}"
                if nova not in alertas_list:
                    alertas_list.append(nova)
                    alertas_list.sort()
                ref_var.set("")
                _render_alertas()
            except Exception:
                pass

        tk.Button(antes_row, text="OK", command=_add_antes,
                  bg=C["surface3"], fg=C["ink_muted"],
                  activebackground=C["accent_dim"], activeforeground=C["ink"],
                  font=("Segoe UI", 8), relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2").pack(side="left")

        make_hairline(pad, bg=C["surface3"]).pack(fill="x", pady=(10, 0))

        tk.Label(pad, text="COR", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(14, 0))
        color_row  = tk.Frame(pad, bg=C["surface"])
        color_row.pack(anchor="w", pady=(6, 0))
        color_btns = {}

        def _pick(c_val):
            color_var.set(c_val)
            title_dot.itemconfig(1, fill=c_val)
            for cv2, cb2 in color_btns.items():
                r2 = 2 if cv2 == c_val else 0
                cb2.configure(highlightthickness=r2,
                              highlightbackground=C["ink"] if cv2 == c_val else C["hair"])

        for dc in DOT_COLORS:
            cv = tk.Canvas(color_row, width=20, height=20, bg=C["surface"],
                           highlightthickness=2 if dc == color_var.get() else 0,
                           highlightbackground=C["ink"] if dc == color_var.get() else C["hair"],
                           cursor="hand2")
            cv.pack(side="left", padx=3)
            cv.create_oval(3, 3, 17, 17, fill=dc, outline="")
            cv.bind("<Button-1>", lambda _, d=dc: _pick(d))
            color_btns[dc] = cv

        tk.Label(pad, text="NOTAS  (opcional)", bg=C["surface"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(14, 0))
        notas_var = tk.StringVar(value=(rot or {}).get("notas", "") or "")
        styled_entry(pad, textvariable=notas_var).pack(fill="x", pady=(4, 12))

        make_hairline(dlg, bg=C["hair"]).pack(fill="x")
        foot = tk.Frame(dlg, bg=C["surface"], padx=24, pady=12)
        foot.pack(fill="x")

        def salvar():
            nome = nome_var.get().strip()
            if not nome:
                messagebox.showwarning("Campo obrigatório", "Informe um nome.",
                                       parent=dlg)
                return
            dias_final = sorted(dias_sel)
            als_final  = alertas_list[:]

            if editing:
                self._data.update_rotina(
                    rot["id"], nome=nome, dias=dias_final,
                    alertas=als_final, cor=color_var.get(),
                    notas=notas_var.get().strip()
                )
            else:
                self._data.add_rotina(nome, dias_final, als_final,
                                      color_var.get(), notas_var.get().strip())
            dlg.destroy()
            self._refresh_gerenciar()
            home = self.controller.frames.get("Home")
            if home and hasattr(home, "refresh_rotinas"):
                home.refresh_rotinas()

        if editing:
            def _excluir():
                if messagebox.askyesno("Excluir", f"Remover '{rot['nome']}'?",
                                       parent=dlg):
                    self._data.delete_rotina(rot["id"])
                    dlg.destroy()
                    self._refresh_gerenciar()
                    home = self.controller.frames.get("Home")
                    if home and hasattr(home, "refresh_rotinas"):
                        home.refresh_rotinas()
            styled_button(foot, "Excluir", _excluir,
                          danger=True, small=True).pack(side="left")

        styled_button(foot, "Cancelar", dlg.destroy, small=True).pack(
            side="right", padx=(6, 0))
        styled_button(foot, "Salvar", salvar, accent=True, small=True).pack(side="right")

    def _start_alert_checker(self):
        self._alerted_flags: set = set()
        self._alerted_date:  str = ""
        self._check_alerts()

    def _check_alerts(self):
        try:
            if not self.winfo_exists():
                return
        except Exception:
            return

        now       = datetime.now()
        now_str   = now.strftime("%H:%M")
        today_str = now.strftime("%Y-%m-%d")

        if self._alerted_date != today_str:
            self._alerted_flags = set()
            self._alerted_date  = today_str

        for rot in self._data.today_rotinas():
            if self._data.is_done(rot["id"]):
                continue
            alertas = rot.get("alertas") or []
            if not alertas:
                h = (rot.get("hora_alerta") or "").strip()
                if h:
                    alertas = [h]
            for hora in alertas:
                if hora.strip() != now_str:
                    continue
                flag = (rot["id"], hora)
                if flag in self._alerted_flags:
                    continue
                self._alerted_flags.add(flag)
                nome = rot["nome"]
                self.after(0, lambda n=nome, h=hora: self._show_alert(n, h))

        self.after(30_000, self._check_alerts)

    def _show_alert(self, nome, hora):
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            pass

        try:
            if not self.controller.winfo_exists():
                return
        except Exception:
            return

        dlg = tk.Toplevel(self.controller)
        dlg.title("Lembrete")
        dlg.configure(bg=C["surface"])
        dlg.resizable(False, False)
        dlg.overrideredirect(True)
        dlg.attributes("-topmost", True)
        dlg.lift()
        dlg.focus_force()

        if sys.platform == "win32":
            try:
                import ctypes as _ctypes
                _hwnd = _ctypes.windll.user32.GetParent(dlg.winfo_id())
                if not _hwnd:
                    _hwnd = dlg.winfo_id()
                _ctypes.windll.user32.SetForegroundWindow(_hwnd)
                _ctypes.windll.user32.BringWindowToTop(_hwnd)
                _ctypes.windll.user32.FlashWindow(_hwnd, True)
            except Exception:
                pass

        tk.Frame(dlg, bg=C["accent"], height=3).pack(fill="x")

        body = tk.Frame(dlg, bg=C["surface"], padx=28, pady=20)
        body.pack(fill="both", expand=True)

        hdr_row = tk.Frame(body, bg=C["surface"])
        hdr_row.pack(fill="x")
        tk.Label(hdr_row, text="\U0001f514", bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 18)).pack(side="left")
        tk.Label(hdr_row, text=hora, bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 22, "bold")).pack(side="left", padx=(10, 0))

        make_hairline(body, bg=C["hair"]).pack(fill="x", pady=(14, 12))

        tk.Label(body, text=nome, bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 13, "bold"),
                 wraplength=300, justify="left").pack(anchor="w")
        tk.Label(body, text="Rotina programada para agora", bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 0))

        make_hairline(body, bg=C["hair"]).pack(fill="x", pady=(16, 0))

        foot = tk.Frame(body, bg=C["surface"])
        foot.pack(fill="x", pady=(12, 0))

        def _dismiss():
            dlg.destroy()

        styled_button(foot, "Dispensar", _dismiss, small=True).pack(side="right")
        styled_button(foot, "Ok, entendido", _dismiss, accent=True, small=True).pack(
            side="right", padx=(0, 6))

        dlg.update_idletasks()
        sw = dlg.winfo_screenwidth()
        sh = dlg.winfo_screenheight()
        w  = dlg.winfo_reqwidth()
        h  = dlg.winfo_reqheight()
        dlg.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

        dlg.bind("<Escape>", lambda _e: _dismiss())
        dlg.grab_set()


class BPMConfigFrame(tk.Frame):
    CLIENTS = list(BPM_CLIENT_DATA.keys())

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._selected  = {}
        self._func_var  = tk.StringVar()
        self._senha_var = tk.StringVar()
        self._mode_var  = tk.StringVar(value="invertido")
        self._custom_entries = []
        self._build()

    def _only_digits(self, s):
        return "".join(c for c in (s or "") if c.isdigit())

    def on_show(self):
        self._func_var.set(getattr(self.controller, "bpm_funcional", "") or "")
        self._senha_var.set(getattr(self.controller, "bpm_password", "") or "")
        if hasattr(self, "_sf"):
            self._sf.refresh_bindings()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        tk.Label(hdr, text="Configurar BPM", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left")
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(14,0))

        # ── Mode selector ──
        mode_card = card_frame(self)
        mode_card.pack(fill="x", padx=32, pady=(20, 0))
        mode_body = tk.Frame(mode_card, bg=C["surface"], padx=18, pady=14)
        mode_body.pack(fill="x")

        tk.Label(mode_body, text="Tipo de Operação", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(mode_body, text="Selecione o modo de abertura BPM.",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 10))

        mode_row = tk.Frame(mode_body, bg=C["surface"])
        mode_row.pack(fill="x")

        for val, label, desc in [
            ("invertido", "BPM Invertido", "Clientes mapeados com dados pré-configurados"),
            ("nova_plataforma", "BPM Nova Plataforma", "Informe CNPJ, Agência, Conta, Plataforma e Valor manualmente"),
        ]:
            rf = tk.Frame(mode_row, bg=C["surface"], pady=4)
            rf.pack(fill="x")
            rb = tk.Radiobutton(rf, text=label, variable=self._mode_var, value=val,
                                bg=C["surface"], fg=C["ink"],
                                selectcolor=C["surface2"],
                                activebackground=C["surface"],
                                activeforeground=C["ink"],
                                font=("Segoe UI", 9, "bold"),
                                command=self._on_mode_change,
                                cursor="hand2")
            rb.pack(side="left")
            tk.Label(rf, text=f"  — {desc}", bg=C["surface"], fg=C["ink_faint"],
                     font=("Segoe UI", 8)).pack(side="left")

        # ── Credenciais ──
        self._cred_sec = card_frame(self)
        self._cred_sec.pack(fill="x", padx=32, pady=(16, 0))
        cred_sec = self._cred_sec
        tk.Label(cred_sec, text="Credenciais do Painel de Serviços", bg=C["surface"],
                 fg=C["ink"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        tk.Label(cred_sec, text="Funcional e senha para login no Painel BPM",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=(2,10))
        make_hairline(cred_sec, bg=C["hair"]).pack(fill="x")

        cred_form = tk.Frame(cred_sec, bg=C["surface"], padx=18, pady=14)
        cred_form.pack(fill="x")
        cred_form.columnconfigure(0, weight=1)
        cred_form.columnconfigure(1, weight=1)

        fc = tk.Frame(cred_form, bg=C["surface"])
        fc.grid(row=0, column=0, sticky="ew", padx=(0,8))
        tk.Label(fc, text="Funcional (somente números)", bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w")
        self._ent_func = styled_entry(fc, textvariable=self._func_var)
        self._ent_func.pack(fill="x", pady=(4,0))

        sc2 = tk.Frame(cred_form, bg=C["surface"])
        sc2.grid(row=0, column=1, sticky="ew", padx=(8,0))
        tk.Label(sc2, text="Senha (até 6 dígitos)", bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w")
        self._ent_senha = styled_entry(sc2, textvariable=self._senha_var, show="•")
        self._ent_senha.pack(fill="x", pady=(4,0))

        def _key_func(e):
            if e.keysym in ("BackSpace","Delete","Tab","Return","Left","Right","Home","End"): return
            if e.state & 0x4: return
            if e.char and e.char.isdigit(): return
            if e.char: return "break"

        def _key_senha(e):
            if e.keysym in ("BackSpace","Delete","Tab","Return","Left","Right","Home","End"): return
            if e.state & 0x4: return
            if not e.char: return
            if not e.char.isdigit(): return "break"
            if len(self._ent_senha.get()) >= 6 and not self._ent_senha.selection_present():
                return "break"

        self._ent_func.bind("<KeyPress>", _key_func)
        self._ent_senha.bind("<KeyPress>", _key_senha)

        # ── Container dinâmico ──
        self._dynamic_area = tk.Frame(self, bg=C["bg"])
        self._dynamic_area.pack(fill="both", expand=True)

        # ── Invertido panel ──
        self._invertido_panel = tk.Frame(self._dynamic_area, bg=C["bg"])
        self._nova_panel = tk.Frame(self._dynamic_area, bg=C["bg"])

        self._build_invertido_panel()
        self._build_nova_panel()

        # ── Footer ──
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(10, 0))

        foot = tk.Frame(self, bg=C["bg"])
        foot.pack(fill="x", padx=32, pady=16)
        styled_button(foot, "← Voltar",
                      lambda: self.controller.show_frame("Home")).pack(side="left")
        self._run_btn = styled_button(foot, "▶  Iniciar BPM",
                                      self._start_bpm, accent=True)
        self._run_btn.pack(side="right")

        self._on_mode_change()

    def _on_mode_change(self):
        self._invertido_panel.pack_forget()
        self._nova_panel.pack_forget()
        if self._mode_var.get() == "invertido":
            self._invertido_panel.pack(fill="both", expand=True, padx=32)
        else:
            self._nova_panel.pack(fill="both", expand=True, padx=32)

    def _build_invertido_panel(self):
        sec2 = card_frame(self._invertido_panel)
        sec2.pack(fill="x", pady=(0, 0))
        tk.Label(sec2, text="Clientes e Valores", bg=C["surface"],
                 fg=C["ink"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        tk.Label(sec2, text="Selecione os clientes e informe o valor da operação",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=(2,10))
        make_hairline(sec2, bg=C["hair"]).pack(fill="x")

        clients_frame = tk.Frame(sec2, bg=C["surface"], padx=18, pady=12)
        clients_frame.pack(fill="x")

        self._client_rows = {}
        for cli in self.CLIENTS:
            self._make_client_row(clients_frame, cli)

    def _make_client_row(self, parent, cli):
        row = tk.Frame(parent, bg=C["surface"], pady=3)
        row.pack(fill="x")
        row.columnconfigure(1, weight=1)

        selected = tk.BooleanVar(value=False)
        val_var  = tk.StringVar(value="R$ 0,00")
        val_digits = [""]

        cb = tk.Checkbutton(row, variable=selected, bg=C["surface"],
                            selectcolor=C["bg"], activebackground=C["surface"],
                            fg=C["ink_muted"], font=("Segoe UI",9))
        cb.grid(row=0, column=0, sticky="w")

        name_lbl = tk.Label(row, text=cli, bg=C["surface"], fg=C["ink_muted"],
                            font=("Segoe UI",9))
        name_lbl.grid(row=0, column=1, sticky="w", padx=(4,8))

        ent = styled_entry(row, textvariable=val_var, width=14)
        ent.grid(row=0, column=2, sticky="e")
        ent.configure(state="disabled", disabledbackground=C["bg"], disabledforeground=C["ink_faint"])

        def fmt_val():
            if not val_digits[0]: val_var.set("R$ 0,00"); return
            d = Decimal(int(val_digits[0])) / Decimal("100")
            val_var.set(_fmt_brl(d))

        def on_key(e):
            if e.keysym == "BackSpace": val_digits[0] = val_digits[0][:-1]; fmt_val(); return "break"
            if e.char and e.char.isdigit(): val_digits[0] += e.char; fmt_val(); return "break"
            if e.keysym in {"Tab","Left","Right","Home","End"}: return
            return "break"

        def on_paste(_):
            try: clip = row.clipboard_get()
            except: return "break"
            d = _parse_brl(clip)
            if d is None: return "break"
            val_digits[0] = str(max(int((d*100).quantize(Decimal("1"))),0))
            fmt_val(); return "break"

        def on_toggle(*_):
            if selected.get():
                ent.configure(state="normal", bg=C["bg"], fg=C["ink"])
                self._selected[cli] = val_var
            else:
                ent.configure(state="disabled", disabledbackground=C["bg"], disabledforeground=C["ink_faint"])
                self._selected.pop(cli, None)

        cb.configure(command=on_toggle)
        ent.bind("<KeyPress>", on_key)
        ent.bind("<<Paste>>", on_paste)
        ent.bind("<Control-v>", on_paste)
        self._client_rows[cli] = {"selected": selected, "val_var": val_var, "val_digits": val_digits, "entry": ent}

    def _build_nova_panel(self):
        sec = card_frame(self._nova_panel)
        sec.pack(fill="x", pady=(0, 0))
        tk.Label(sec, text="Dados da Nova Plataforma", bg=C["surface"],
                 fg=C["ink"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        tk.Label(sec, text="Adicione entradas com CNPJ, Agência, Conta, Plataforma e Valor.",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=(2,10))
        make_hairline(sec, bg=C["hair"]).pack(fill="x")

        self._nova_list_frame = tk.Frame(sec, bg=C["surface"], padx=18, pady=12)
        self._nova_list_frame.pack(fill="x")

        add_btn_frame = tk.Frame(sec, bg=C["surface"], padx=18, pady=(0, 12))
        add_btn_frame.pack(fill="x")
        styled_button(add_btn_frame, "+ Adicionar entrada", self._add_nova_entry,
                      accent=True, small=True).pack(side="left")

        self._custom_entries = []
        self._add_nova_entry()

    def _add_nova_entry(self):
        idx = len(self._custom_entries)
        entry_data = {
            "ref": tk.StringVar(value=f"Cliente {idx + 1}"),
            "cnpj": tk.StringVar(),
            "ag": tk.StringVar(),
            "conta": tk.StringVar(),
            "plataforma": tk.StringVar(),
            "valor_digits": [""],
            "valor_var": tk.StringVar(value="R$ 0,00"),
        }
        self._custom_entries.append(entry_data)

        item = tk.Frame(self._nova_list_frame, bg=C["surface2"],
                        highlightthickness=1, highlightbackground=C["hair"])
        item.pack(fill="x", pady=(0, 8))

        head = tk.Frame(item, bg=C["surface2"], padx=12, pady=8)
        head.pack(fill="x")
        tk.Label(head, text=f"Entrada {idx + 1}", bg=C["surface2"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(side="left")
        styled_button(head, "✕", lambda e=entry_data, i=item: self._remove_nova_entry(e, i),
                      danger=True, small=True).pack(side="right")

        body = tk.Frame(item, bg=C["surface2"], padx=12, pady=10)
        body.pack(fill="x")
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.columnconfigure(2, weight=1)

        fields = [
            ("Referência", entry_data["ref"], 0, 0),
            ("CNPJ (somente números)", entry_data["cnpj"], 0, 1),
            ("Plataforma", entry_data["plataforma"], 1, 0),
            ("Agência", entry_data["ag"], 1, 1),
            ("Conta Corrente", entry_data["conta"], 2, 0),
        ]
        for label, var, r, c in fields:
            f = tk.Frame(body, bg=C["surface2"])
            f.grid(row=r, column=c, sticky="ew", padx=(0 if c == 0 else 6, 0), pady=3)
            tk.Label(f, text=label, bg=C["surface2"], fg=C["ink_faint"],
                     font=("Segoe UI", 7)).pack(anchor="w")
            styled_entry(f, textvariable=var, width=18).pack(fill="x", pady=(2, 0))

        # Valor field
        vf = tk.Frame(body, bg=C["surface2"])
        vf.grid(row=2, column=1, sticky="ew", padx=(6, 0), pady=3)
        tk.Label(vf, text="Valor da Operação", bg=C["surface2"], fg=C["ink_faint"],
                 font=("Segoe UI", 7)).pack(anchor="w")
        val_ent = styled_entry(vf, textvariable=entry_data["valor_var"], width=18)
        val_ent.pack(fill="x", pady=(2, 0))

        def fmt_val(e=entry_data):
            if not e["valor_digits"][0]: e["valor_var"].set("R$ 0,00"); return
            d = Decimal(int(e["valor_digits"][0])) / Decimal("100")
            e["valor_var"].set(_fmt_brl(d))

        def on_key(e_arg, e=entry_data):
            if e_arg.keysym == "BackSpace": e["valor_digits"][0] = e["valor_digits"][0][:-1]; fmt_val(e); return "break"
            if e_arg.char and e_arg.char.isdigit(): e["valor_digits"][0] += e_arg.char; fmt_val(e); return "break"
            if e_arg.keysym in {"Tab","Left","Right","Home","End"}: return
            return "break"

        val_ent.bind("<KeyPress>", on_key)

    def _remove_nova_entry(self, entry_data, item_frame):
        if entry_data in self._custom_entries:
            self._custom_entries.remove(entry_data)
        item_frame.destroy()

    def _start_bpm(self):
        func = self._only_digits(self._func_var.get())
        senha = self._only_digits(self._senha_var.get())
        if not func:
            messagebox.showwarning("Funcional obrigatório", "Informe o funcional (somente números)."); return
        if not (1 <= len(senha) <= 6):
            messagebox.showwarning("Senha inválida", "Senha deve ter 1 a 6 dígitos."); return

        mode = self._mode_var.get()
        selection = []

        if mode == "invertido":
            for cli in self.CLIENTS:
                row = self._client_rows[cli]
                if not row["selected"].get(): continue
                raw = row["val_var"].get()
                d = _parse_brl(raw)
                if d is None or d == Decimal("0.00"):
                    messagebox.showwarning("Valor inválido", f"Informe um valor válido para {cli}."); return
                selection.append({"cliente": cli, "valor": _fmt_brl_from_raw(raw), "mode": "invertido"})
        else:
            if not self._custom_entries:
                messagebox.showwarning("Nenhuma entrada", "Adicione ao menos uma entrada para BPM Nova Plataforma."); return
            for e in self._custom_entries:
                ref = e["ref"].get().strip()
                cnpj = self._only_digits(e["cnpj"].get())
                ag = e["ag"].get().strip()
                conta = e["conta"].get().strip()
                plat = e["plataforma"].get().strip()
                if not cnpj or not ag or not conta or not plat:
                    messagebox.showwarning("Dados incompletos", f"Preencha CNPJ, Agência, Conta e Plataforma para '{ref or 'entrada'}'."); return
                raw = e["valor_var"].get()
                d = _parse_brl(raw)
                if d is None or d == Decimal("0.00"):
                    messagebox.showwarning("Valor inválido", f"Informe um valor válido para '{ref or 'entrada'}'."); return
                selection.append({
                    "cliente": ref or cnpj,
                    "valor": _fmt_brl_from_raw(raw),
                    "mode": "nova_plataforma",
                    "custom_cnpj": cnpj,
                    "custom_ag": ag,
                    "custom_conta": conta,
                    "custom_plataforma": plat,
                })

        if not selection:
            messagebox.showwarning("Seleção vazia", "Selecione ao menos um cliente/entrada."); return

        self.controller.bpm_funcional    = func
        self.controller.bpm_password     = senha
        self.controller.bpm_run_selection = selection
        self.controller.show_frame("BPM")


class ShareFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self.vars = {k: tk.StringVar() for k in [
            "razao_social","cnpj","conta","plataforma","regiao",
            "trader_espec","valor_operacao","spread","prazo_min","prazo_max","modalidade"
        ]}
        self.hidden = {"premio":"sem prêmio","liquidacao":"Débito em CC","cnpj8":""}
        self.vars["modalidade"].trace_add("write", lambda *_: self._update_resumo())
        self.vars["regiao"].trace_add("write",    lambda *_: self._update_trader_espec())
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        tk.Label(hdr, text="Cadastro Share", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left")
        styled_button(hdr, "📄  Abrir PDF…", self._on_open_pdf, accent=True).pack(side="right")
        styled_button(hdr, "← Voltar",
                      lambda: self.controller.show_frame("Home")).pack(side="right", padx=(0,6))
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(14,0))

        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._sf.link_wheel(self)
        body = self._sf.inner
        body.configure(bg=C["bg"])

        fields_card = card_frame(body)
        fields_card.pack(fill="x", padx=32, pady=(20,0))
        tk.Label(fields_card, text="Campos extraídos", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        tk.Label(fields_card, text="Edite se necessário antes de gerar o resumo.",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=(2,10))
        make_hairline(fields_card, bg=C["hair"]).pack(fill="x")

        grid = tk.Frame(fields_card, bg=C["surface"], padx=18, pady=14)
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        def field(key, label, row, col, combo=False, width=28, readonly=False):
            b = tk.Frame(grid, bg=C["surface"])
            b.grid(row=row, column=col, sticky="ew", padx=(0 if col==0 else 10, 0), pady=6)
            tk.Label(b, text=label, bg=C["surface"], fg=C["ink_muted"],
                     font=("Segoe UI",8)).pack(anchor="w", pady=(0,4))
            row_f = tk.Frame(b, bg=C["surface"])
            row_f.pack(fill="x")
            row_f.columnconfigure(0, weight=1)
            if combo:
                w = ttk.Combobox(row_f, textvariable=self.vars[key],
                                 values=["Autorizadas CSV","Autorizadas SISPAG"],
                                 state="normal", width=width)
                w.grid(row=0, column=0, sticky="ew", padx=(0,6))
            else:
                e = styled_entry(row_f, textvariable=self.vars[key], width=width)
                if readonly: e.configure(state="readonly", readonlybackground=C["bg"])
                e.grid(row=0, column=0, sticky="ew", padx=(0,6))
            styled_button(row_f, "↗", lambda k=key: self._copy(k), small=True).grid(row=0,column=1)

        specs = [
            ("razao_social","Razão Social",0,0,False,28,False),
            ("cnpj","CNPJ",0,1,False,22,False),
            ("conta","Conta Corrente",1,0,False,22,False),
            ("valor_operacao","Valor da Operação",1,1,False,22,False),
            ("plataforma","Plataforma",2,0,False,14,False),
            ("regiao","Região da Plataforma",2,1,False,14,False),
            ("trader_espec","Trader / Espec",3,0,False,28,True),
            ("spread","Spread",3,1,False,14,False),
            ("prazo_min","Prazo Mínimo NF",4,0,False,14,False),
            ("prazo_max","Prazo Máximo NF",4,1,False,14,False),
            ("modalidade","Modalidade",5,0,True,28,False),
        ]
        for k,l,r,c,combo,w,ro in specs:
            field(k,l,r,c,combo,w,ro)

        act = tk.Frame(body, bg=C["bg"])
        act.pack(fill="x", padx=32, pady=(14,0))
        styled_button(act,"🔄  Gerar Resumo",  self._update_resumo, accent=True).pack(side="left")
        styled_button(act,"🧽  Limpar",          self._clear_all).pack(side="left", padx=(6,0))

        res_card = card_frame(body)
        res_card.pack(fill="x", padx=32, pady=(14,0))
        tk.Label(res_card, text="Resumo", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        make_hairline(res_card, bg=C["hair"]).pack(fill="x")

        txt_wrap = tk.Frame(res_card, bg=C["surface"], padx=18, pady=14)
        txt_wrap.pack(fill="x")
        txt_wrap.columnconfigure(0, weight=1)
        self.txt_resumo = tk.Text(txt_wrap, height=7, wrap="word", bd=0, relief="flat",
                                  bg=C["bg"], fg=C["ink"], insertbackground=C["accent"],
                                  highlightthickness=1, highlightbackground=C["hair"],
                                  font=("Segoe UI",9), padx=10, pady=8)
        self.txt_resumo.grid(row=0, column=0, sticky="ew", padx=(0,6))
        sc = MinimalScrollbar(txt_wrap, command=self.txt_resumo.yview, bg=C["surface"])
        sc.grid(row=0, column=1, sticky="ns")
        self.txt_resumo.configure(yscrollcommand=sc.set)
        bind_text_mousewheel(self.txt_resumo)

        foot = tk.Frame(body, bg=C["bg"])
        foot.pack(fill="x", padx=32, pady=(12,30))
        styled_button(foot,"📋  Copiar Resumo", self._copy_resumo, accent=True).pack(side="left")
        styled_button(foot,"💾  Salvar .txt",   self._save_resumo).pack(side="left", padx=(6,0))

    def on_show(self):
        if hasattr(self, "_sf"):
            self._sf.refresh_bindings()

    def _copy(self, k):
        v = self.vars[k].get()
        self.clipboard_clear(); self.clipboard_append(v)

    def _copy_resumo(self):
        v = self.txt_resumo.get("1.0","end-1c")
        self.clipboard_clear(); self.clipboard_append(v)

    def _save_resumo(self):
        c = self.txt_resumo.get("1.0","end-1c").strip()
        if not c: messagebox.showinfo("Vazio","Nada para salvar."); return
        p = filedialog.asksaveasfilename(defaultextension=".txt",
                                        filetypes=[("Texto","*.txt")], title="Salvar Resumo")
        if p:
            try: open(p,"w",encoding="utf-8").write(c)
            except Exception as e: messagebox.showerror("Erro",str(e))

    def _update_trader_espec(self):
        reg = self.vars["regiao"].get().strip()
        tr, sp = trader_espec_from_regiao(reg)
        self.vars["trader_espec"].set(f"{tr} / {sp}".strip(" /") if (tr or sp) else "")
        self._update_resumo()

    def _on_open_pdf(self):
        p = filedialog.askopenfilename(title="Selecione um PDF",
                                       filetypes=[("PDF","*.pdf")])
        if not p: return
        try:
            tl, tp = extract_text_from_pdf(p)
            t = tl
            if not t.strip():
                messagebox.showwarning("PDF sem texto","Não foi possível extrair texto."); return
            lines = t.splitlines()
            tn, tc = normalize_text_variants(t)
            raz  = extract_razao_social(t,lines,tn,tc,t_plain=tp) or ""
            cn   = extract_cnpj(t,tn,tc) or ""
            conta= extract_conta_corrente(lines,tn,tc) or ""
            plat = extract_plataforma(t,tn,tc) or ""
            reg  = extract_regiao(t,tn,tc) or ""
            val  = extract_valor(t,tn,tc) or ""
            sp   = extract_spread(t,tn,tc) or ""
            pmin = extract_prazo_min(t,tn,tc) or ""
            pmax = extract_prazo_max(t,tn,tc) or ""
            mod  = extract_modalidade(t,tn,tc) or ""
            liq  = "Débito em CC" if not RE_LIQ_CRED.search(t) else "Crédito em CC"
            prem = "com prêmio" if RE_PREMIO.search(t) else "sem prêmio"
            self.vars["razao_social"].set(raz)
            self.vars["cnpj"].set(only_digits(cn) if cn else "")
            self.vars["conta"].set(conta)
            self.vars["plataforma"].set(plat)
            self.vars["regiao"].set(reg)
            self.vars["valor_operacao"].set(val)
            self.vars["spread"].set(normalize_percent_br(sp) if sp else "")
            self.vars["prazo_min"].set(pmin)
            self.vars["prazo_max"].set(pmax)
            self.vars["modalidade"].set(mod)
            tr, spe = trader_espec_from_regiao(reg)
            self.vars["trader_espec"].set(f"{tr} / {spe}".strip(" /") if (tr or spe) else "")
            self.hidden["premio"] = prem
            self.hidden["liquidacao"] = liq
            self._update_resumo()
        except Exception as e:
            messagebox.showerror("Erro ao ler PDF", str(e))

    def _update_resumo(self):
        spread = self.vars["spread"].get().strip()
        pmin   = self.vars["prazo_min"].get().strip()
        pmax   = self.vars["prazo_max"].get().strip()
        plat   = self.vars["plataforma"].get().strip()
        reg    = self.vars["regiao"].get().strip()
        mod    = normalize_modalidade(self.vars["modalidade"].get() or "")
        prem   = self.hidden.get("premio","sem prêmio")
        liq    = self.hidden.get("liquidacao","Débito em CC")
        te     = self.vars["trader_espec"].get().strip()
        if spread and "a.a" not in spread.lower(): spread = normalize_percent_br(spread)
        troca  = infer_troca(mod)
        prazo  = ""
        if pmin and pmax: prazo = f"Prazo {pmin} a {pmax} dias"
        elif pmin: prazo = f"Prazo mínimo {pmin} dias"
        elif pmax: prazo = f"Prazo máximo {pmax} dias"
        pr = ""
        if plat and reg: pr = f"Plataforma {plat} – Região {reg}"
        elif plat: pr = f"Plataforma {plat}"
        elif reg: pr = f"Região {reg}"
        L = [f"Spread mínimo {spread} – {prem} – Troca de Arquivo {troca}" if spread
             else f"Spread mínimo – {prem} – Troca de Arquivo {troca}"]
        if prazo: L.append(prazo)
        if liq:   L.append(f"Liquidação {liq}")
        if pr:    L.append(pr)
        if te:    L.append(f"Trader/Espec {te}")
        self.txt_resumo.delete("1.0","end")
        self.txt_resumo.insert("1.0", "\n".join(L).strip())

    def _clear_all(self):
        for v in self.vars.values(): v.set("")
        self.hidden["premio"] = "sem prêmio"
        self.hidden["liquidacao"] = "Débito em CC"
        self.txt_resumo.delete("1.0","end")

class OperacoesInvertidoFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._xlsx_path = None
        self._overlay   = None
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=44, pady=(36, 0))
        eyebrow_label(hdr, "MESA DE OPERAÇÕES").pack(anchor="w")
        tk.Label(hdr, text="Operações Invertido", bg=C["bg"], fg=C["ink"],
                 font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(6, 0))
        tk.Label(hdr, text="Escolha uma ferramenta para continuar.",
                 bg=C["bg"], fg=C["ink_muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", pady=(20, 0))

        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=44, pady=(32, 0))
        for c in range(3):
            body.columnconfigure(c, weight=1, uniform="opc")

        self._make_option_card(
            body, 0, "▤", "Analisar Operações",
            "Importe uma planilha .xlsx para análise das operações.",
            self._open_analisar_overlay, "#5a9e72")

        self._make_option_card(
            body, 1, "⬡", "Limites Invertido",
            "Consulta o LTC e o limite disponível de cada cliente.",
            lambda: self.controller.show_frame("LimitesInvertido"), C["accent"])

        self._make_option_card(
            body, 2, "▦", "Histórico Operações",
            "Consulta o histórico de operações já realizadas.",
            self._open_historico_placeholder, "#8b72c9")

    def on_show(self):
        pass

    def _make_option_card(self, parent, col, icon, title, sub, command, color):
        pad = {0: (0, 6), 1: (6, 6), 2: (6, 0)}.get(col, (6, 6))
        outer = tk.Frame(parent, bg=C["surface"],
                         highlightthickness=1, highlightbackground=C["hair"],
                         cursor="hand2")
        outer.grid(row=0, column=col, sticky="nsew", padx=pad)

        top_line = tk.Frame(outer, bg=C["hair"], height=2)
        top_line.pack(fill="x")

        body_f = tk.Frame(outer, bg=C["surface"], padx=22, pady=22)
        body_f.pack(fill="both", expand=True)

        icon_lbl = tk.Label(body_f, text=icon, bg=C["surface"], fg=color,
                            font=("Segoe UI", 22))
        icon_lbl.pack(anchor="w")
        name_lbl = tk.Label(body_f, text=title, bg=C["surface"], fg=C["ink"],
                            font=("Segoe UI", 12, "bold"), anchor="w")
        name_lbl.pack(anchor="w", pady=(12, 4))
        sub_lbl = tk.Label(body_f, text=sub, bg=C["surface"], fg=C["ink_muted"],
                           font=("Segoe UI", 9), anchor="w", wraplength=200,
                           justify="left")
        sub_lbl.pack(anchor="w")
        arrow_lbl = tk.Label(body_f, text="Abrir →", bg=C["surface"], fg=C["ink_faint"],
                             font=("Segoe UI", 8, "bold"))
        arrow_lbl.pack(anchor="w", pady=(18, 0))

        widgets = [outer, top_line, body_f, icon_lbl, name_lbl, sub_lbl, arrow_lbl]

        def _enter(_e=None):
            outer.configure(bg=C["surface2"], highlightbackground=color)
            top_line.configure(bg=color)
            for w in (body_f, icon_lbl, name_lbl, sub_lbl):
                w.configure(bg=C["surface2"])
            arrow_lbl.configure(bg=C["surface2"], fg=color)

        def _leave(_e=None):
            outer.configure(bg=C["surface"], highlightbackground=C["hair"])
            top_line.configure(bg=C["hair"])
            for w in (body_f, icon_lbl, name_lbl, sub_lbl):
                w.configure(bg=C["surface"])
            arrow_lbl.configure(bg=C["surface"], fg=C["ink_faint"])

        for w in widgets:
            w.bind("<Button-1>", lambda _e: command())
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)

        return outer

    def _open_historico_placeholder(self):
        messagebox.showinfo(
            "Histórico Operações",
            "Histórico Operações estará disponível em breve.",
            parent=self.controller)

    def _open_analisar_overlay(self):
        if self._overlay is not None:
            return
        overlay = tk.Frame(self, bg="#0c0c0c")
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        overlay.bind("<Button-1>", lambda _e: self._close_analisar_overlay())
        self._overlay = overlay

        card = tk.Frame(overlay, bg=C["surface"],
                        highlightthickness=1, highlightbackground=C["hair"])
        card.place(relx=0.5, rely=0.5, anchor="center", width=460, height=320)
        card.bind("<Button-1>", lambda _e: "break")

        pad = tk.Frame(card, bg=C["surface"], padx=26, pady=22)
        pad.pack(fill="both", expand=True)

        top = tk.Frame(pad, bg=C["surface"])
        top.pack(fill="x")
        tk.Label(top, text="Analisar Operações", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        styled_button(top, "✕", self._close_analisar_overlay,
                      small=True).pack(side="right")

        tk.Label(pad, text="Selecione uma planilha .xlsx com as operações para análise.",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 9),
                 wraplength=400, justify="left").pack(anchor="w", pady=(10, 0))

        make_hairline(pad, bg=C["hair"]).pack(fill="x", pady=(16, 0))

        drop = tk.Frame(pad, bg=C["surface2"], highlightthickness=1,
                        highlightbackground=C["hair"])
        drop.pack(fill="x", pady=(16, 0))
        inner_drop = tk.Frame(drop, bg=C["surface2"], pady=24)
        inner_drop.pack(fill="x")

        has_file = bool(self._xlsx_path)
        icon_lbl = tk.Label(inner_drop, text=("✓" if has_file else "▤"),
                            bg=C["surface2"],
                            fg=(C["ok"] if has_file else C["ink_faint"]),
                            font=("Segoe UI", 20))
        icon_lbl.pack()
        file_lbl = tk.Label(inner_drop,
                            text=(os.path.basename(self._xlsx_path) if has_file
                                  else "Nenhum arquivo selecionado"),
                            bg=C["surface2"],
                            fg=(C["ink"] if has_file else C["ink_muted"]),
                            font=("Segoe UI", 9))
        file_lbl.pack(pady=(8, 0))

        self._overlay_file_lbl = file_lbl
        self._overlay_icon_lbl = icon_lbl

        foot = tk.Frame(pad, bg=C["surface"])
        foot.pack(fill="x", pady=(20, 0))
        self._overlay_action_btn = styled_button(
            foot, "Selecionar arquivo .xlsx…",
            self._overlay_action_click, accent=True)
        self._overlay_action_btn.pack(side="left")
        self._overlay_ready = bool(self._xlsx_path)
        if self._overlay_ready:
            self._set_overlay_analyze_state(animate=False)
        styled_button(foot, "Cancelar",
                      self._close_analisar_overlay).pack(side="right")

    def _overlay_action_click(self):
        if self._overlay_ready:
            self._start_analyze()
        else:
            self._pick_xlsx()

    def _set_overlay_analyze_state(self, animate=True):
        btn = getattr(self, "_overlay_action_btn", None)
        if btn is None or not btn.winfo_exists():
            return
        self._overlay_ready = True
        btn.configure(command=self._start_analyze)
        if not animate:
            btn.configure(text="Analisar", bg=C["ok_dim"], fg=C["ok"],
                          activebackground=C["ok"], activeforeground=C["bg"])
            btn.bind("<Enter>", lambda _: btn.configure(bg=C["ok"], fg=C["bg"]))
            btn.bind("<Leave>", lambda _: btn.configure(bg=C["ok_dim"], fg=C["ok"]))
            return

        steps = [
            (C["accent_dim"], C["accent"], "Selecionar arquivo .xlsx…"),
            (C["surface3"],   C["ink"],    "Preparando…"),
            ("#2d4a38",       "#6fbf96",   "Analisar"),
            (C["ok_dim"],     C["ok"],     "Analisar"),
        ]
        def _step(i=0):
            if btn is None or not btn.winfo_exists():
                return
            bg, fg, text = steps[min(i, len(steps) - 1)]
            btn.configure(text=text, bg=bg, fg=fg,
                          activebackground=C["ok"], activeforeground=C["bg"])
            if i < len(steps) - 1:
                self.after(55, lambda: _step(i + 1))
            else:
                btn.bind("<Enter>", lambda _: btn.configure(bg=C["ok"], fg=C["bg"]))
                btn.bind("<Leave>", lambda _: btn.configure(bg=C["ok_dim"], fg=C["ok"]))
        _step()

    def _start_analyze(self):
        if not self._xlsx_path:
            return
        self.controller.invertido_xlsx_path = self._xlsx_path
        self._close_analisar_overlay()
        self.controller.show_frame("AnalisarOperacoes")

    def _pick_xlsx(self):
        p = filedialog.askopenfilename(
            title="Selecionar planilha de operações",
            filetypes=[("Planilha Excel", "*.xlsx")])
        if not p:
            return
        self._xlsx_path = p
        if hasattr(self, "_overlay_file_lbl") and self._overlay_file_lbl.winfo_exists():
            self._overlay_file_lbl.configure(text=os.path.basename(p), fg=C["ink"])
            self._overlay_icon_lbl.configure(text="✓", fg=C["ok"])
        self._set_overlay_analyze_state(animate=True)

    def _close_analisar_overlay(self):
        if self._overlay is not None:
            try: self._overlay.destroy()
            except Exception: pass
            self._overlay = None
        self._overlay_ready = bool(self._xlsx_path)


class AnalisarOperacoesFrame(tk.Frame, ThreadSafeUIMixin):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._groups = []
        self._cards = []
        self._limit_btns = {}
        self._worker_running = False
        self._last_path = None
        self._detail_overlay = None
        self._limite_overlay = None
        self._init_ui_queue()
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24, 0))
        styled_button(hdr, "← Voltar",
                      lambda: (self._close_detalhes(),
                               self.controller.show_frame("OperacoesInvertido"))).pack(side="left")
        tk.Label(hdr, text="Analisar Operações", bg=C["bg"], fg=C["ink"],
                 font=("Georgia", 18, "bold")).pack(side="left", padx=(14, 0))

        sub = tk.Frame(self, bg=C["bg"])
        sub.pack(fill="x", padx=32)
        self._sub_lbl = tk.Label(
            sub, text="Importe e analise as operações da planilha selecionada.",
            bg=C["bg"], fg=C["ink_muted"], font=("Segoe UI", 9))
        self._sub_lbl.pack(anchor="w", pady=(4, 0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(16, 0))

        self._loading_outer = tk.Frame(self, bg=C["bg"])
        self._loading_outer.pack(fill="x", padx=32, pady=(28, 0))
        loading_card = tk.Frame(self._loading_outer, bg=C["surface"],
                                highlightthickness=1, highlightbackground=C["hair"])
        loading_card.pack(fill="x")
        loading_body = tk.Frame(loading_card, bg=C["surface"], padx=18, pady=16)
        loading_body.pack(fill="x")
        self._loading_icon = tk.Label(loading_body, text="◐", bg=C["surface"],
                                      fg=C["ok"], font=("Segoe UI", 16, "bold"))
        self._loading_icon.pack()
        self._loading_lbl = tk.Label(
            loading_body, text="Analisando planilha…", bg=C["surface"],
            fg=C["ink_muted"], font=("Segoe UI", 9))
        self._loading_lbl.pack(pady=(8, 0))
        self._loading_spin = [None]
        self._loading_angle = [0]

        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._sf.link_wheel(self)
        self._grid_outer = self._sf.inner
        self._grid_outer.configure(bg=C["bg"])

        self._grid = tk.Frame(self._grid_outer, bg=C["bg"])
        self._grid.pack(padx=32, pady=(16, 24), fill="x")
        for c in range(3):
            self._grid.columnconfigure(c, weight=1, uniform="acards")

    def on_show(self):
        self._sf.refresh_bindings()
        path = getattr(self.controller, "invertido_xlsx_path", None)
        if not path:
            self.controller.show_frame("OperacoesInvertido")
            return
        fname = os.path.basename(path)
        self._sub_lbl.configure(text=f"Planilha: {fname}")
        if path == self._last_path and self._groups and not self._worker_running:
            self._refresh_all_limit_buttons()
            return
        self._last_path = path
        self._reset_view()
        self._start_worker(path)
        self.controller.register_limites_listener(self._on_limites_update)

    def _on_limites_update(self, _cnpj=None):
        self._ui(self._refresh_all_limit_buttons)

    def _reset_view(self):
        self._groups = []
        self._cards = []
        for w in self._grid.winfo_children():
            w.destroy()
        self._loading_lbl.configure(text="Analisando planilha…", fg=C["ink_muted"])
        self._loading_icon.configure(text="◐", fg=C["ok"])
        self._loading_outer.pack(fill="x", padx=32, pady=(28, 0))
        self._start_loading_anim()

    def _start_loading_anim(self):
        if self._loading_spin[0]:
            try:
                self.after_cancel(self._loading_spin[0])
            except Exception:
                pass
        self._loading_angle[0] = 0

        def tick():
            self._loading_angle[0] = (self._loading_angle[0] + 1) % 4
            if self._loading_icon.winfo_exists():
                self._loading_icon.configure(
                    text=["◐", "◓", "◑", "◒"][self._loading_angle[0]])
            self._loading_spin[0] = self.after(120, tick)

        tick()

    def _stop_loading_anim(self):
        if self._loading_spin[0]:
            try:
                self.after_cancel(self._loading_spin[0])
            except Exception:
                pass
            self._loading_spin[0] = None

    def _make_group_card(self, group, row, col):
        bg = C["surface"]
        outer = tk.Frame(self._grid, bg=C["bg"])
        outer.grid(row=row, column=col, sticky="new", padx=5, pady=5)

        card = tk.Frame(outer, bg=bg, highlightthickness=1,
                        highlightbackground=C["hair"], bd=0)
        card.pack(fill="x")
        top_bar = tk.Frame(card, bg=C["ok"], height=2)
        top_bar.pack(fill="x")

        body = tk.Frame(card, bg=bg, padx=14, pady=12)
        body.pack(fill="x")

        tk.Label(body, text=group["nome_sacado"], bg=bg, fg=C["ink"],
                 font=("Segoe UI", 9, "bold"), wraplength=170,
                 justify="center").pack()

        info = tk.Frame(body, bg=bg)
        info.pack(pady=(10, 0), fill="x")
        for label, value, accent in (
            ("Notas", str(group["count"]), False),
            ("Montante", group["valor_total"], True),
        ):
            row_f = tk.Frame(info, bg=bg)
            row_f.pack(fill="x", pady=(0, 3))
            tk.Label(row_f, text=label, bg=bg, fg=C["ink_faint"],
                     font=("Segoe UI", 7, "bold"), width=9, anchor="w").pack(side="left")
            fg = C["ok"] if accent else C["ink_muted"]
            tk.Label(row_f, text=value, bg=bg, fg=fg,
                     font=("Segoe UI", 8, "bold" if accent else "normal"),
                     anchor="w").pack(side="left", fill="x", expand=True)

        btn_row = tk.Frame(body, bg=bg)
        btn_row.pack(fill="x", pady=(12, 0))
        status, label = self._evaluate_group_limite(group)
        lim_btn = styled_button_limite(
            btn_row, label,
            lambda g=group: self._open_limite_modal(g),
            variant=self._limite_variant(status),
            small=True)
        lim_btn.pack(side="left")
        styled_button(btn_row, "Detalhes →",
                      lambda k=group["nome_sacado"]: self._open_detalhes(k),
                      accent=True, small=True).pack(side="right")

        self._limit_btns[self._group_limite_key(group)] = lim_btn

        return {"outer": outer, "lim_btn": lim_btn, "group": group}

    def _limite_variant(self, status):
        return {"ok": "ok", "quase": "warn", "nao_validado": "warn",
                "nao_encontrado": "warn", "validando": "idle",
                "insuficiente": "err"}.get(status, "warn")

    def _group_limite_key(self, group):
        return group.get("doc_sacado") or _normalize_sacado_key(group["nome_sacado"])

    def _get_limite_data(self, group):
        cnpj = group.get("doc_sacado") or ""
        if not cnpj:
            return None
        return getattr(self.controller, "invertido_limites_cache", {}).get(cnpj)

    def _evaluate_group_limite(self, group):
        cnpj = group.get("doc_sacado") or ""
        if not cnpj or cnpj not in LIMITE_INVERTIDO_CNPJS:
            return "nao_encontrado", "Limite não encontrado"
        limite_data = self._get_limite_data(group)
        lf = self.controller.frames.get("LimitesInvertido")
        if lf and lf._worker_running and not limite_data:
            return "validando", "Validando…"
        if not limite_data:
            return "nao_validado", "Limites não validados"
        return _evaluate_limite_operacao(group["total"], limite_data)

    def _refresh_all_limit_buttons(self):
        for key, btn in list(self._limit_btns.items()):
            if not btn.winfo_exists():
                continue
            group = next((g for g in self._groups if self._group_limite_key(g) == key), None)
            if not group:
                continue
            status, label = self._evaluate_group_limite(group)
            _set_limite_button(btn, label, self._limite_variant(status))

    def _close_limite_modal(self):
        if self._limite_overlay is not None:
            try:
                self._limite_overlay.destroy()
            except Exception:
                pass
            self._limite_overlay = None

    def _start_limites_validation(self):
        lf = self.controller.frames.get("LimitesInvertido")
        if lf is None or lf._worker_running:
            return
        self._close_limite_modal()
        if lf._started and not lf._worker_running:
            lf._restart_limites()
        else:
            lf._started = True
            lf._cancel_requested = False
            lf._setup_cards()
            lf._start_worker()
        self._refresh_all_limit_buttons()

    def _open_limite_modal(self, group):
        if self._limite_overlay is not None:
            return
        status, label = self._evaluate_group_limite(group)
        limite_data = self._get_limite_data(group)

        overlay = tk.Frame(self, bg="#0c0c0c")
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._limite_overlay = overlay

        card = tk.Frame(overlay, bg=C["surface"],
                        highlightthickness=1, highlightbackground=C["hair"])
        card.place(relx=0.5, rely=0.5, anchor="center", width=420, height=300)
        card.bind("<Button-1>", lambda _e: "break")

        pad = tk.Frame(card, bg=C["surface"], padx=24, pady=20)
        pad.pack(fill="both", expand=True)

        top = tk.Frame(pad, bg=C["surface"])
        top.pack(fill="x")
        title_fg = {"ok": C["ok"], "quase": C["warn"], "insuficiente": C["err"],
                    "nao_encontrado": C["warn"]}.get(status, C["warn"])
        tk.Label(top, text=label, bg=C["surface"], fg=title_fg,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        styled_button(top, "✕", self._close_limite_modal, small=True).pack(side="right")

        tk.Label(pad, text=group["nome_sacado"], bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI", 9), wraplength=360, justify="left").pack(
            anchor="w", pady=(10, 0))

        make_hairline(pad, bg=C["hair"]).pack(fill="x", pady=(14, 0))

        body = tk.Frame(pad, bg=C["surface"])
        body.pack(fill="both", expand=True, pady=(12, 0))

        if status == "nao_validado":
            tk.Label(
                body,
                text=("Os limites ainda não foram consultados para este sacado.\n"
                      "Deseja iniciar a validação agora?"),
                bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 9),
                wraplength=360, justify="left",
            ).pack(anchor="w")
            foot = tk.Frame(pad, bg=C["surface"])
            foot.pack(fill="x", pady=(16, 0))
            styled_button(foot, "Cancelar", self._close_limite_modal, small=True).pack(side="right")
            styled_button(foot, "Iniciar validação",
                          self._start_limites_validation, accent=True, small=True).pack(
                side="right", padx=(0, 8))
            return

        if status == "nao_encontrado":
            doc = group.get("doc_sacado") or "—"
            tk.Label(
                body,
                text=(f"CNPJ {doc} não consta na lista de Limites Invertido.\n"
                      "Este sacado não possui consulta de limite cadastrada no app."),
                bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 9),
                wraplength=360, justify="left",
            ).pack(anchor="w")
            foot = tk.Frame(pad, bg=C["surface"])
            foot.pack(fill="x", pady=(16, 0))
            styled_button(foot, "Fechar", self._close_limite_modal, accent=True, small=True).pack(
                side="right")
            return

        mont = group["valor_total"]
        lines = [f"Montante do grupo: {mont}"]
        if limite_data:
            if limite_data.get("ltc_str"):
                lines.append(f"LTC ativo · vence {limite_data['ltc_str']}")
            lines.append(f"Limite disp. (fornecedor): {_fmt_limite_int(limite_data.get('limite_disp'))}")
            limite = limite_data.get("limite_disp")
            if limite is not None:
                slack = limite - int(group["total"].quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                lines.append(f"Sobra: {_fmt_limite_int(slack)}")
            if limite_data.get("via"):
                lines.append(f"Consulta via {limite_data['via']}")
            if status == "quase":
                lines.append(f"\nSobra abaixo de R$ {LIMITE_SOBRA_MIN:,}".replace(",", "."))
            elif status == "insuficiente":
                if limite_data.get("state") == "ltc_expired":
                    lines.append("\nLTC vencido ou indisponível.")
                else:
                    lines.append("\nMontante superior ao limite disponível.")

        tk.Label(body, text="\n".join(lines), bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI", 9), wraplength=360, justify="left").pack(anchor="w")

        foot = tk.Frame(pad, bg=C["surface"])
        foot.pack(fill="x", pady=(16, 0))
        styled_button(foot, "Fechar", self._close_limite_modal, accent=True, small=True).pack(
            side="right")

    def _find_group(self, nome_sacado):
        key = _normalize_sacado_key(nome_sacado)
        for g in self._groups:
            if _normalize_sacado_key(g["nome_sacado"]) == key:
                return g
        return None

    def _bind_modal_scroll(self, widgets, sf):
        def _mw(event):
            sf._scroll_mousewheel(event)
            return "break"
        for widget in widgets:
            widget.bind("<MouseWheel>", _mw)
            widget.bind("<Button-4>", _mw)
            widget.bind("<Button-5>", _mw)

    def _open_detalhes(self, nome_sacado):
        group = self._find_group(nome_sacado)
        if group is None or self._detail_overlay is not None:
            return

        overlay = tk.Frame(self, bg="#0c0c0c")
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._detail_overlay = overlay

        card = tk.Frame(overlay, bg=C["surface"],
                        highlightthickness=1, highlightbackground=C["hair"])
        card.place(relx=0.5, rely=0.5, anchor="center", width=720, height=580)
        card.bind("<Button-1>", lambda _e: "break")

        pad = tk.Frame(card, bg=C["surface"], padx=28, pady=22)
        pad.pack(fill="both", expand=True)

        top = tk.Frame(pad, bg=C["surface"])
        top.pack(fill="x")
        tk.Label(top, text=group["nome_sacado"], bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 13, "bold"), wraplength=560,
                 justify="left").pack(side="left", fill="x", expand=True)
        styled_button(top, "✕", self._close_detalhes, small=True).pack(side="right")

        tk.Label(
            pad,
            text=f"{group['count']} nota(s) · {group['valor_total']}",
            bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(8, 0))

        make_hairline(pad, bg=C["hair"]).pack(fill="x", pady=(16, 0))

        scroll_wrap = tk.Frame(pad, bg=C["surface"], height=420)
        scroll_wrap.pack(fill="both", expand=True, pady=(12, 0))
        scroll_wrap.pack_propagate(False)

        sf = ScrollableFrame(scroll_wrap, bg=C["surface"])
        sf.pack(fill="both", expand=True)
        sf.link_wheel(scroll_wrap)
        sf.link_wheel(pad)
        sf.link_wheel(card)
        list_outer = sf.inner
        list_outer.configure(bg=C["surface"])

        notas = list(group["notas"])
        for idx, op in enumerate(notas, start=1):
            item = tk.Frame(list_outer, bg=C["surface2"],
                            highlightthickness=1, highlightbackground=C["hair"])
            item.pack(fill="x", pady=(0, 8))

            head = tk.Frame(item, bg=C["surface2"], padx=14, pady=10)
            head.pack(fill="x")
            tk.Label(head, text=f"Nota {idx} de {len(notas)}", bg=C["surface2"],
                     fg=C["ink_faint"], font=("Segoe UI", 7, "bold")).pack(anchor="w")

            body = tk.Frame(item, bg=C["surface2"], padx=14, pady=12)
            body.pack(fill="x")
            fields = [
                ("NF", op["nf"] or "—", False),
                ("Valor", op["valor"], True),
                ("Inclusão", op["data_inclusao"] or "—", False),
                ("Vencimento", op["data_vencimento"] or "—", False),
                ("Prazo", f"{op['prazo']} dias" if op.get("prazo") else "—", False),
            ]
            for label, value, accent in fields:
                row_f = tk.Frame(body, bg=C["surface2"])
                row_f.pack(fill="x", pady=(0, 3))
                tk.Label(row_f, text=label, bg=C["surface2"], fg=C["ink_faint"],
                         font=("Segoe UI", 7, "bold"), width=10, anchor="w").pack(side="left")
                fg = C["ok"] if accent else C["ink"]
                tk.Label(row_f, text=value, bg=C["surface2"], fg=fg,
                         font=("Segoe UI", 8, "bold" if accent else "normal"),
                         anchor="w").pack(side="left", fill="x", expand=True)

        self._bind_modal_scroll([overlay, card, pad, scroll_wrap], sf)

        def _sync_modal():
            sf.update_idletasks()
            sf._sync_scrollregion()
            sf.refresh_bindings()
        self.after_idle(_sync_modal)

    def _close_detalhes(self):
        if self._detail_overlay is not None:
            try:
                self._detail_overlay.destroy()
            except Exception:
                pass
            self._detail_overlay = None

    # ── Alerta de Validação ──────────────────────────────────────────────────
    def _show_alert_dialog(self, ops, alerts_map):
        """Janela modal com alertas de validação. O usuário aceita ou rejeita cada nota."""
        dlg = tk.Toplevel(self.controller)
        dlg.title("Alertas de Validação")
        dlg.configure(bg=C["surface"])
        dlg.geometry("780x620")
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.transient(self.controller)
        dlg.attributes("-topmost", True)

        # Impede fechar sem confirmar
        def _on_close():
            pass
        dlg.protocol("WM_DELETE_WINDOW", _on_close)

        # Header
        hdr = tk.Frame(dlg, bg=C["surface"], padx=24, pady=18)
        hdr.pack(fill="x")

        hdr_top = tk.Frame(hdr, bg=C["surface"])
        hdr_top.pack(fill="x")

        tk.Label(hdr_top, text="⚠", bg=C["surface"], fg=C["warn"],
                 font=("Segoe UI", 18)).pack(side="left")
        tk.Label(hdr_top, text="Alertas de Validação", bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=(8, 0))

        alert_count = len(alerts_map)
        tk.Label(hdr, text=f"Foram encontrados alertas em {alert_count} nota{'s' if alert_count != 1 else ''}. "
                           f"Revise e decida se deseja manter cada nota.",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI", 9),
                 wraplength=700, justify="left").pack(anchor="w", pady=(8, 0))

        # Bulk actions
        bulk_row = tk.Frame(hdr, bg=C["surface"])
        bulk_row.pack(anchor="w", pady=(10, 0))

        decisions = {}  # op_index -> "accept" or "reject"

        def _accept_all():
            for idx in alerts_map:
                decisions[idx] = "accept"
            _refresh_all_items()

        def _reject_all():
            for idx in alerts_map:
                decisions[idx] = "reject"
            _refresh_all_items()

        styled_button(bulk_row, "✓ Aceitar Todos", _accept_all,
                      accent=True, small=True).pack(side="left")
        styled_button(bulk_row, "✕ Rejeitar Todos", _reject_all,
                      danger=True, small=True).pack(side="left", padx=(6, 0))

        make_hairline(dlg, bg=C["hair"]).pack(fill="x")

        # Scrollable alert list
        list_wrap = tk.Frame(dlg, bg=C["surface"])
        list_wrap.pack(fill="both", expand=True)

        sf = ScrollableFrame(list_wrap, bg=C["surface"])
        sf.pack(fill="both", expand=True)
        sf.link_wheel(dlg)
        alert_list = sf.inner
        alert_list.configure(bg=C["surface"])

        item_frames = {}

        def _refresh_all_items():
            for idx in alerts_map:
                _refresh_item(idx)

        def _refresh_item(idx):
            frame_data = item_frames.get(idx)
            if frame_data is None:
                return
            card_f = frame_data["card"]
            accept_btn = frame_data["accept_btn"]
            reject_btn = frame_data["reject_btn"]
            status_lbl = frame_data["status_lbl"]

            dec = decisions.get(idx, "accept")
            if dec == "accept":
                card_f.configure(highlightbackground=C["ok"])
                accept_btn.configure(bg=C["ok"], fg=C["bg"],
                                     activebackground=C["ok"], activeforeground=C["bg"])
                reject_btn.configure(bg=C["surface3"], fg=C["ink_muted"],
                                     activebackground=C["surface3"], activeforeground=C["ink"])
                status_lbl.configure(text="✓ Aceita", fg=C["ok"])
            else:
                card_f.configure(highlightbackground=C["err"])
                accept_btn.configure(bg=C["surface3"], fg=C["ink_muted"],
                                     activebackground=C["surface3"], activeforeground=C["ink"])
                reject_btn.configure(bg=C["err"], fg=C["bg"],
                                     activebackground=C["err"], activeforeground=C["bg"])
                status_lbl.configure(text="✕ Rejeitada", fg=C["err"])

        for op_idx in sorted(alerts_map.keys()):
            op = ops[op_idx]
            op_alerts = alerts_map[op_idx]
            decisions[op_idx] = "accept"  # Default: accept

            card_f = tk.Frame(alert_list, bg=C["surface2"],
                              highlightthickness=2, highlightbackground=C["ok"],
                              padx=0, pady=0)
            card_f.pack(fill="x", padx=20, pady=(8, 0))

            # Header da nota
            note_hdr = tk.Frame(card_f, bg=C["surface2"], padx=16, pady=10)
            note_hdr.pack(fill="x")
            tk.Label(note_hdr, text=op.get("nome_sacado", "—"), bg=C["surface2"],
                     fg=C["ink"], font=("Segoe UI", 10, "bold")).pack(side="left")

            status_lbl = tk.Label(note_hdr, text="✓ Aceita", bg=C["surface2"],
                                  fg=C["ok"], font=("Segoe UI", 8, "bold"))
            status_lbl.pack(side="right")

            # Detalhes
            detail = tk.Frame(card_f, bg=C["surface2"], padx=16, pady=(0, 8))
            detail.pack(fill="x")

            detail_row = tk.Frame(detail, bg=C["surface2"])
            detail_row.pack(fill="x")
            for lbl, val in [
                ("NF", op.get("nf") or "—"),
                ("Valor", op.get("valor") or "—"),
                ("Vencimento", op.get("data_vencimento") or "—"),
                ("Prazo", f"{op.get('prazo', '—')} dias" if op.get("prazo") else "—"),
            ]:
                f = tk.Frame(detail_row, bg=C["surface2"])
                f.pack(side="left", padx=(0, 16))
                tk.Label(f, text=lbl, bg=C["surface2"], fg=C["ink_faint"],
                         font=("Segoe UI", 7, "bold")).pack(anchor="w")
                tk.Label(f, text=val, bg=C["surface2"], fg=C["ink"],
                         font=("Segoe UI", 9)).pack(anchor="w")

            # Motivos do alerta
            motives_f = tk.Frame(card_f, bg=C["surface2"], padx=16, pady=(0, 8))
            motives_f.pack(fill="x")
            for motive in op_alerts:
                mf = tk.Frame(motives_f, bg=C["surface2"])
                mf.pack(fill="x", pady=(1, 0))
                tk.Label(mf, text="⚠", bg=C["surface2"], fg=C["warn"],
                         font=("Segoe UI", 8)).pack(side="left", padx=(0, 4))
                tk.Label(mf, text=motive, bg=C["surface2"], fg=C["warn"],
                         font=("Segoe UI", 8)).pack(side="left")

            # Botões Aceitar/Rejeitar
            btn_row = tk.Frame(card_f, bg=C["surface2"], padx=16, pady=(0, 12))
            btn_row.pack(fill="x")

            accept_btn = styled_button(btn_row, "✓ Aceitar",
                                       lambda idx=op_idx: (decisions.__setitem__(idx, "accept"), _refresh_item(idx)),
                                       accent=True, small=True)
            accept_btn.pack(side="left")

            reject_btn = styled_button(btn_row, "✕ Rejeitar",
                                       lambda idx=op_idx: (decisions.__setitem__(idx, "reject"), _refresh_item(idx)),
                                       danger=True, small=True)
            reject_btn.pack(side="left", padx=(6, 0))

            item_frames[op_idx] = {
                "card": card_f,
                "accept_btn": accept_btn,
                "reject_btn": reject_btn,
                "status_lbl": status_lbl,
            }

        # Footer
        make_hairline(dlg, bg=C["hair"]).pack(fill="x")
        foot = tk.Frame(dlg, bg=C["surface"], padx=24, pady=14)
        foot.pack(fill="x")

        rejected_summary = tk.Label(foot, text="", bg=C["surface"], fg=C["ink_faint"],
                                    font=("Segoe UI", 8))
        rejected_summary.pack(side="left")

        def _confirm():
            rejected_indices = {idx for idx, dec in decisions.items() if dec == "reject"}
            # Notas sem alerta são sempre aceitas
            filtered_ops = [op for i, op in enumerate(ops) if i not in rejected_indices]
            dlg.destroy()
            self._render_results(filtered_ops)

        styled_button(foot, "Confirmar Decisões", _confirm,
                      accent=True).pack(side="right")

        # Atualiza resumo ao vivo
        def _update_summary(*_):
            n_rejected = sum(1 for d in decisions.values() if d == "reject")
            n_accepted = len(decisions) - n_rejected
            rejected_summary.configure(
                text=f"{n_accepted} aceita{'s' if n_accepted != 1 else ''} · "
                     f"{n_rejected} rejeitada{'s' if n_rejected != 1 else ''}")

        # Observer pattern simples: refresh summary on any decision change
        _orig_setitem = decisions.__setitem__
        def _patched_setitem(key, value):
            _orig_setitem(key, value)
            _update_summary()
        decisions.__setitem__ = _patched_setitem
        _update_summary()

        dlg.update_idletasks()
        sw = dlg.winfo_screenwidth()
        sh = dlg.winfo_screenheight()
        w = dlg.winfo_reqwidth()
        h = dlg.winfo_reqheight()
        dlg.geometry(f"+{max(0,(sw - w) // 2)}+{max(0,(sh - h) // 2)}")

    def _render_results(self, ops):
        self._stop_loading_anim()
        self._loading_outer.pack_forget()
        for w in self._grid.winfo_children():
            w.destroy()
        self._cards = []
        groups = _group_invertido_ops(ops)
        self._groups = groups
        if not groups:
            empty = tk.Frame(self._grid, bg=C["bg"])
            empty.grid(row=0, column=0, columnspan=3, sticky="ew", pady=24)
            tk.Label(empty, text="Nenhuma operação encontrada na planilha.",
                     bg=C["bg"], fg=C["ink_muted"],
                     font=("Segoe UI", 10)).pack()
            return
        for idx, group in enumerate(groups):
            row, col = divmod(idx, 3)
            self._cards.append(self._make_group_card(group, row, col))
        total_notas = sum(g["count"] for g in groups)
        self._sub_lbl.configure(
            text=(f"{len(groups)} grupo(s) · {total_notas} nota(s) · "
                  f"{os.path.basename(self._last_path or '')}"))
        self._limit_btns = {}
        for card in self._cards:
            if card.get("lim_btn") and card.get("group"):
                self._limit_btns[self._group_limite_key(card["group"])] = card["lim_btn"]
        self.update_idletasks()
        self._sf._sync_scrollregion()
        self._sf.refresh_bindings()
        self._refresh_all_limit_buttons()

    def _show_error(self, msg):
        self._stop_loading_anim()
        self._loading_lbl.configure(text=msg, fg=C["err"])
        self._loading_icon.configure(text="✗", fg=C["err"])

    def _start_worker(self, path):
        if self._worker_running:
            return
        self._worker_running = True
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _worker(self, path):
        try:
            ops = _parse_invertido_xlsx(path)
            # Validar todas as operações
            alerts_map = {}
            for i, op in enumerate(ops):
                op_alerts = _validate_operation(op)
                if op_alerts:
                    alerts_map[i] = op_alerts

            if alerts_map:
                # Mostrar janela de alertas na thread UI
                self._ui(lambda: self._show_alert_dialog(ops, alerts_map))
            else:
                self._ui(lambda: self._render_results(ops))
        except Exception as e:
            self._ui(lambda m=str(e): self._show_error(f"Erro ao analisar: {m}"))
        finally:
            self._worker_running = False


class LimitesInvertidoFrame(tk.Frame, ThreadSafeUIMixin):
    COL_OK   = C["ok"]
    COL_WARN = C["warn"]
    COL_ERR  = C["err"]

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller  = controller
        self._cards      = []
        self._worker_running  = False
        self._cancel_requested= False
        self._browser    = None
        self._started    = False
        self._restart_pending = False
        self._init_ui_queue()
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        styled_button(hdr, "← Voltar",
                      lambda: self.controller.show_frame("OperacoesInvertido")).pack(side="left")
        self._restart_btn = styled_button(hdr, "↻  Atualizar",
                                          self._restart_limites)
        self._restart_btn.pack(side="right")

        tk.Label(hdr, text="Limites Invertido", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left", padx=(14,0))

        sub = tk.Frame(self, bg=C["bg"])
        sub.pack(fill="x", padx=32)
        tk.Label(sub, text="Consulta o LTC e o limite disponível de cada cliente.",
                 bg=C["bg"], fg=C["ink_muted"], font=("Segoe UI",9)).pack(anchor="w", pady=(4,0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(16,0))

        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._sf.link_wheel(self)
        self._grid_outer = self._sf.inner
        self._grid_outer.configure(bg=C["bg"])

        self._grid = tk.Frame(self._grid_outer, bg=C["bg"])
        self._grid.pack(padx=32, pady=(16,24), fill="x")
        for c in range(3):
            self._grid.columnconfigure(c, weight=1, uniform="lcards")

    def on_show(self):
        self._sf.refresh_bindings()
        if self._started: return
        self._started = True
        self._cancel_requested = False
        self._setup_cards()
        self._start_worker()

    def _restart_limites(self):
        self._restart_pending  = True
        self._cancel_requested = True
        br = getattr(self,"_browser",None)
        if br:
            try: br.close()
            except: pass
        self.after(200, self._check_restart)

    def _check_restart(self):
        if not self._restart_pending: return
        if self._worker_running: self.after(220, self._check_restart); return
        self._restart_pending = False
        self._cards = []
        for w in self._grid.winfo_children(): w.destroy()
        self._started = False
        self._cancel_requested = False
        self._setup_cards()
        self._started = True
        self._start_worker()

    def _setup_cards(self):
        for w in self._grid.winfo_children(): w.destroy()
        self._cards = []
        all_clients = list(BPM_CLIENT_DATA.keys())
        for idx, name in enumerate(all_clients):
            is_mapped = name in MAPPED_CLIENTS
            is_mirror = name in MIRROR_CLIENTS
            row, col  = divmod(idx, 3)
            c = self._make_card(name, row, col, is_mapped, is_mirror)
            self._cards.append(c)
        for i in range(len(self._cards)):
            self.after(50*i, lambda i=i: self._reveal_card(i))

    def _make_card(self, name, row, col, is_mapped, is_mirror):
        bg = C["surface"]; bord = C["hair"]
        outer = tk.Frame(self._grid, bg=C["bg"])
        outer.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)
        outer.grid_remove()

        card = tk.Frame(outer, bg=bg, highlightthickness=1, highlightbackground=bord, bd=0)
        card.pack(fill="both", expand=True)
        top_bar = tk.Frame(card, bg=bord, height=2)
        top_bar.pack(fill="x")

        body = tk.Frame(card, bg=bg, padx=14, pady=12)
        body.pack(fill="both", expand=True)

        icon_lbl = tk.Label(body, text="–", bg=bg, fg=C["ink_faint"],
                            font=("Segoe UI",14,"bold"))
        icon_lbl.pack()
        tk.Label(body, text=name, bg=bg, fg=C["ink_muted"],
                 font=("Segoe UI",9,"bold")).pack(pady=(6,0))
        if is_mirror:
            src = MIRROR_CLIENTS[name]
            tk.Label(body, text=f"via {src}", bg=bg, fg=C["ink_faint"],
                     font=("Segoe UI",7,"italic")).pack()

        info_lbl = tk.Label(body, text="", bg=bg, fg=C["ink_muted"],
                            font=("Segoe UI",8), wraplength=160, justify="center")
        info_lbl.pack(pady=(6,0))

        status_row = tk.Frame(body, bg=bg)
        status_row.pack(pady=(8,0))
        status_lbl = tk.Label(status_row, text="EM ESPERA" if is_mapped or is_mirror else "A MAPEAR",
                              bg=bg, fg=C["ink_faint"], font=("Segoe UI",7,"bold"))
        status_lbl.pack()

        spin_id = [None]
        angle   = [0]

        def tick():
            angle[0] = (angle[0]+20)%360
            if spin_id[0] is not None:
                status_lbl.configure(text=["◐","◓","◑","◒"][angle[0]//90])
            spin_id[0] = outer.after(120, tick) if spin_id[0] else None

        return {
            "outer":outer,"card":card,"top_bar":top_bar,"body":body,
            "icon_lbl":icon_lbl,"info_lbl":info_lbl,"status_lbl":status_lbl,
            "status_row":status_row,
            "spin_id":spin_id,"angle":angle,"tick":tick,
            "is_mapped":is_mapped,"is_mirror":is_mirror,"name":name,"state":"init"
        }

    def _reveal_card(self, idx):
        if idx >= len(self._cards): return
        c = self._cards[idx]
        c["outer"].grid()
        state = "waiting" if c["is_mapped"] else ("mirror_waiting" if c["is_mirror"] else "pending")
        self._set_state(idx, state)

    def _set_state(self, idx, state, info=""):
        if idx >= len(self._cards): return
        c = self._cards[idx]
        if c["spin_id"][0]:
            try: c["outer"].after_cancel(c["spin_id"][0])
            except: pass
            c["spin_id"][0] = None
        c["state"] = state

        cfg = {
            "waiting":       (C["surface"],  C["hair"],     "–",  C["ink_faint"],  "EM ESPERA",   C["ink_faint"]),
            "mirror_waiting":(C["surface"],  C["hair"],     "–",  C["ink_faint"],  "AGUARDANDO",  C["ink_faint"]),
            "pending":       (C["surface2"], C["hair"],     "–",  C["ink_faint"],  "A MAPEAR",    C["ink_faint"]),
            "processing":    (C["surface"],  C["ok"],       "…",  C["ok"],         "CONSULTANDO", C["ok"]),
            "ok":            (C["surface"],  C["ok"],       "✓",  C["ok"],         "LIMITE OK",   C["ok"]),
            "warn":          (C["surface"],  C["warn"],     "!",  C["warn"],        "ATENÇÃO",     C["warn"]),
            "error":         (C["surface"],  C["err"],      "✗",  C["err"],         "ERRO",        C["err"]),
            "ltc_expired":   (C["surface"],  C["err"],      "✗",  C["err"],         "LTC VENCIDO", C["err"]),
        }.get(state, (C["surface"], C["hair"], "–", C["ink_faint"], state.upper(), C["ink_faint"]))

        bg, bord, icon_t, icon_fg, status_t, status_fg = cfg
        c["card"].configure(bg=bg, highlightbackground=bord)
        c["top_bar"].configure(bg=bord)
        c["body"].configure(bg=bg)
        c["icon_lbl"].configure(bg=bg, text=icon_t, fg=icon_fg)
        c["info_lbl"].configure(bg=bg, text=info, fg=C["ink_muted"])
        c["status_lbl"].configure(bg=bg, text=status_t, fg=status_fg)
        c["status_row"].configure(bg=bg)
        for w in c["body"].winfo_children():
            try: w.configure(bg=bg)
            except: pass

        if state == "processing":
            c["spin_id"][0] = True
            c["tick"]()

    def _publish_limite(self, name, ltc_str=None, ltc_date=None, limite_disp=None,
                        state="processing", info=""):
        cnpj = only_digits(BPM_CLIENT_DATA.get(name, {}).get("CNPJ", ""))
        if not cnpj:
            return
        data = {
            "client_name": name,
            "cnpj": cnpj,
            "ltc_str": ltc_str,
            "ltc_date": ltc_date,
            "limite_disp": limite_disp,
            "state": state,
            "info": info,
        }
        self.controller.publish_limite_result(cnpj, data)
        for mn in LIMITE_SHARED_RESULTS.get(name, []):
            mcnpj = only_digits(BPM_CLIENT_DATA.get(mn, {}).get("CNPJ", ""))
            if not mcnpj:
                continue
            mdata = dict(data)
            mdata["client_name"] = mn
            mdata["via"] = name
            self.controller.publish_limite_result(mcnpj, mdata)

    def _start_worker(self):
        if self._worker_running: return
        self._worker_running = True
        threading.Thread(target=self._worker, daemon=True).start()

    def _find_idx(self, name):
        for i, c in enumerate(self._cards):
            if c["name"] == name: return i
        return -1

    def _worker(self):
        if not PLAYWRIGHT_OK:
            self._ui(lambda: messagebox.showerror("Erro","Playwright não disponível."))
            self._worker_running = False; return
        today = date.today()
        with sync_playwright() as p:
            browser = None
            try:
                for ch in ("chrome","msedge"):
                    try: browser = p.chromium.launch(channel=ch, headless=False); break
                    except: pass
                if browser is None:
                    try: browser = p.chromium.launch(headless=False)
                    except Exception as e: raise RuntimeError(f"Não foi possível iniciar navegador: {e}")
                self._browser = browser
                page = browser.new_page()
                page.set_default_timeout(60_000)
                RE_DATE = re.compile(r"\d{2}/\d{2}/\d{4}")

                for name in list(BPM_CLIENT_DATA.keys()):
                    if self._cancel_requested: break
                    if name not in MAPPED_CLIENTS: continue
                    idx = self._find_idx(name)
                    if idx == -1: continue
                    self._publish_limite(name, state="processing")
                    self._ui(lambda i=idx: self._set_state(i, "processing"))
                    url = LIMITE_CLIENT_URLS