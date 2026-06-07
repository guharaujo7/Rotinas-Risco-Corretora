"""
Mesa Itaú — Risco Sacado
Redesign editorial dark/warm (Tkinter nativo).
Mantém toda a lógica de automação BPM, Share, Limites Invertido
e adiciona o sistema de Rotinas com seleção de módulo.
"""

import os, sys, re, random, tkinter as tk, threading, time, tempfile, shutil, webbrowser
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

# ─────────────────────────────────────────────
#  PALETA  dark / warm-ink
# ─────────────────────────────────────────────
C = {
    "bg":           "#1a1712",   # carvão quente
    "surface":      "#211e1a",   # superfície primária
    "surface2":     "#28241f",   # superfície elevada
    "surface3":     "#312d27",   # card hover / selecionado
    "ink":          "#ede8df",   # texto principal
    "ink_muted":    "#8a8070",   # texto secundário
    "ink_faint":    "#4a4540",   # hairline / divisor
    "accent":       "#c8923a",   # âmbar
    "accent_dim":   "#7a5520",   # âmbar apagado
    "accent_soft":  "#2d2015",   # âmbar fundo suave
    "ok":           "#4ea87a",   # verde sucesso
    "ok_dim":       "#1e3d2e",
    "warn":         "#c8923a",
    "err":          "#b85050",
    "err_dim":      "#3d1a1a",
    "hair":         "#2e2a25",   # hairline
    "log_step":     "#6a6258",
    "log_ok":       "#4ea87a",
    "log_warn":     "#c8923a",
    "log_err":      "#b85050",
}

LOGO_FILENAME   = "itau-logo-png_seeklogo-74122.png"

# ─────────────────────────────────────────────
#  HELPERS DE FORMATAÇÃO MONETÁRIA
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
#  DADOS DOS CLIENTES
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
#  REGEX EXTRAÇÃO PDF
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
#  FUNÇÕES UTILITÁRIAS
# ─────────────────────────────────────────────
def app_base_dir():
    return os.path.dirname(sys.executable) if getattr(sys,"frozen",False) else os.path.dirname(os.path.abspath(__file__))

def resource_path(p):
    return os.path.join(getattr(sys,"_MEIPASS",app_base_dir()), p)

def only_digits(s):
    return re.sub(r"\D","",s or "")

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

class BPMUserCancelled(Exception): pass

# ═══════════════════════════════════════════════════════════════════════════════
#  COMPONENTES VISUAIS REUTILIZÁVEIS
# ═══════════════════════════════════════════════════════════════════════════════

def make_hairline(parent, orient="h", **kwargs):
    kw = {"bg": C["hair"]}
    kw.update(kwargs)
    if orient == "h":
        return tk.Frame(parent, height=1, **kw)
    else:
        return tk.Frame(parent, width=1, **kw)


def styled_label(parent, text, size=10, weight="normal", color=None, **kwargs):
    return tk.Label(parent, text=text,
                    font=("Georgia", size, weight) if weight == "bold" or size >= 14 else ("Segoe UI", size, weight),
                    fg=color or C["ink"],
                    bg=kwargs.pop("bg", C["surface"]),
                    **kwargs)


def styled_button(parent, text, command, accent=False, danger=False, small=False, **kwargs):
    bg   = C["accent_dim"] if accent else (C["err_dim"] if danger else C["surface3"])
    fg   = C["accent"] if accent else (C["err"] if danger else C["ink"])
    abg  = C["accent"] if accent else (C["err"] if danger else C["surface2"])
    afg  = C["bg"] if accent else C["ink"]
    pad  = (8, 4) if small else (14, 7)
    btn  = tk.Button(parent, text=text, command=command,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=afg,
                     font=("Segoe UI", 8 if small else 9),
                     relief="flat", bd=0, padx=pad[0], pady=pad[1],
                     cursor="hand2", **kwargs)
    btn.bind("<Enter>", lambda _: btn.configure(bg=abg, fg=afg))
    btn.bind("<Leave>", lambda _: btn.configure(bg=bg, fg=fg))
    return btn


def styled_entry(parent, textvariable=None, width=20, show=None, **kwargs):
    e = tk.Entry(parent, textvariable=textvariable, width=width, show=show or "",
                 bg=C["bg"], fg=C["ink"], insertbackground=C["accent"],
                 relief="flat", highlightthickness=1,
                 highlightbackground=C["hair"], highlightcolor=C["accent"],
                 font=("Segoe UI", 10), **kwargs)
    return e


def card_frame(parent, **kwargs):
    kw = {"bg": C["surface"], "highlightthickness": 1,
          "highlightbackground": C["hair"], "bd": 0}
    kw.update(kwargs)
    return tk.Frame(parent, **kw)


# ─────────────────────────────────────────────
#  SCROLLABLE FRAME
# ─────────────────────────────────────────────
class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg=None, **kwargs):
        bg = bg or C["bg"]
        super().__init__(parent, bg=bg, **kwargs)
        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self._vbar   = tk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vbar.set)
        self._vbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self.inner   = tk.Frame(self._canvas, bg=bg)
        self._win    = self._canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner)
        self._canvas.bind("<Configure>", self._on_canvas)
        self._canvas.bind_all("<MouseWheel>",  self._mw)
        self._canvas.bind_all("<Button-4>",    self._mw)
        self._canvas.bind_all("<Button-5>",    self._mw)

    def _on_inner(self, _): self._canvas.configure(scrollregion=self._canvas.bbox("all"))
    def _on_canvas(self, e): self._canvas.itemconfigure(self._win, width=e.width)
    def _mw(self, e):
        if getattr(e,"delta",0): self._canvas.yview_scroll(int(-e.delta/120),"units")
        elif e.num==4: self._canvas.yview_scroll(-3,"units")
        elif e.num==5: self._canvas.yview_scroll(3,"units")


# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR  (navegação permanente à esquerda)
# ═══════════════════════════════════════════════════════════════════════════════
class Sidebar(tk.Frame):
    NAV = [
        ("Home",             "⌂",  "Início"),
        ("Rotinas",          "◈",  "Rotinas"),
        ("Share",            "⊕",  "Cadastro Share"),
        ("BPM",              "⚡",  "BPM"),
        ("LimitesInvertido", "⬡",  "Limites Invertido"),
    ]

    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=C["surface"], width=200, **kwargs)
        self.pack_propagate(False)
        self.controller = controller
        self._btns = {}
        self._build()

    def _build(self):
        # Logo / título
        top = tk.Frame(self, bg=C["surface"])
        top.pack(fill="x", padx=16, pady=(22, 0))

        logo_row = tk.Frame(top, bg=C["surface"])
        logo_row.pack(fill="x")
        tk.Label(logo_row, text="Mesa", bg=C["surface"], fg=C["ink"],
                 font=("Georgia", 14, "bold")).pack(side="left")
        tk.Label(logo_row, text=" Itaú", bg=C["surface"], fg=C["accent"],
                 font=("Georgia", 14, "bold")).pack(side="left")
        tk.Label(top, text="Risco Sacado", bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(2,0))

        make_hairline(self).pack(fill="x", padx=0, pady=(16, 12))

        # Nav links
        nav_frame = tk.Frame(self, bg=C["surface"])
        nav_frame.pack(fill="x", padx=10)

        for name, icon, label in self.NAV:
            btn_frame = tk.Frame(nav_frame, bg=C["surface"], cursor="hand2")
            btn_frame.pack(fill="x", pady=1)
            icon_lbl = tk.Label(btn_frame, text=icon, bg=C["surface"], fg=C["ink_muted"],
                                font=("Segoe UI", 11), width=3, anchor="e")
            icon_lbl.pack(side="left")
            text_lbl = tk.Label(btn_frame, text=label, bg=C["surface"], fg=C["ink_muted"],
                                font=("Segoe UI", 9), anchor="w")
            text_lbl.pack(side="left", padx=(6,0))
            accent_bar = tk.Frame(btn_frame, bg=C["surface"], width=3)
            accent_bar.pack(side="right", fill="y")

            def on_click(n=name): self.controller.show_frame(n)
            def on_enter(e, f=btn_frame, il=icon_lbl, tl=text_lbl):
                f.configure(bg=C["surface2"]); il.configure(bg=C["surface2"]); tl.configure(bg=C["surface2"])
            def on_leave(e, f=btn_frame, il=icon_lbl, tl=text_lbl, n=name):
                active = getattr(self.controller, "_active_frame", None)
                if active != n:
                    f.configure(bg=C["surface"]); il.configure(bg=C["surface"]); tl.configure(bg=C["surface"])

            for w in (btn_frame, icon_lbl, text_lbl):
                w.bind("<Button-1>", lambda _,n=name: on_click(n))
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)

            self._btns[name] = {
                "frame": btn_frame, "icon": icon_lbl,
                "text": text_lbl, "bar": accent_bar
            }

        # Data/hora rodapé
        self._clock_lbl = tk.Label(self, text="", bg=C["surface"], fg=C["ink_faint"],
                                   font=("Segoe UI", 8))
        self._clock_lbl.pack(side="bottom", pady=10)
        self._tick_clock()

    def _tick_clock(self):
        now = datetime.now()
        self._clock_lbl.configure(text=now.strftime("%d/%m  %H:%M"))
        self.after(30_000, self._tick_clock)

    def set_active(self, name):
        for n, w in self._btns.items():
            if n == name:
                w["frame"].configure(bg=C["surface2"])
                w["icon"].configure(bg=C["surface2"], fg=C["accent"])
                w["text"].configure(bg=C["surface2"], fg=C["ink"])
                w["bar"].configure(bg=C["accent"])
            else:
                w["frame"].configure(bg=C["surface"])
                w["icon"].configure(bg=C["surface"], fg=C["ink_muted"])
                w["text"].configure(bg=C["surface"], fg=C["ink_muted"])
                w["bar"].configure(bg=C["surface"])


# ═══════════════════════════════════════════════════════════════════════════════
#  HOME FRAME — dashboard
# ═══════════════════════════════════════════════════════════════════════════════
class HomeFrame(tk.Frame):
    MODULES = [
        {
            "name":  "Cadastro Share",
            "sub":   "Extração e análise de PDF",
            "icon":  "⊕",
            "frame": "Share",
            "accent": False,
        },
        {
            "name":  "BPM",
            "sub":   "Abertura de solicitações",
            "icon":  "⚡",
            "frame": "BPM_CONFIG",
            "accent": False,
        },
        {
            "name":  "Limites Invertido",
            "sub":   "Consulta LTC e limites disponíveis",
            "icon":  "⬡",
            "frame": "LimitesInvertido",
            "accent": True,
        },
        {
            "name":  "Rotinas",
            "sub":   "Sequências configuráveis",
            "icon":  "◈",
            "frame": "Rotinas",
            "accent": False,
        },
    ]

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._build()

    def _build(self):
        sf = ScrollableFrame(self)
        sf.pack(fill="both", expand=True)
        inner = sf.inner
        inner.configure(bg=C["bg"])
        inner.columnconfigure(0, weight=1)

        # ── Saudação ──────────────────────────────
        greet_frame = tk.Frame(inner, bg=C["bg"])
        greet_frame.pack(fill="x", padx=40, pady=(36, 0))
        now = datetime.now()
        hour = now.hour
        saudacao = "Bom dia" if hour < 12 else ("Boa tarde" if hour < 18 else "Boa noite")
        tk.Label(greet_frame, text=f"{saudacao}.", bg=C["bg"], fg=C["ink"],
                 font=("Georgia", 22, "bold"), anchor="w").pack(anchor="w")
        tk.Label(greet_frame, text=now.strftime("%A, %d de %B de %Y").capitalize(),
                 bg=C["bg"], fg=C["ink_muted"], font=("Segoe UI", 10)).pack(anchor="w", pady=(4,0))

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=40, pady=(24, 28))

        # ── Módulos grid ─────────────────────────
        tk.Label(inner, text="MÓDULOS", bg=C["bg"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=40, pady=(0,12))

        grid_frame = tk.Frame(inner, bg=C["bg"])
        grid_frame.pack(fill="x", padx=40)
        grid_frame.columnconfigure(0, weight=1, uniform="mod")
        grid_frame.columnconfigure(1, weight=1, uniform="mod")

        for i, mod in enumerate(self.MODULES):
            row, col = divmod(i, 2)
            self._make_module_card(grid_frame, mod, row, col)

        make_hairline(inner, bg=C["hair"]).pack(fill="x", padx=40, pady=(32, 24))

        # ── Atalhos rápidos ───────────────────────
        tk.Label(inner, text="ATALHOS RÁPIDOS", bg=C["bg"], fg=C["ink_faint"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=40, pady=(0,10))

        quick = tk.Frame(inner, bg=C["bg"])
        quick.pack(fill="x", padx=40, pady=(0, 40))
        for label, frame in [
            ("→  Nova solicitação BPM",   "BPM_CONFIG"),
            ("→  Consultar limites",       "LimitesInvertido"),
            ("→  Extrair dados de PDF",    "Share"),
            ("→  Gerenciar rotinas",       "Rotinas"),
        ]:
            btn = tk.Button(quick, text=label, command=lambda f=frame: self.controller.show_frame(f),
                            bg=C["bg"], fg=C["ink_muted"], activebackground=C["surface"],
                            activeforeground=C["ink"], font=("Segoe UI", 9),
                            relief="flat", bd=0, anchor="w", padx=0, pady=5, cursor="hand2")
            btn.pack(anchor="w")
            btn.bind("<Enter>", lambda e, b=btn: b.configure(fg=C["accent"]))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(fg=C["ink_muted"]))

        # rodapé
        tk.Label(inner, text="Mesa de Operação  ·  Risco Sacado  ·  Middle",
                 bg=C["bg"], fg=C["ink_faint"], font=("Segoe UI", 7)).pack(pady=(0, 20))

    def _make_module_card(self, parent, mod, row, col):
        is_accent = mod["accent"]
        bg   = C["surface"]
        bord = C["accent_dim"] if is_accent else C["hair"]
        fg_name = C["accent"] if is_accent else C["ink"]
        pad  = (0, 6) if col == 0 else (6, 0)

        outer = tk.Frame(parent, bg=bg, highlightthickness=1, highlightbackground=bord,
                         cursor="hand2")
        outer.grid(row=row, column=col, sticky="nsew", padx=pad, pady=6)

        top_bar = tk.Frame(outer, bg=C["accent"] if is_accent else C["hair"], height=2)
        top_bar.pack(fill="x")

        body = tk.Frame(outer, bg=bg, padx=18, pady=16)
        body.pack(fill="both", expand=True)

        head = tk.Frame(body, bg=bg)
        head.pack(fill="x")
        tk.Label(head, text=mod["icon"], bg=bg, fg=C["accent"], font=("Segoe UI",16)).pack(side="left")
        tk.Label(body, text=mod["name"], bg=bg, fg=fg_name,
                 font=("Segoe UI", 11, "bold"), anchor="w").pack(anchor="w", pady=(8,2))
        tk.Label(body, text=mod["sub"], bg=bg, fg=C["ink_muted"],
                 font=("Segoe UI", 8), anchor="w", wraplength=160, justify="left").pack(anchor="w")

        def cmd(f=mod["frame"]): self.controller.show_frame(f)
        def enter(e, o=outer, b=body, bg_h=C["surface2"]):
            o.configure(bg=bg_h, highlightbackground=C["accent"])
            b.configure(bg=bg_h)
            for w in b.winfo_children():
                try: w.configure(bg=bg_h)
                except: pass
            for w in head.winfo_children():
                try: w.configure(bg=bg_h)
                except: pass
        def leave(e, o=outer, b=body):
            o.configure(bg=bg, highlightbackground=bord)
            b.configure(bg=bg)
            for w in b.winfo_children():
                try: w.configure(bg=bg)
                except: pass
            for w in head.winfo_children():
                try: w.configure(bg=bg)
                except: pass

        for w in [outer, body, head] + list(body.winfo_children()) + list(head.winfo_children()):
            try:
                w.bind("<Button-1>", lambda _: cmd())
                w.bind("<Enter>", enter)
                w.bind("<Leave>", leave)
            except: pass


# ═══════════════════════════════════════════════════════════════════════════════
#  ROTINAS FRAME — gestão de rotinas configuráveis
# ═══════════════════════════════════════════════════════════════════════════════

# Rotinas pré-definidas com seus módulos
ROTINAS_PREDEFINIDAS = [
    {
        "nome": "Abertura de Solicitações (Invertido)",
        "icon": "⚡",
        "descricao": "Sequência completa: verificar limites → abrir BPM para clientes selecionados.",
        "passos": [
            {"nome": "Consultar limites (Limites Invertido)", "modulo": "LimitesInvertido", "obrigatorio": True},
            {"nome": "Configurar e executar BPM",              "modulo": "BPM_CONFIG",       "obrigatorio": True},
        ],
    },
    {
        "nome": "Cadastro e Revisão Share",
        "icon": "⊕",
        "descricao": "Extrai dados do PDF e gera resumo para cadastro.",
        "passos": [
            {"nome": "Extrair PDF (Cadastro Share)", "modulo": "Share",  "obrigatorio": True},
        ],
    },
    {
        "nome": "Rotina Completa do Dia",
        "icon": "◈",
        "descricao": "Limites → Share → BPM: fluxo completo de operações.",
        "passos": [
            {"nome": "Verificar limites",    "modulo": "LimitesInvertido", "obrigatorio": True},
            {"nome": "Revisar cadastros Share", "modulo": "Share",         "obrigatorio": False},
            {"nome": "Executar BPM",         "modulo": "BPM_CONFIG",       "obrigatorio": True},
        ],
    },
]


class RotinasFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._rotinas   = list(ROTINAS_PREDEFINIDAS)  # cópia mutável
        self._custom    = []   # rotinas criadas pelo usuário
        self._exec_idx  = None
        self._exec_step = 0
        self._build()

    # ─── build ──────────────────────────────────────────────────
    def _build(self):
        # header
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24, 0))
        tk.Label(hdr, text="Rotinas", bg=C["bg"], fg=C["ink"],
                 font=("Georgia", 18, "bold")).pack(side="left")
        styled_button(hdr, "+ Nova rotina", self._open_new_rotina_dialog,
                      accent=True).pack(side="right")

        tk.Label(hdr, text="Sequências configuráveis de módulos", bg=C["bg"],
                 fg=C["ink_muted"], font=("Segoe UI", 9)).pack(side="left", padx=(12,0))

        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(18,0))

        # lista de rotinas
        self._list_frame = ScrollableFrame(self, bg=C["bg"])
        self._list_frame.pack(fill="both", expand=True)
        self._refresh_list()

    def _refresh_list(self):
        for w in self._list_frame.inner.winfo_children():
            w.destroy()
        inner = self._list_frame.inner
        inner.configure(bg=C["bg"])

        all_rotinas = self._rotinas + self._custom
        if not all_rotinas:
            tk.Label(inner, text="Nenhuma rotina cadastrada.", bg=C["bg"],
                     fg=C["ink_muted"], font=("Segoe UI",10)).pack(pady=40)
            return

        for i, rot in enumerate(all_rotinas):
            self._make_rotina_card(inner, rot, i)

    def _make_rotina_card(self, parent, rot, idx):
        is_custom = idx >= len(self._rotinas)
        card = card_frame(parent)
        card.pack(fill="x", padx=32, pady=(12 if idx==0 else 6, 0))

        # cabeçalho
        head = tk.Frame(card, bg=C["surface"], padx=18, pady=14)
        head.pack(fill="x")

        tk.Label(head, text=rot["icon"], bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI",15)).pack(side="left")
        title_col = tk.Frame(head, bg=C["surface"])
        title_col.pack(side="left", padx=(10,0), fill="x", expand=True)
        tk.Label(title_col, text=rot["nome"], bg=C["surface"], fg=C["ink"],
                 font=("Segoe UI", 11, "bold"), anchor="w").pack(anchor="w")
        tk.Label(title_col, text=rot.get("descricao",""), bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI",8), anchor="w",
                 wraplength=380, justify="left").pack(anchor="w", pady=(2,0))

        # passos
        steps_frame = tk.Frame(card, bg=C["surface"], padx=18, pady=(0,10))
        steps_frame.pack(fill="x")
        for j, passo in enumerate(rot.get("passos",[])):
            row = tk.Frame(steps_frame, bg=C["surface"])
            row.pack(fill="x", pady=2)
            dot_color = C["accent"] if passo.get("obrigatorio") else C["ink_faint"]
            tk.Label(row, text="●", bg=C["surface"], fg=dot_color,
                     font=("Segoe UI",7)).pack(side="left")
            tk.Label(row, text=passo["nome"], bg=C["surface"], fg=C["ink_muted"],
                     font=("Segoe UI",8)).pack(side="left", padx=(6,0))
            mod_tag = passo.get("modulo","")
            if mod_tag:
                tk.Label(row, text=f"[{mod_tag}]", bg=C["surface"], fg=C["ink_faint"],
                         font=("Segoe UI",7)).pack(side="left", padx=(4,0))

        make_hairline(card, bg=C["hair"]).pack(fill="x", padx=0)

        # footer ações
        foot = tk.Frame(card, bg=C["surface"], padx=18, pady=10)
        foot.pack(fill="x")
        styled_button(foot, "▶  Executar rotina",
                      lambda r=rot: self._execute_rotina(r),
                      accent=True, small=True).pack(side="left")
        styled_button(foot, "Editar",
                      lambda r=rot, i=idx: self._open_edit_dialog(r, i),
                      small=True).pack(side="left", padx=(6,0))
        if is_custom:
            styled_button(foot, "Remover",
                          lambda i=idx: self._remove_rotina(i),
                          danger=True, small=True).pack(side="left", padx=(6,0))

    def _execute_rotina(self, rot):
        """Inicia a rotina: navega pelos módulos na sequência."""
        passos = rot.get("passos", [])
        if not passos:
            messagebox.showinfo("Rotina vazia", "Esta rotina não tem passos configurados.")
            return
        primeiro = passos[0].get("modulo","")
        if primeiro:
            # Guarda contexto de execução no controller
            self.controller.rotina_em_execucao = {
                "nome": rot["nome"],
                "passos": passos,
                "step_atual": 0,
            }
            messagebox.showinfo(
                "Iniciar Rotina",
                f"Rotina: {rot['nome']}\n\nPasso 1 de {len(passos)}: {passos[0]['nome']}\n\nClique OK para iniciar.",
                parent=self
            )
            self.controller.show_frame(primeiro)

    def _remove_rotina(self, idx):
        ci = idx - len(self._rotinas)
        if ci >= 0 and ci < len(self._custom):
            nome = self._custom[ci]["nome"]
            if messagebox.askyesno("Remover rotina", f"Remover '{nome}'?"):
                self._custom.pop(ci)
                self._refresh_list()

    # ─── diálogo: nova / editar ──────────────────────────────────
    def _open_new_rotina_dialog(self):
        self._open_rotina_dialog(None, None)

    def _open_edit_dialog(self, rot, idx):
        self._open_rotina_dialog(rot, idx)

    def _open_rotina_dialog(self, rot, idx):
        dlg = tk.Toplevel(self)
        dlg.title("Nova Rotina" if rot is None else "Editar Rotina")
        dlg.configure(bg=C["surface"])
        dlg.geometry("520x560")
        dlg.resizable(False, True)
        dlg.grab_set()

        tk.Label(dlg, text="Nova Rotina" if rot is None else "Editar Rotina",
                 bg=C["surface"], fg=C["ink"],
                 font=("Georgia", 14, "bold")).pack(anchor="w", padx=24, pady=(20,4))
        make_hairline(dlg, bg=C["hair"]).pack(fill="x", padx=0, pady=(0,16))

        form = tk.Frame(dlg, bg=C["surface"])
        form.pack(fill="both", expand=True, padx=24)

        # Nome
        tk.Label(form, text="Nome", bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI",8)).pack(anchor="w")
        nome_var = tk.StringVar(value=(rot or {}).get("nome",""))
        styled_entry(form, textvariable=nome_var).pack(fill="x", pady=(4,12))

        # Ícone
        tk.Label(form, text="Ícone (emoji)", bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI",8)).pack(anchor="w")
        icon_var = tk.StringVar(value=(rot or {}).get("icon","◈"))
        styled_entry(form, textvariable=icon_var, width=5).pack(anchor="w", pady=(4,12))

        # Descrição
        tk.Label(form, text="Descrição", bg=C["surface"], fg=C["ink_muted"],
                 font=("Segoe UI",8)).pack(anchor="w")
        desc_var = tk.StringVar(value=(rot or {}).get("descricao",""))
        styled_entry(form, textvariable=desc_var).pack(fill="x", pady=(4,12))

        # Passos
        tk.Label(form, text="Passos (módulos em sequência)", bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w")

        MODULOS_DISP = [
            ("LimitesInvertido", "Limites Invertido"),
            ("Share",            "Cadastro Share"),
            ("BPM_CONFIG",       "BPM"),
        ]

        passos_vars = []  # lista de (nome_var, modulo_var, obr_var)

        passos_frame = tk.Frame(form, bg=C["surface"])
        passos_frame.pack(fill="x", pady=(4,0))

        def refresh_passos():
            for w in passos_frame.winfo_children(): w.destroy()
            for pi, (nv, mv, ov) in enumerate(passos_vars):
                pr = tk.Frame(passos_frame, bg=C["surface2"], pady=6, padx=8)
                pr.pack(fill="x", pady=2)
                tk.Label(pr, text=f"{pi+1}.", bg=C["surface2"], fg=C["ink_muted"],
                         font=("Segoe UI",9), width=2).pack(side="left")
                styled_entry(pr, textvariable=nv, width=16).pack(side="left", padx=(4,4))
                # dropdown módulo
                opts = [m[1] for m in MODULOS_DISP]
                combo = ttk.Combobox(pr, textvariable=mv, values=opts, width=14, state="readonly")
                # mapeia label → key
                cur_key = mv.get()
                for k, lbl in MODULOS_DISP:
                    if k == cur_key: combo.set(lbl); break
                else: combo.set(opts[0])
                def on_combo_change(e, mv2=mv, c=combo):
                    lbl = c.get()
                    for k, l in MODULOS_DISP:
                        if l == lbl: mv2.set(k); break
                combo.bind("<<ComboboxSelected>>", on_combo_change)
                combo.pack(side="left", padx=(0,4))
                ck = tk.Checkbutton(pr, text="Obrig.", variable=ov,
                                    bg=C["surface2"], fg=C["ink_muted"],
                                    selectcolor=C["bg"], activebackground=C["surface2"],
                                    font=("Segoe UI",7))
                ck.pack(side="left")
                styled_button(pr, "✕", lambda p=pi: remove_passo(p), danger=True, small=True).pack(side="right")

        def add_passo():
            passos_vars.append((tk.StringVar(value="Novo passo"),
                                tk.StringVar(value="LimitesInvertido"),
                                tk.BooleanVar(value=True)))
            refresh_passos()

        def remove_passo(pi):
            if pi < len(passos_vars): passos_vars.pop(pi)
            refresh_passos()

        # preenche passos existentes
        for p in (rot or {}).get("passos",[]):
            passos_vars.append((tk.StringVar(value=p.get("nome","Passo")),
                                tk.StringVar(value=p.get("modulo","LimitesInvertido")),
                                tk.BooleanVar(value=p.get("obrigatorio",True))))
        refresh_passos()

        styled_button(form, "+ Adicionar passo", add_passo, small=True).pack(anchor="w", pady=(6,0))

        # footer
        make_hairline(dlg, bg=C["hair"]).pack(fill="x", padx=0, pady=(12,0))
        foot = tk.Frame(dlg, bg=C["surface"])
        foot.pack(fill="x", padx=24, pady=12)

        def salvar():
            nome = nome_var.get().strip()
            if not nome:
                messagebox.showwarning("Campo obrigatório", "Informe um nome para a rotina.", parent=dlg)
                return
            nova = {
                "nome": nome,
                "icon": icon_var.get().strip() or "◈",
                "descricao": desc_var.get().strip(),
                "passos": [
                    {"nome": nv.get(), "modulo": mv.get(), "obrigatorio": bool(ov.get())}
                    for nv, mv, ov in passos_vars
                ],
            }
            if rot is None or idx is None or idx >= len(self._rotinas):
                self._custom.append(nova)
            else:
                # edita predefinida → vira custom substituindo na lista
                self._rotinas[idx] = nova
            dlg.destroy()
            self._refresh_list()

        styled_button(foot, "Cancelar", dlg.destroy, small=True).pack(side="right", padx=(6,0))
        styled_button(foot, "Salvar", salvar, accent=True, small=True).pack(side="right")


# ═══════════════════════════════════════════════════════════════════════════════
#  BPM CONFIG FRAME — seleção de clientes + credenciais
# ═══════════════════════════════════════════════════════════════════════════════
class BPMConfigFrame(tk.Frame):
    CLIENTS = list(BPM_CLIENT_DATA.keys())

    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._selected  = {}  # nome → StringVar(valor)
        self._func_var  = tk.StringVar()
        self._senha_var = tk.StringVar()
        self._build()

    def _only_digits(self, s):
        return "".join(c for c in (s or "") if c.isdigit())

    def on_show(self):
        # Preenche credenciais já salvas
        self._func_var.set(getattr(self.controller, "bpm_funcional", "") or "")
        self._senha_var.set(getattr(self.controller, "bpm_password", "") or "")

    def _build(self):
        # header
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        tk.Label(hdr, text="Configurar BPM", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left")
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(14,0))

        sf = ScrollableFrame(self, bg=C["bg"])
        sf.pack(fill="both", expand=True)
        body = sf.inner
        body.configure(bg=C["bg"])

        # ── Credenciais ─────────────────────────────────────────
        sec = card_frame(body)
        sec.pack(fill="x", padx=32, pady=(20,0))
        tk.Label(sec, text="Credenciais do Painel de Serviços", bg=C["surface"],
                 fg=C["ink"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=18, pady=(14,0))
        tk.Label(sec, text="Funcional e senha para login no Painel BPM",
                 bg=C["surface"], fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w", padx=18, pady=(2,10))
        make_hairline(sec, bg=C["hair"]).pack(fill="x")

        cred_form = tk.Frame(sec, bg=C["surface"], padx=18, pady=14)
        cred_form.pack(fill="x")
        cred_form.columnconfigure(0, weight=1)
        cred_form.columnconfigure(1, weight=1)

        # Funcional
        fc = tk.Frame(cred_form, bg=C["surface"])
        fc.grid(row=0, column=0, sticky="ew", padx=(0,8))
        tk.Label(fc, text="Funcional (somente números)", bg=C["surface"],
                 fg=C["ink_muted"], font=("Segoe UI",8)).pack(anchor="w")
        self._ent_func = styled_entry(fc, textvariable=self._func_var)
        self._ent_func.pack(fill="x", pady=(4,0))

        # Senha
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

        # ── Seleção de clientes ──────────────────────────────────
        sec2 = card_frame(body)
        sec2.pack(fill="x", padx=32, pady=(16,0))
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

        make_hairline(body, bg=C["hair"]).pack(fill="x", padx=32, pady=(20,0))

        # ── Footer ──────────────────────────────────────────────
        foot = tk.Frame(body, bg=C["bg"])
        foot.pack(fill="x", padx=32, pady=16)
        styled_button(foot, "← Voltar",
                      lambda: self.controller.show_frame("Home")).pack(side="left")
        self._run_btn = styled_button(foot, "▶  Iniciar BPM",
                                      self._start_bpm, accent=True)
        self._run_btn.pack(side="right")

    def _make_client_row(self, parent, cli):
        row = tk.Frame(parent, bg=C["surface"], pady=3)
        row.pack(fill="x")
        row.columnconfigure(1, weight=1)

        selected = tk.BooleanVar(value=False)
        val_var  = tk.StringVar(value="R$ 0,00")
        val_digits = [""]   # mutável via lista

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

    def _start_bpm(self):
        func = self._only_digits(self._func_var.get())
        senha = self._only_digits(self._senha_var.get())
        if not func:
            messagebox.showwarning("Funcional obrigatório", "Informe o funcional (somente números)."); return
        if not (1 <= len(senha) <= 6):
            messagebox.showwarning("Senha inválida", "Senha deve ter 1 a 6 dígitos."); return

        selection = []
        for cli in self.CLIENTS:
            row = self._client_rows[cli]
            if not row["selected"].get(): continue
            raw = row["val_var"].get()
            d = _parse_brl(raw)
            if d is None or d == Decimal("0.00"):
                messagebox.showwarning("Valor inválido", f"Informe um valor válido para {cli}."); return
            selection.append({"cliente": cli, "valor": _fmt_brl_from_raw(raw)})

        if not selection:
            messagebox.showwarning("Seleção vazia", "Selecione ao menos um cliente."); return

        self.controller.bpm_funcional    = func
        self.controller.bpm_password     = senha
        self.controller.bpm_run_selection = selection
        self.controller.show_frame("BPM")


# ═══════════════════════════════════════════════════════════════════════════════
#  SHARE FRAME — extração PDF + resumo
# ═══════════════════════════════════════════════════════════════════════════════
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
        # header
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        tk.Label(hdr, text="Cadastro Share", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left")
        styled_button(hdr, "📄  Abrir PDF…", self._on_open_pdf, accent=True).pack(side="right")
        styled_button(hdr, "← Voltar",
                      lambda: self.controller.show_frame("Home")).pack(side="right", padx=(0,6))
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(14,0))

        sf = ScrollableFrame(self, bg=C["bg"])
        sf.pack(fill="both", expand=True)
        body = sf.inner
        body.configure(bg=C["bg"])

        # grid campos
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

        # ações
        act = tk.Frame(body, bg=C["bg"])
        act.pack(fill="x", padx=32, pady=(14,0))
        styled_button(act,"🔄  Gerar Resumo",  self._update_resumo, accent=True).pack(side="left")
        styled_button(act,"🧽  Limpar",          self._clear_all).pack(side="left", padx=(6,0))

        # resumo
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
        sc = tk.Scrollbar(txt_wrap, orient="vertical", command=self.txt_resumo.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.txt_resumo.configure(yscrollcommand=sc.set)

        foot = tk.Frame(body, bg=C["bg"])
        foot.pack(fill="x", padx=32, pady=(12,30))
        styled_button(foot,"📋  Copiar Resumo", self._copy_resumo, accent=True).pack(side="left")
        styled_button(foot,"💾  Salvar .txt",   self._save_resumo).pack(side="left", padx=(6,0))

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


# ═══════════════════════════════════════════════════════════════════════════════
#  LIMITES INVERTIDO FRAME
# ═══════════════════════════════════════════════════════════════════════════════
class LimitesInvertidoFrame(tk.Frame):
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
        self._build()

    def _ui(self, fn):
        try: self.after(0, fn)
        except: pass

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        styled_button(hdr, "← Voltar",
                      lambda: self.controller.show_frame("Home")).pack(side="left")
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
        self._grid_outer = self._sf.inner
        self._grid_outer.configure(bg=C["bg"])

        self._grid = tk.Frame(self._grid_outer, bg=C["bg"])
        self._grid.pack(padx=32, pady=(16,24), fill="x")
        for c in range(3):
            self._grid.columnconfigure(c, weight=1, uniform="lcards")

    def on_show(self):
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
        # cancel spinner
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
            c["spin_id"][0] = True  # flag
            c["tick"]()

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
                    self._ui(lambda i=idx: self._set_state(i, "processing"))
                    url = LIMITE_CLIENT_URLS.get(name)
                    if not url:
                        self._ui(lambda i=idx: self._set_state(i,"error","URL não mapeada")); continue
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
                        time.sleep(7)
                        if self._cancel_requested: break

                        # ler LTC
                        ltc_str = None
                        try:
                            all_spans = page.locator("span.u-font-size--14.u-ml--8.u-mt--8.u-block")
                            for si in range(all_spans.count()):
                                txt = (all_spans.nth(si).inner_text() or "").strip()
                                if RE_DATE.match(txt): ltc_str = txt; break
                        except: pass
                        if not ltc_str:
                            try:
                                body_text = page.locator("body").inner_text()
                                m = RE_DATE.search(body_text)
                                if m: ltc_str = m.group(0)
                            except: pass

                        try:
                            ltc_date = datetime.strptime(ltc_str,"%d/%m/%Y").date() if ltc_str else None
                        except: ltc_date = None

                        if ltc_date is None:
                            inf = "Não foi possível ler\ndata do LTC"
                            self._ui(lambda i=idx,inf=inf: self._set_state(i,"error",inf))
                            for mn in LIMITE_SHARED_RESULTS.get(name,[]):
                                mi = self._find_idx(mn)
                                if mi!=-1: self._ui(lambda i=mi,inf=inf: self._set_state(i,"error",inf))
                            continue

                        if ltc_date <= today:
                            inf = f"LTC vencido em {ltc_str}"
                            self._ui(lambda i=idx,inf=inf: self._set_state(i,"ltc_expired",inf))
                            for mn in LIMITE_SHARED_RESULTS.get(name,[]):
                                mi = self._find_idx(mn)
                                if mi!=-1: self._ui(lambda i=mi,inf=inf: self._set_state(i,"ltc_expired",inf))
                            continue

                        # ler limite
                        limite_disp = None
                        BAIXO = 1_000_000
                        try:
                            val_js = page.evaluate("""
                                (function(){
                                    var rows=document.querySelectorAll('table.atual tbody tr');
                                    for(var i=0;i<rows.length;i++){
                                        var n=rows[i].querySelector('td.tdNomeFinalidade');
                                        if(!n||String(n.textContent).trim().toLowerCase()!=='fornecedor') continue;
                                        var d=rows[i].querySelector('td[id^="valorDisponibilidade_"]');
                                        if(!d) continue;
                                        return String(d.textContent).trim();
                                    }
                                    return null;
                                })()
                            """)
                            if val_js:
                                s = val_js.replace(".","").replace(",","")
                                try: limite_disp = int(s)
                                except: pass
                            if limite_disp is None:
                                tds = page.locator("td.tdValorDisp")
                                for ti in range(tds.count()):
                                    s = re.sub(r"[^\d]","", tds.nth(ti).inner_text() or "")
                                    try:
                                        v = int(s)
                                        if limite_disp is None or v > limite_disp: limite_disp = v
                                    except: pass
                        except: pass

                        disp_fmt = f"R$ {limite_disp:,}".replace(",",".") if limite_disp is not None else "N/D"
                        inf = f"LTC ativo · vence {ltc_str}\nLimite Disp. (fornecedor): {disp_fmt}"
                        warn = limite_disp is not None and limite_disp < BAIXO
                        final_state = "warn" if warn else "ok"
                        if warn: inf += f"\n⚠ Limite abaixo de R$ {BAIXO:,}".replace(",",".")

                        self._ui(lambda i=idx,s=final_state,inf=inf: self._set_state(i,s,inf))
                        for mn in LIMITE_SHARED_RESULTS.get(name,[]):
                            mi = self._find_idx(mn)
                            if mi!=-1:
                                minf = inf + f"\n(via {name})"
                                self._ui(lambda i=mi,s=final_state,inf=minf: self._set_state(i,s,inf))

                    except BPMUserCancelled: break
                    except Exception as e:
                        if self._cancel_requested: break
                        es = str(e)[:80]
                        self._ui(lambda i=idx,es=es: self._set_state(i,"error",es))

            except Exception as e:
                if not self._cancel_requested:
                    em = str(e)
                    self._ui(lambda: messagebox.showerror("Erro na consulta",em))
            finally:
                self._browser = None
                if browser:
                    try: browser.close()
                    except: pass
                self._worker_running = False


# ═══════════════════════════════════════════════════════════════════════════════
#  BPM FRAME — execução com log scrollável
# ═══════════════════════════════════════════════════════════════════════════════
class BPMFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=C["bg"])
        self.controller = controller
        self._cards   = []
        self._selected= []
        self._worker_running   = False
        self._cancel_requested = False
        self._browser = None
        self._started = False
        self._build()

    def _ui(self, fn):
        try: self.after(0, fn)
        except: pass

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=32, pady=(24,0))
        tk.Label(hdr, text="BPM — Operações em andamento", bg=C["bg"], fg=C["ink"],
                 font=("Georgia",18,"bold")).pack(side="left")
        self._cancel_btn = styled_button(hdr, "✕  Cancelar",
                                         self._on_cancel, danger=True)
        self._cancel_btn.pack(side="right")
        make_hairline(self, bg=C["hair"]).pack(fill="x", padx=0, pady=(14,0))

        # grid de cards
        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._grid_wrap = self._sf.inner
        self._grid_wrap.configure(bg=C["bg"])

        self._grid = tk.Frame(self._grid_wrap, bg=C["bg"])
        self._grid.pack(padx=32, pady=(16,8), fill="x")
        for c in range(3):
            self._grid.columnconfigure(c, weight=1, uniform="bcards")

        # log
        make_hairline(self._grid_wrap, bg=C["hair"]).pack(fill="x", padx=32, pady=(8,0))
        log_hdr = tk.Frame(self._grid_wrap, bg=C["bg"])
        log_hdr.pack(fill="x", padx=32, pady=(10,4))
        tk.Label(log_hdr, text="Log de execução", bg=C["bg"], fg=C["ink_muted"],
                 font=("Segoe UI",8,"bold")).pack(side="left")
        styled_button(log_hdr, "Limpar log",
                      self._clear_log, small=True).pack(side="right")

        log_frame = tk.Frame(self._grid_wrap, bg=C["bg"])
        log_frame.pack(fill="x", padx=32, pady=(0,24))
        log_frame.columnconfigure(0, weight=1)
        self._log = tk.Text(log_frame, height=10, wrap="word", bd=0, relief="flat",
                            bg=C["surface"], fg=C["log_step"],
                            font=("Consolas",8), padx=10, pady=8,
                            state="disabled")
        self._log.grid(row=0, column=0, sticky="ew")
        lsb = tk.Scrollbar(log_frame, orient="vertical", command=self._log.yview)
        lsb.grid(row=0, column=1, sticky="ns")
        self._log.configure(yscrollcommand=lsb.set)
        self._log.tag_configure("step",    foreground=C["log_step"])
        self._log.tag_configure("heading", foreground=C["ink"])
        self._log.tag_configure("ok",      foreground=C["log_ok"])
        self._log.tag_configure("warn",    foreground=C["log_warn"])
        self._log.tag_configure("err",     foreground=C["log_err"])

    def _log_line(self, msg, tag="step"):
        def _do():
            self._log.configure(state="normal")
            self._log.insert("end", f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n", tag)
            self._log.see("end")
            self._log.configure(state="disabled")
        self._ui(_do)

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0","end")
        self._log.configure(state="disabled")

    def on_show(self):
        sel = getattr(self.controller,"bpm_run_selection",None) or []
        if not sel: return
        self._cancel_requested = False
        if self._started: return
        self._started = True
        self._selected = sel
        self._setup_cards()
        self._start_worker()

    def _on_cancel(self):
        self._cancel_requested = True
        self.controller.bpm_run_selection = []
        br = getattr(self,"_browser",None)
        if br:
            try: br.close()
            except: pass
        self._reset()
        self._started = False
        self.controller.show_frame("BPM_CONFIG")

    def _reset(self):
        self._cards = []
        for w in self._grid.winfo_children(): w.destroy()

    def _setup_cards(self):
        self._reset()
        for idx, item in enumerate(self._selected):
            cli = item.get("cliente","")
            raw = (item.get("valor","")).strip()
            display = raw[2:].strip() if raw.lower().startswith("r$") else raw
            row, col = divmod(idx,3)
            c = self._make_card(cli, display, row, col)
            self._cards.append(c)
        for i in range(1, len(self._cards)):
            self.after(60*i, lambda i=i: self._set_card_state(i,"waiting",show=True))
        if self._cards:
            self.after(0, lambda: self._set_card_state(0,"processing",show=True))

    def _make_card(self, name, val_display, row, col):
        bg = C["surface"]; bord = C["hair"]
        outer = tk.Frame(self._grid, bg=C["bg"])
        outer.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)
        outer.grid_remove()
        card = tk.Frame(outer, bg=bg, highlightthickness=1, highlightbackground=bord)
        card.pack(fill="both", expand=True)
        top_bar = tk.Frame(card, bg=bord, height=2)
        top_bar.pack(fill="x")
        body = tk.Frame(card, bg=bg, padx=14, pady=12)
        body.pack(fill="both", expand=True)
        icon_lbl = tk.Label(body, text="–", bg=bg, fg=C["ink_faint"],
                            font=("Segoe UI",14,"bold"))
        icon_lbl.pack()
        tk.Label(body, text=name, bg=bg, fg=C["ink"],
                 font=("Segoe UI",10,"bold")).pack(pady=(6,0))
        tk.Label(body, text=f"R$ {val_display}", bg=bg, fg=C["ink_muted"],
                 font=("Segoe UI",8)).pack(pady=(2,0))
        status_lbl = tk.Label(body, text="EM ESPERA", bg=bg, fg=C["ink_faint"],
                              font=("Segoe UI",7,"bold"))
        status_lbl.pack(pady=(8,0))
        angle = [0]; spin_id = [None]
        def tick():
            angle[0]=(angle[0]+90)%360
            status_lbl.configure(text=["◐","◓","◑","◒"][angle[0]//90])
            spin_id[0] = outer.after(180, tick) if spin_id[0] else None
        return {"outer":outer,"card":card,"top_bar":top_bar,"body":body,
                "icon_lbl":icon_lbl,"status_lbl":status_lbl,
                "spin_id":spin_id,"angle":angle,"tick":tick,"state":"init"}

    def _set_card_state(self, idx, state, show=False):
        if idx >= len(self._cards): return
        c = self._cards[idx]
        if c["spin_id"][0]:
            try: c["outer"].after_cancel(c["spin_id"][0])
            except: pass
            c["spin_id"][0] = None
        c["state"] = state

        cfg = {
            "waiting":    (C["surface"],  C["hair"],  "–",  C["ink_faint"], "EM ESPERA",  C["ink_faint"]),
            "processing": (C["surface"],  C["ok"],    "…",  C["ok"],        "EXECUTANDO", C["ok"]),
            "done":       (C["surface"],  C["ok"],    "✓",  C["ok"],        "CONCLUÍDO",  C["ok"]),
            "error":      (C["surface"],  C["err"],   "✗",  C["err"],       "ERRO",       C["err"]),
        }.get(state, (C["surface"], C["hair"], "–", C["ink_faint"], state, C["ink_faint"]))

        bg,bord,ic,ic_fg,st,st_fg = cfg
        c["card"].configure(bg=bg, highlightbackground=bord)
        c["top_bar"].configure(bg=bord)
        c["body"].configure(bg=bg)
        c["icon_lbl"].configure(bg=bg, text=ic, fg=ic_fg)
        c["status_lbl"].configure(bg=bg, text=st, fg=st_fg)
        for w in c["body"].winfo_children():
            try: w.configure(bg=bg)
            except: pass
        if show: c["outer"].grid()
        if state=="processing":
            c["spin_id"][0] = True
            c["tick"]()

    def _update_card_status(self, idx, text):
        if idx >= len(self._cards): return
        self._cards[idx]["status_lbl"].configure(text=text)

    def _start_worker(self):
        if self._worker_running: return
        self._worker_running = True
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        if not PLAYWRIGHT_OK:
            self._ui(lambda: messagebox.showerror("Erro","Playwright não disponível."))
            self._worker_running = False; return

        funcional = (getattr(self.controller,"bpm_funcional","") or "").strip()
        senha     = (getattr(self.controller,"bpm_password","") or "").strip()
        if not funcional or not senha:
            self._ui(lambda: messagebox.showerror("BPM","Credenciais não informadas."))
            self._worker_running = False; return

        PAINEL = "https://painelservicos.cloud.ihf/AplicAutSolicitacoesMiddle"
        PACE   = 0.85

        def pace(a=1.15, b=1.75):
            time.sleep(random.uniform(a*PACE, b*PACE))

        def pick_browser(p):
            for ch in ("chrome","msedge"):
                try: return p.chromium.launch(channel=ch, headless=False)
                except: pass
            return p.chromium.launch(headless=False)

        def wait_enabled(loc, timeout_ms=120_000):
            end = time.time() + timeout_ms/1000
            loc.wait_for(state="visible", timeout=timeout_ms)
            while time.time() < end:
                if self._cancel_requested: raise BPMUserCancelled()
                try:
                    if not loc.is_disabled(): return
                except: pass
                time.sleep(0.2)
            raise TimeoutError("Timeout aguardando campo habilitar.")

        def _to_dec(s):
            s = re.sub(r"[^\d,.\-]","",s or "")
            if not s: return None
            ld,lc = s.rfind("."),s.rfind(",")
            si = max(ld,lc)
            try:
                if si==-1: d=Decimal(re.sub(r"[^\d\-]","",s))
                else:
                    ip=re.sub(r"[^\d\-]","",s[:si]); dp=re.sub(r"[^\d]","",s[si+1:])
                    if not dp: return None
                    d=Decimal((ip if ip not in {"","-"} else "0")+"."+dp)
            except: return None
            return d.quantize(Decimal("0.01"),rounding=ROUND_HALF_UP)

        def fill_currency(ctx, loc, amount, timeout_ms=120_000):
            wait_enabled(loc, timeout_ms)
            loc.scroll_into_view_if_needed(); loc.click()
            pace(0.18,0.35); loc.press("Control+A"); loc.press("Backspace")
            pace(0.12,0.25); loc.press("Delete"); pace(0.15,0.30)
            digits = re.sub(r"\D","",amount or "")
            if not digits: raise RuntimeError(f"Valor inválido: {amount}")
            for ch in digits:
                loc.type(ch, delay=random.randint(int(85*PACE),int(150*PACE)))
            loc.press("Tab"); pace(0.45,0.85)
            exp = _to_dec(amount); got = _to_dec(loc.input_value())
            if exp is None or got is None or got != exp:
                raise RuntimeError(f"Valor mascarado divergente. Esperado {amount}, obtido '{loc.input_value()}'.")

        def resolve_frame(timeout_ms=120_000):
            deadline = time.time()+timeout_ms/1000
            while time.time()<deadline:
                if self._cancel_requested: raise BPMUserCancelled()
                for frame in list(page.frames):
                    try:
                        if frame.is_detached(): continue
                        if frame.locator("#nova-solicitacao").count()>0: return frame
                    except: pass
                time.sleep(0.35*PACE)
            raise TimeoutError("Timeout: #nova-solicitacao não encontrado.")

        def wait_continuar(ctx, timeout_ms=120_000):
            deadline = time.time()+timeout_ms/1000
            while time.time()<deadline:
                if self._cancel_requested: raise BPMUserCancelled()
                ok = ctx.evaluate("""(function(){
                    var el=document.querySelector('#continuar');
                    if(!el||el.disabled) return false;
                    var s=window.getComputedStyle(el);
                    if(s.display==='none'||s.visibility==='hidden') return false;
                    return true;
                })()""")
                if ok: return
                time.sleep(0.25)
            raise TimeoutError("Timeout: #continuar não ficou clicável.")

        def click_continuar(ctx, timeout_ms=120_000):
            loc = ctx.locator('input#continuar[type="submit"]').first
            for action in [
                lambda: loc.click(timeout=12_000),
                lambda: loc.click(force=True, timeout=12_000),
                lambda: ctx.evaluate("document.querySelector('#continuar').click()"),
            ]:
                try: action(); return
                except: pass
            raise RuntimeError("Falha ao clicar #continuar.")

        def wait_nova(ctx, timeout_ms=120_000):
            deadline = time.time()+timeout_ms/1000
            while time.time()<deadline:
                if self._cancel_requested: raise BPMUserCancelled()
                ok = ctx.evaluate("""(function(){
                    var el=document.querySelector('#nova-solicitacao');
                    if(!el||el.disabled) return false;
                    var s=window.getComputedStyle(el);
                    return !(s.display==='none'||s.visibility==='hidden');
                })()""")
                if ok: return
                time.sleep(0.4)
            raise TimeoutError("Timeout: #nova-solicitacao.")

        def click_nova(ctx, timeout_ms=120_000):
            wait_nova(ctx, timeout_ms)
            def done():
                try: ctx.locator("#tipooper").wait_for(state="visible",timeout=5_000); return True
                except: return False
            for action in [
                lambda: ctx.locator("#nova-solicitacao").first.click(timeout=8_000),
                lambda: ctx.evaluate("document.querySelector('#nova-solicitacao').click()"),
                lambda: ctx.locator("#nova-solicitacao").first.click(force=True,timeout=8_000),
            ]:
                try: action()
                except: pass
                for _ in range(12):
                    if done(): return
                    time.sleep(0.5)
            raise RuntimeError("Falha ao clicar #nova-solicitacao.")

        def ensure_painel():
            if self._cancel_requested: raise BPMUserCancelled()
            page.goto(PAINEL, wait_until="domcontentloaded")
            if "/Home" in page.url: page.goto(PAINEL, wait_until="domcontentloaded")
            if page.locator("#username").count()>0 or "login.itau/oauth" in page.url:
                page.locator("#username").fill(funcional)
                page.locator("#password").fill(senha)
                page.locator("#btLogin").click()
                page.wait_for_load_state("domcontentloaded")
                page.goto(PAINEL, wait_until="domcontentloaded")
            ctx = resolve_frame(timeout_ms=120_000)
            wait_nova(ctx, timeout_ms=120_000)
            return ctx

        STEP = 15_000
        with sync_playwright() as p:
            browser = pick_browser(p)
            self._browser = browser
            page = browser.new_page()
            page.set_default_timeout(120_000)
            try:
                for idx, item in enumerate(self._selected):
                    if self._cancel_requested: raise BPMUserCancelled()
                    client_name = item.get("cliente","")
                    raw_amount  = item.get("valor","")
                    info = BPM_CLIENT_DATA.get(client_name)
                    if not info:
                        self._log_line(f"[{client_name}] Sem mapeamento interno.", "err"); continue
                    amount_web = _fmt_brl_plain_web(raw_amount)
                    if not amount_web:
                        self._log_line(f"[{client_name}] Valor inválido: {raw_amount}", "err"); continue

                    self._log_line(f"━━ {client_name} ━━", "heading")
                    tentativa = 0
                    while True:
                        if self._cancel_requested: raise BPMUserCancelled()
                        tentativa += 1
                        try:
                            ctx = ensure_painel()
                            self._log_line(f"  Painel pronto.", "step")
                            click_nova(ctx)
                            pace()
                            self._log_line("  Nova solicitação aberta.", "step")

                            ctx.locator("#tipooper").wait_for(state="visible",timeout=STEP)
                            ctx.locator("#tipooper").select_option(value="AT")
                            pace()
                            ctx.locator("#CPNJ").fill(info["CNPJ"]); pace()
                            ctx.locator('input[name="agenciafiltro"]').fill(info["AG"]); pace()
                            ctx.locator('input[name="contadacfiltro"]').fill(info["CONTA"]); pace()
                            self._log_line("  Dados do cliente preenchidos.", "step")

                            loc_proc = ctx.locator("#processar-nova")
                            wait_enabled(loc_proc, STEP); loc_proc.click(); pace()

                            loc_plat = ctx.locator('input[name="plataforma"]')
                            wait_enabled(loc_plat,STEP); loc_plat.fill(info["PLATAFORMA"]); pace()

                            loc_fn = ctx.locator("#funcionalac")
                            wait_enabled(loc_fn,STEP); loc_fn.fill(funcional)
                            loc_fn.press("Tab")
                            try: loc_fn.evaluate("el=>{el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));el.blur();}")
                            except: pass
                            time.sleep(0.35*PACE)
                            wait_continuar(ctx,STEP); click_continuar(ctx,STEP); pace()
                            self._log_line("  Filtro AplicAut preenchido.", "step")

                            ctx.locator('select[name="produto"]').wait_for(state="visible",timeout=STEP)
                            ctx.locator('select[name="produto"]').select_option(value="34"); pace()
                            ctx.locator('select[name="tpOperacao"]').wait_for(state="visible",timeout=STEP)
                            ctx.locator('select[name="tpOperacao"]').select_option(value="1"); pace()

                            loc_buscar = ctx.locator('input[type="submit"][data-ng-click="BuscarListaPN()"]')
                            if loc_buscar.count()==0: loc_buscar=ctx.locator('input[type="submit"][value="Processar"]')
                            wait_enabled(loc_buscar,STEP); loc_buscar.click(); pace()

                            checkbox = ctx.locator('input[type="checkbox"]')
                            checkbox.first.wait_for(state="visible",timeout=STEP)
                            checkbox.first.click(); pace()

                            loc_val = ctx.locator('input[name="ValordaOperacao0"]')
                            self._log_line(f"  Preenchendo valor {raw_amount}…", "step")
                            fill_currency(ctx, loc_val, amount_web, STEP); pace()
                            self._log_line("  Valor preenchido.", "step")

                            loc_final = ctx.locator('input[type="submit"][ng-click="vm.finalizarSol()"]')
                            if loc_final.count()==0: loc_final=ctx.locator('input[type="submit"][value="Continuar"]')
                            wait_enabled(loc_final,STEP); loc_final.click(); pace()

                            loc_inc = ctx.locator('input[type="submit"][value="Incluir"]')
                            if loc_inc.count()==0: loc_inc=ctx.locator('input[type="submit"][ng-click="vm.IncluirConta()"]')
                            wait_enabled(loc_inc,STEP); loc_inc.click(); pace(1.4,2.1)

                            def try_grey():
                                loc_g=ctx.locator('input[type="button"].grey[ng-click="vm.continuar()"]')
                                if loc_g.count()==0: return False
                                wait_enabled(loc_g.first,STEP); loc_g.first.click(); return True

                            try_grey(); pace()
                            try: try_grey()
                            except: pass

                            loc_v = ctx.locator('input[type="button"].grey[ng-click="vm.voltarAplicAut()"]')
                            if loc_v.count()==0: loc_v=ctx.locator('input.grey[ng-click="vm.voltarAplicAut()"]')
                            if loc_v.count()>0:
                                wait_enabled(loc_v.first,max(STEP,25_000)); loc_v.first.click()

                            ctx_after = resolve_frame(120_000)
                            wait_nova(ctx_after,120_000)
                            self._log_line(f"  {client_name} → CONCLUÍDO ✓", "ok")
                            self._ui(lambda i=idx: self._set_card_state(i,"done",show=True))
                            break

                        except BPMUserCancelled: raise
                        except Exception as e:
                            self._log_line(f"  Tentativa {tentativa} falhou: {e}", "warn")
                            if tentativa >= 20:
                                self._log_line(f"  {client_name} → FALHA PERMANENTE ✗", "err")
                                self._ui(lambda i=idx: self._set_card_state(i,"error",show=True))
                                raise RuntimeError(f"Falha recorrente {client_name}: {e}") from e
                            try: page.goto(PAINEL, wait_until="domcontentloaded")
                            except: pass

            except BPMUserCancelled: self._log_line("Operações canceladas pelo usuário.", "warn")
            except Exception as e:
                if not self._cancel_requested:
                    em = str(e)
                    self._log_line(f"ERRO GERAL: {em}", "err")
                    self._ui(lambda: messagebox.showerror("Erro na rotina", em))
            finally:
                self._browser = None
                try: browser.close()
                except: pass
                self._worker_running = False
                self._started = False


# ═══════════════════════════════════════════════════════════════════════════════
#  APP — janela principal com sidebar
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mesa Itaú — Risco Sacado")
        self.geometry("1060x720")
        self.minsize(860, 580)
        self.configure(bg=C["bg"])
        self._setup_ttk_styles()
        self.bpm_run_selection = []
        self.bpm_funcional     = ""
        self.bpm_password      = ""
        self.rotina_em_execucao= None
        self._active_frame     = "Home"

        # Layout: sidebar | conteúdo
        self._sidebar = Sidebar(self, self)
        self._sidebar.pack(side="left", fill="y")

        make_hairline(self, orient="v", bg=C["hair"]).pack(side="left", fill="y")

        self._content = tk.Frame(self, bg=C["bg"])
        self._content.pack(side="left", fill="both", expand=True)
        self._content.rowconfigure(0, weight=1)
        self._content.columnconfigure(0, weight=1)

        self.frames = {}
        for Cls, name in [
            (HomeFrame,              "Home"),
            (RotinasFrame,           "Rotinas"),
            (ShareFrame,             "Share"),
            (BPMConfigFrame,         "BPM_CONFIG"),
            (BPMFrame,               "BPM"),
            (LimitesInvertidoFrame,  "LimitesInvertido"),
        ]:
            f = Cls(self._content, self)
            self.frames[name] = f
            f.grid(row=0, column=0, sticky="nsew")

        self.show_frame("Home")

    def _setup_ttk_styles(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except: pass
        s.configure("TCombobox",
                    fieldbackground=C["bg"], background=C["surface"],
                    foreground=C["ink"], selectbackground=C["surface2"],
                    selectforeground=C["ink"], borderwidth=1,
                    lightcolor=C["hair"], darkcolor=C["hair"])
        s.map("TCombobox", fieldbackground=[("readonly",C["bg"])],
              selectbackground=[("!focus",C["surface2"])])
        s.configure("TScrollbar", background=C["surface2"], troughcolor=C["bg"],
                    borderwidth=0, arrowsize=10)
        s.map("TScrollbar", background=[("active",C["surface3"])])

    def show_frame(self, name):
        # remap: "BPM" sem seleção → config
        if name == "BPM" and not getattr(self,"bpm_run_selection",[]):
            name = "BPM_CONFIG"
        f = self.frames.get(name)
        if f is None: return
        self._active_frame = name
        sidebar_name = name if name in ("Home","Rotinas","Share","LimitesInvertido") \
                       else ("BPM" if name in ("BPM","BPM_CONFIG") else name)
        self._sidebar.set_active(sidebar_name)
        f.tkraise()
        if hasattr(f,"on_show"):
            try: f.on_show()
            except: pass


# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()