import os, sys, re, random, struct, tkinter as tk, threading, time, tempfile, shutil, webbrowser, ctypes, json as _json_mod, uuid as _uuid_mod
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

def make_hairline(parent, orient="h", **kwargs):
    kw = {"bg": C["hair"]}
    kw.update(kwargs)
    if orient == "h":
        return tk.Frame(parent, height=1, **kw)
    else:
        return tk.Frame(parent, width=1, **kw)


def _make_dot(parent, color, size=10, bg=None):
    """Bolinha colorida via Canvas — estilo Notion."""
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
    "Home":             "Início",
    "Rotinas":          "Rotinas",
    "Share":            "Cadastro Share",
    "BPM_CONFIG":       "Configurar BPM",
    "BPM":              "BPM — Operações",
    "LimitesInvertido": "Limites Invertido",
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
        WM_NCLBUTTONDOWN = 0x00A1
        HTCAPTION = 2
        ctypes.windll.user32.ReleaseCapture()
        ctypes.windll.user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
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
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        self.refresh_bindings()

    def _on_canvas(self, event):
        if self._canvas_alive():
            self._canvas.itemconfigure(self._win, width=event.width)

    def _on_destroy(self, event):
        pass


class Sidebar(tk.Frame):
    NAV = [
        ("Home",             "⌂",  "Início"),
        ("Rotinas",          "◈",  "Rotinas"),
        ("Share",            "⊕",  "Cadastro Share"),
        ("BPM",              "⚡",  "BPM"),
        ("LimitesInvertido", "⬡",  "Limites Invertido"),
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
        {"name": "Cadastro Share",   "sub": "Extração e análise de PDF",          "icon": "⊕", "frame": "Share",           "color": "#5a9e72"},
        {"name": "BPM",              "sub": "Abertura de solicitações",            "icon": "⚡", "frame": "BPM_CONFIG",      "color": "#EC7000"},
        {"name": "Limites Invertido","sub": "Consulta LTC e limites disponíveis",  "icon": "⬡", "frame": "LimitesInvertido","color": "#c87941"},
        {"name": "Rotinas",          "sub": "Sequências configuráveis",            "icon": "◈", "frame": "Rotinas",         "color": "#8b72c9"},
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
            ("Consultar limites",        "LimitesInvertido"),
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

    # ── Alertas ──────────────────────────────────────────────────────────────
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

    # ── PATCH 1: popup de alerta visual customizado ──────────────────────────
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

        # Força ao topo no Windows — aparece sobre qualquer app aberto
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

        # Borda laranja no topo (identidade visual do app)
        tk.Frame(dlg, bg=C["accent"], height=3).pack(fill="x")

        body = tk.Frame(dlg, bg=C["surface"], padx=28, pady=20)
        body.pack(fill="both", expand=True)

        # Cabeçalho: sino + hora
        hdr_row = tk.Frame(body, bg=C["surface"])
        hdr_row.pack(fill="x")
        tk.Label(hdr_row, text="\U0001f514", bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 18)).pack(side="left")
        tk.Label(hdr_row, text=hora, bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 22, "bold")).pack(side="left", padx=(10, 0))

        make_hairline(body, bg=C["hair"]).pack(fill="x", pady=(14, 12))

        # Nome da rotina
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

        # Centraliza na tela
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

        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._sf.link_wheel(self)
        body = self._sf.inner
        body.configure(bg=C["bg"])

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
                        # ── PATCH 2: espera inteligente de carregamento ──────────
                        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
                        if self._cancel_requested: break

                        # Espera networkidle (max 30s) — sem sleep fixo
                        try:
                            page.wait_for_load_state("networkidle", timeout=30_000)
                        except Exception:
                            pass

                        # Polling JS: aguarda spinners/loaders sumirem (max 25s)
                        _deadline = time.time() + 25
                        while time.time() < _deadline:
                            if self._cancel_requested:
                                break
                            try:
                                _loading = page.evaluate(
                                    "(function(){"
                                    "var sels=['.loading','.spinner','[class*=\"loading\"]',"
                                    "'[class*=\"spinner\"]','.u-loading',"
                                    "'[aria-busy=\"true\"]','.itau-spinner'];"
                                    "for(var i=0;i<sels.length;i++){"
                                    "var els=document.querySelectorAll(sels[i]);"
                                    "for(var j=0;j<els.length;j++){"
                                    "var s=window.getComputedStyle(els[j]);"
                                    "if(s.display!=='none'&&s.visibility!=='hidden'"
                                    "&&s.opacity!=='0') return true;}}"
                                    "return false;})()"
                                )
                                if not _loading:
                                    break
                            except Exception:
                                break
                            time.sleep(0.5)

                        # Pausa mínima de segurança para renderização final
                        time.sleep(2)
                        if self._cancel_requested: break

                        # ler LTC — 3 tentativas com pausa entre elas
                        ltc_str = None
                        for _ltc_try in range(3):
                            try:
                                all_spans = page.locator("span.u-font-size--14.u-ml--8.u-mt--8.u-block")
                                for si in range(all_spans.count()):
                                    txt = (all_spans.nth(si).inner_text() or "").strip()
                                    if RE_DATE.match(txt):
                                        ltc_str = txt
                                        break
                            except Exception:
                                pass
                            if not ltc_str:
                                try:
                                    body_text = page.locator("body").inner_text()
                                    m = RE_DATE.search(body_text)
                                    if m:
                                        ltc_str = m.group(0)
                                except Exception:
                                    pass
                            if ltc_str:
                                break
                            time.sleep(2)
                        # ────────────────────────────────────────────────────────

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

        self._sf = ScrollableFrame(self, bg=C["bg"])
        self._sf.pack(fill="both", expand=True)
        self._sf.link_wheel(self)
        self._grid_wrap = self._sf.inner
        self._grid_wrap.configure(bg=C["bg"])

        self._grid = tk.Frame(self._grid_wrap, bg=C["bg"])
        self._grid.pack(padx=32, pady=(16,8), fill="x")
        for c in range(3):
            self._grid.columnconfigure(c, weight=1, uniform="bcards")

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
        lsb = MinimalScrollbar(log_frame, command=self._log.yview, bg=C["bg"])
        lsb.grid(row=0, column=1, sticky="ns")
        self._log.configure(yscrollcommand=lsb.set)
        bind_text_mousewheel(self._log)
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
        self._sf.refresh_bindings()
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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mesa Itaú — Risco Sacado")
        self.geometry("1060x720")
        self.minsize(860, 580)
        self.configure(bg=C["bg"])
        self._setup_ttk_styles()
        ico_path = _ensure_ico_path()
        if ico_path:
            try:
                self.iconbitmap(default=ico_path)
            except Exception:
                pass
        self.overrideredirect(True)
        self.bpm_run_selection = []
        self.bpm_funcional     = ""
        self.bpm_password      = ""
        self.rotina_em_execucao= None
        self._active_frame     = "Home"

        self._shell = tk.Frame(self, bg=C["bg"])
        self._shell.pack(fill="both", expand=True)

        self._titlebar = AppTitleBar(self._shell, self)
        self._titlebar.pack(side="top", fill="x")

        self._main = tk.Frame(self._shell, bg=C["bg"])
        self._main.pack(fill="both", expand=True)

        self._sidebar = Sidebar(self._main, self)
        self._sidebar.pack(side="left", fill="y")

        make_hairline(self._main, orient="v", bg=C["hair"]).pack(side="left", fill="y")

        self._content = tk.Frame(self._main, bg=C["bg"])
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

        self._statusbar = AppStatusBar(self._shell, self)
        self._statusbar.pack(side="bottom", fill="x")

        self.show_frame("Home")
        self.after(120, self._apply_window_chrome)

    def _apply_window_chrome(self):
        apply_modern_window_chrome(self)
        apply_frameless_resize(self)
        apply_windows_shell(self)

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

    def _refresh_frame_scroll(self, frame):
        if isinstance(frame, ScrollableFrame):
            frame.refresh_bindings()
        for child in frame.winfo_children():
            self._refresh_frame_scroll(child)

    def show_frame(self, name):
        if name == "BPM" and not getattr(self,"bpm_run_selection",[]):
            name = "BPM_CONFIG"
        f = self.frames.get(name)
        if f is None: return
        self._active_frame = name
        sidebar_name = name if name in ("Home","Rotinas","Share","LimitesInvertido") \
                       else ("BPM" if name in ("BPM","BPM_CONFIG") else name)
        self._sidebar.set_active(sidebar_name)
        self._titlebar.set_module(name)
        self._statusbar.set_module(name)
        f.tkraise()
        self._refresh_frame_scroll(f)
        if hasattr(f,"on_show"):
            try: f.on_show()
            except: pass

if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
        except Exception:
            pass
    App().mainloop()