from __future__ import annotations

import os
import re
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Optional


PLAYWRIGHT_OK = False
PLAYWRIGHT_ERRO = ""
try:
    from playwright.sync_api import Page, sync_playwright 
    PLAYWRIGHT_OK = True
except ImportError as _e:
    Page = object 
    sync_playwright = None  
    PLAYWRIGHT_ERRO = f"Playwright não instalado: {_e}"
except Exception as _e:
    Page = object  
    sync_playwright = None  
    PLAYWRIGHT_ERRO = f"Erro ao importar Playwright: {_e}"

LOGO_FILENAME = "itau-logo-png_seeklogo-74122.png"

URL_LOGIN  = "https://investidor.libercapital.com.br/investidores/entrar"
URL_PAINEL = "https://investidor.libercapital.com.br/investidor/painel_taxas"
CNPJ  = os.environ.get("TAXAS_CNPJ",  "60701190481609")
SENHA = os.environ.get("TAXAS_SENHA", "mesaibbarecebiveis")

CLIENTES = [
    "ELECNOR",
    "ELECNOR II",
    "ALPARGATAS",
    "ALPARGATAS II",
    "CBA",
    "CBA II",
    "KLABIN",
    "KLABIN II",
    "AJINOMOTO",
    "AJINOMOTO II",
    "COPA ENERGIA",
    "COPA ENERGIA II",
    "CASSOL",
    "CASSOL II",
    "FARTURA",
    "FARTURA II",
    "Nexa",
    "Nexa II",
]

DEFAULT_TIMEOUT_MS = 22_000
CLICK_TIMEOUT_MS   = 10_000
NAV_TIMEOUT_MS     = 35_000


def app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(p: str) -> str:
    base = getattr(sys, "_MEIPASS", app_base_dir())
    return os.path.join(base, p)


def _normalizar_nome_cliente(nome: str) -> str:
    return " ".join(nome.strip().split()).lower()



def listar_subpastas_clientes(pasta_boletos: str) -> list[str]:
    """Usa scandir — uma única chamada de sistema, sem isdir extra."""
    try:
        with os.scandir(pasta_boletos) as it:
            return [e.name for e in it if e.is_dir(follow_symlinks=False)]
    except OSError:
        return []


def arquivo_mais_recente_em(pasta: str) -> Optional[str]:
    """
    Encontra o arquivo mais recente via scandir + stat inline.
    Uma única passagem, sem lista intermediária nem chamadas separadas a getmtime.
    Ignora temporários do Office (~) e arquivos ocultos (.).
    """
    melhor_path: Optional[str] = None
    melhor_mtime: float = -1.0
    try:
        with os.scandir(pasta) as it:
            for entry in it:
                if not entry.is_file(follow_symlinks=False):
                    continue
                nome = entry.name
                if nome.startswith("~") or nome.startswith("."):
                    continue
                try:
                    mtime = entry.stat(follow_symlinks=False).st_mtime
                except OSError:
                    continue
                if mtime > melhor_mtime:
                    melhor_mtime = mtime
                    melhor_path  = entry.path
    except OSError:
        pass
    return melhor_path



_SubpastaIdx = dict[str, str] 


def _indexar_subpastas(subpastas: list[str]) -> _SubpastaIdx:
    """Constrói índice normalizado sem tocar em arquivos."""
    return {
        _normalizar_nome_cliente(s).replace("_", " "): s
        for s in subpastas
    }


def planilha_taxas_para_cliente(
    pasta_boletos: str,
    cliente: str,
    idx_subpastas: _SubpastaIdx,
) -> Optional[str]:
    """
    Busca lazy: só abre o diretório 'Taxas Enviadas' do cliente quando chamada.
    Usa o índice de subpastas pré-montado (sem I/O de arquivo).
    """
    chave = _normalizar_nome_cliente(cliente).replace("_", " ")
    sub   = idx_subpastas.get(chave)
    if not sub:
        return None
    taxas_dir = os.path.join(pasta_boletos, sub, "Taxas Enviadas")
    if not os.path.isdir(taxas_dir):
        return None
    return arquivo_mais_recente_em(taxas_dir)


LogFn    = Callable[[str, str], None]
CancelFn = Callable[[], bool]


@dataclass
class TaxasPlaywrightSession:
    """Uma sessão: login, painel e envio por cliente. Não depende de Tk."""

    page: object 
    log: LogFn
    cancelado: CancelFn

    def _abort(self) -> None:
        if self.cancelado():
            raise InterruptedError()

    def _aguardar_botao_clicavel(self, texto: str, timeout_ms: int) -> None:
        """
        Espera o botão aparecer e estar habilitado via locators nativos.
        Substitui o hit-test JS (elementFromPoint) que falha quando o botão
        está fora da viewport ou coberto por overlay de carregamento.
        """
        self._abort()
        deadline = time.time() + (timeout_ms / 1000.0)
        last_err: Optional[Exception] = None
        esc = re.escape(texto.strip())
        locs = [
            self.page.get_by_role("button", name=re.compile(rf"^\s*{esc}\s*$", re.IGNORECASE)).first,
            self.page.get_by_role("button", name=re.compile(esc, re.IGNORECASE)).first,
            self.page.locator(f'button:has-text("{texto}")').first,
        ]
        while time.time() < deadline:
            self._abort()
            for loc in locs:
                try:
                    if loc.count() > 0 and loc.first.is_visible():
                        disabled = loc.first.get_attribute("disabled")
                        aria_dis = loc.first.get_attribute("aria-disabled")
                        if disabled is None and aria_dis not in ("true", "1"):
                            return  
                except Exception as e:
                    last_err = e
            time.sleep(0.15)
        raise TimeoutError(
            f'Botão "{texto}" não ficou visível/habilitado em {timeout_ms}ms. '
            f'Último erro: {last_err}'
        )

    def _locators_botao(self, texto: str):
        esc = re.escape(texto.strip())
        texto_norm = " ".join((texto or "").strip().lower().split())
        locs = []

        
        if texto_norm == "criar nova taxa":
            locs.extend(
                [
                    self.page.locator("#InvestorSpreadsheetRates-react-component > div.sc-dZtSGr.hlfYDf > button").first,
                    self.page.locator('xpath=//*[@id="InvestorSpreadsheetRates-react-component"]/div[1]/button').first,
                    self.page.locator('xpath=/html/body/div[1]/div[1]/main/div/div[1]/button').first,
                    self.page.locator("#InvestorSpreadsheetRates-react-component button").first,
                ]
            )

        locs.extend([
            self.page.get_by_role("button", name=re.compile(rf"^\s*{esc}\s*$", re.IGNORECASE)).first,
            self.page.get_by_role("button", name=re.compile(esc, re.IGNORECASE)).first,
            self.page.locator(f'button:has-text("{texto}")').first,
            self.page.locator(f'text="{texto}"').first,
        ])
        return locs

    def _clicar_botao(self, texto: str, timeout_pronto_ms: int) -> None:
        self._abort()
        deadline  = time.time() + (timeout_pronto_ms / 1000.0)
        last_err: Optional[Exception] = None

        self._fechar_popups(timeout_ms=800)

        # Primeira passada rápida: só 1.2s por locator. A maioria das tentativas
        # nesta lista não vai bater (são fallbacks), então gastar 4s em cada
        # uma sequencialmente era o principal gargalo de tempo aqui.
        for loc in self._locators_botao(texto):
            self._abort()
            try:
                loc.wait_for(state="visible", timeout=min(1_200, timeout_pronto_ms))
                if hasattr(loc, "is_disabled") and loc.is_disabled():
                    continue
                loc.click(timeout=CLICK_TIMEOUT_MS)
                return
            except Exception as e:
                last_err = e

       
        while time.time() < deadline:
            self._abort()
            for loc in self._locators_botao(texto):
                try:
                    if loc.count() == 0:
                        continue
                    if loc.first.is_visible() and not loc.first.is_disabled():
                        try:
                            loc.first.scroll_into_view_if_needed(timeout=2_000)
                        except Exception:
                            pass
                        loc.first.click(timeout=CLICK_TIMEOUT_MS)
                        return
                except Exception as e:
                    last_err = e
            try:
                clicked = self.page.evaluate(
                    """(texto) => {
                        const t = (texto || '').trim().toLowerCase();
                        const exact = document.querySelector('#InvestorSpreadsheetRates-react-component > div.sc-dZtSGr.hlfYDf > button');
                        if (exact && (exact.innerText || '').trim().toLowerCase().includes(t) && !exact.disabled) {
                          exact.click();
                          return true;
                        }
                        const btn = [...document.querySelectorAll('button')]
                          .find(b => (b.innerText || '').trim().toLowerCase().includes(t));
                        if (!btn || btn.disabled) return false;
                        btn.click();
                        return true;
                    }""",
                    texto,
                )
                if clicked:
                    return
            except Exception as e:
                last_err = e
            time.sleep(0.12)

        raise TimeoutError(
            f'Botão "{texto}" não ficou clicável a tempo. Último erro: {last_err}'
        )

    def configurar_timeouts(self) -> None:
        self.page.set_default_timeout(DEFAULT_TIMEOUT_MS)
        self.page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

    def _fechar_popups(self, timeout_ms: int = 3_000) -> None:
        """
        Fecha popups/overlays que atrapalham (ex.: pesquisa Hotjar).
        Não lança erro se não houver popup — é best-effort e rápido.
        """
        self._abort()
        seletores = [
            'dialog[open] button[aria-label="Fechar"]',
            'button[aria-label="Fechar"]',
            '[data-testid="survey-root"] button[aria-label="Fechar"]',
        ]
        deadline = time.time() + (timeout_ms / 1000.0)
        while time.time() < deadline:
            fechou_algo = False
            for sel in seletores:
                try:
                    loc = self.page.locator(sel).first
                    if loc.count() > 0 and loc.is_visible():
                        loc.click(timeout=1_000)
                        fechou_algo = True
                        self.log("  Popup fechado.", "info")
                except Exception:
                    continue
            if not fechou_algo:
                return
            time.sleep(0.1)

    def login(self) -> None:
        p = self.page
        self.log("  Abrindo login…", "info")
        p.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)

        # O popup de pesquisa (Hotjar) pode surgir logo após o load, cobrindo
        # os campos. Fecha antes de qualquer interação.
        self._fechar_popups()

        self.log("  Tipo CNPJ…", "info")
        for seletor in (
            'label.MuiFormControlLabel-root:has(p[data-testid="Typography"]:text-is("CNPJ"))',
            'p[data-testid="Typography"]:text-is("CNPJ")',
            'text="CNPJ"',
        ):
            self._abort()
            try:
                loc = p.locator(seletor).first
                loc.wait_for(state="visible", timeout=6_000)
                loc.click(timeout=4_000)
                break
            except Exception:
                continue

        self._fechar_popups(timeout_ms=1_500)

        self.log("  Credenciais…", "info")
        doc = p.locator('[data-testid="DocumentInput"] input')
        doc.wait_for(state="visible", timeout=12_000)
        doc.click()
        doc.fill(CNPJ)

        pwd = p.locator('[data-testid="PasswordInput"] input')
        pwd.wait_for(state="visible", timeout=12_000)
        pwd.click()
        pwd.fill(SENHA)

        self.log("  Enviando login…", "info")
        submit  = p.locator('[data-testid="SubmitButton"]')
        submit.wait_for(state="visible", timeout=12_000)
        deadline = time.time() + 25
        while time.time() < deadline:
            self._abort()
            try:
                if not submit.is_disabled():
                    break
            except Exception:
                pass
            time.sleep(0.12)
        else:
            raise TimeoutError("Botão de entrar não habilitou.")

        submit.click(timeout=5_000)
        p.wait_for_function(
            """() => {
                const h = window.location.href;
                return h.includes('investidor') && !h.includes('/entrar');
            }""",
            timeout=NAV_TIMEOUT_MS,
        )
        try:
            p.wait_for_load_state("domcontentloaded", timeout=8_000)
        except Exception:
            pass

    def ir_painel_taxas(self, recarregar: bool = False) -> None:
        p = self.page
        self._abort()
        if recarregar or URL_PAINEL not in (p.url or ""):
            p.goto(URL_PAINEL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
     
        try:
            p.wait_for_load_state("networkidle", timeout=4_000)
        except Exception:
            pass

        self._fechar_popups(timeout_ms=1_500)

        try:
            p.evaluate("window.scrollTo(0, 0)")
        except Exception:
            pass
        self._aguardar_botao_clicavel("Criar Nova Taxa", 30_000)

    def garantir_painel(self) -> None:
        if URL_PAINEL not in (self.page.url or ""):
            self.log("  Voltando ao painel de taxas…", "info")
            self.ir_painel_taxas(recarregar=False)

    def enviar_taxa_cliente(self, nome_cliente: str, caminho_planilha: str) -> None:
        p    = self.page
        arquivo = os.path.basename(caminho_planilha)

        self.garantir_painel()

        self.log(f"  [{nome_cliente}] Nova taxa…", "info")
        self._clicar_botao("Criar Nova Taxa", 28_000)

        self.log(f"  [{nome_cliente}] Anexo: {arquivo}", "info")
        inp = p.locator('input[type="file"]').first
        inp.wait_for(state="attached", timeout=10_000)
        inp.set_input_files(caminho_planilha)

        self.log(f"  [{nome_cliente}] Enviar…", "info")
        self._clicar_botao("ENVIAR", 18_000)

        self.log(f"  [{nome_cliente}] Confirmar (1/2)…", "info")
        self._clicar_botao("Confirmar", 22_000)

        self.log(f"  [{nome_cliente}] Confirmar (2/2)…", "info")
        self._clicar_botao("CONFIRMAR", 22_000)

        try:
            p.wait_for_load_state("domcontentloaded", timeout=15_000)
        except Exception:
            pass
        if URL_PAINEL not in (p.url or ""):
            self.ir_painel_taxas(recarregar=False)


def executar_rotina_taxas(
    pasta_boletos: str,
    log: LogFn,
    cancelado: CancelFn,
    browser_holder: list,
) -> tuple[int, int, int]:
    """
    Abre o navegador, faz login, percorre CLIENTES.
    Retorna (sucesso, pulados, erros).
    browser_holder[0] = instância do browser (para fechar no Voltar).
    """
 
    if not PLAYWRIGHT_OK or sync_playwright is None:
        msg = PLAYWRIGHT_ERRO or "Playwright não está instalado/disponível."
        log(f"ERRO: {msg}", "err")
        log("Execute: pip install playwright && python -m playwright install chromium", "warn")
        return (0, 0, 0)

    log(f"  Pasta: {pasta_boletos}", "info")
    subpastas = listar_subpastas_clientes(pasta_boletos)
    if not subpastas:
        log("Aviso: não há subpastas na pasta Boletos.", "warn")
    else:
        log(f"  Subpastas encontradas: {len(subpastas)} — iniciando browser…", "info")


    idx_subpastas = _indexar_subpastas(subpastas)

    sucesso = pulados = erros = 0

    try:
        with sync_playwright() as pw:
            browser = None
            try:

                launched = False
                for canal in ("chrome", "msedge"):
                    try:
                        log(f"  Tentando abrir browser: {canal}…", "info")
                        browser = pw.chromium.launch(channel=canal, headless=False)
                        log(f"  Browser aberto ({canal}).", "ok")
                        launched = True
                        break
                    except Exception as e_launch:
                        log(f"  {canal} indisponível: {e_launch}", "warn")

                if not launched:
                    try:
                        log("  Tentando Chromium embutido…", "info")
                        browser = pw.chromium.launch(headless=False)
                        log("  Chromium embutido aberto.", "ok")
                    except Exception as e_chromium:
                        raise RuntimeError(
                            f"Nenhum browser disponível (Chrome, Edge ou Chromium). "
                            f"Detalhes: {e_chromium}"
                        )

                browser_holder.clear()
                browser_holder.append(browser)

                ctx  = browser.new_context()
                page = ctx.new_page()
                sess = TaxasPlaywrightSession(page, log, cancelado)
                sess.configurar_timeouts()

                log("\n─── Login ────────────────────────────", "heading")
                sess.login()
                log("  Indo para painel de taxas…", "info")
                sess.ir_painel_taxas(recarregar=True)
                log("Login concluído.", "ok")

                total = len(CLIENTES)
                log("\n─── Envio de taxas ───────────────────", "heading")

                for i, cliente in enumerate(CLIENTES):
                    if cancelado():
                        log("\nCancelado pelo usuário.", "warn")
                        break

                    log(f"\n[{i + 1}/{total}] {cliente}", "heading")

                   
                    planilha = planilha_taxas_para_cliente(
                        pasta_boletos, cliente, idx_subpastas
                    )
                    if not planilha:
                        log("  Planilha não encontrada — pulando.", "warn")
                        pulados += 1
                        continue

                    log(f"  Arquivo: {os.path.basename(planilha)}", "info")
                    try:
                        sess.enviar_taxa_cliente(cliente, planilha)
                        sucesso += 1
                        log("  ✓ Enviado.", "ok")
                    except InterruptedError:
                        log("\nCancelado pelo usuário.", "warn")
                        break
                    except Exception as e:
                        erros += 1
                        log(f"  ✗ Erro: {str(e)[:200]}", "err")

                log(
                    f"\n─── Concluído ────────────────────────\n"
                    f"  Enviados: {sucesso}\n"
                    f"  Pulados:  {pulados}\n"
                    f"  Erros:    {erros}",
                    "heading",
                )

            finally:
                
                browser_holder.clear()
                if browser:
                    try:
                        browser.close()
                    except Exception:
                        pass

    except InterruptedError:
        log("Execução interrompida.", "warn")
    except Exception as e_geral:
        log(f"\nERRO GERAL: {e_geral}", "err")

    return (sucesso, pulados, erros)

class TaxasLargeFrame(ttk.Frame):
    ACCENT = "#e8a020"

    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self._folder  = tk.StringVar()
        self._worker_running      = False
        self._cancel_requested    = False
        self._browser_ref: list   = []
        self._worker_thread: Optional[threading.Thread] = None
        self._build_ui()

    def _ui(self, fn):
        """Agenda execução segura no thread da UI."""
        try:
            self.after(0, fn)
        except Exception:
            pass

    def _build_ui(self):
        bg = self.COLORS["bg"]

       
        top = tk.Frame(self, bg=bg)
        top.pack(fill="x", padx=20, pady=(18, 0))
        tk.Button(
            top,
            text="← Voltar",
            command=self._go_back,
            bg=bg, fg="#8c8c8c", bd=0, relief="flat",
            activebackground=bg, activeforeground="#f2f2f2",
            cursor="hand2", font=("Segoe UI", 9),
        ).pack(side="left")

     
        tk.Label(self, text="Taxas Large", bg=bg, fg="#f2f2f2",
                 font=("Segoe UI", 15, "bold")).pack(pady=(20, 0))
        tk.Label(self, text="Envio automático de planilhas de taxas para cada cliente.",
                 bg=bg, fg="#666666", font=("Segoe UI", 10)).pack(pady=(4, 0))

       
        if not PLAYWRIGHT_OK:
            tk.Label(
                self,
                text=f"⚠ Playwright indisponível — {PLAYWRIGHT_ERRO}",
                bg=bg, fg="#e74c3c",
                font=("Segoe UI", 9),
                wraplength=700, justify="left",
            ).pack(padx=30, pady=(6, 0), anchor="w")

       
        folder_bar = tk.Frame(self, bg=bg)
        folder_bar.pack(fill="x", padx=30, pady=(20, 0))

        tk.Label(folder_bar, text="Pasta Boletos:", bg=bg, fg="#8c8c8c",
                 font=("Segoe UI", 9)).pack(side="left")

        tk.Entry(
            folder_bar, textvariable=self._folder,
            bg="#1a1a1a", fg="#c0c0c0", insertbackground="#c0c0c0",
            relief="flat", highlightthickness=1,
            highlightbackground="#2e2e2e", highlightcolor=self.ACCENT,
            font=("Segoe UI", 9), width=50,
        ).pack(side="left", padx=(10, 8), ipady=4)

        tk.Button(
            folder_bar, text="Selecionar…", command=self._pick_folder,
            bg="#1e2a1a", fg=self.ACCENT,
            activebackground="#253220", activeforeground=self.ACCENT,
            bd=0, relief="flat", cursor="hand2",
            font=("Segoe UI", 9), padx=12, pady=4,
        ).pack(side="left")

        
        btn_bar = tk.Frame(self, bg=bg)
        btn_bar.pack(fill="x", padx=30, pady=(14, 0))

        self._start_btn = tk.Button(
            btn_bar, text="▶  Iniciar Envio",
            command=self._on_start,
            bg="#1e2a1a", fg=self.ACCENT,
            activebackground="#253220", activeforeground=self.ACCENT,
            bd=0, relief="flat", cursor="hand2",
            font=("Segoe UI", 10, "bold"), padx=20, pady=8,
        )
        self._start_btn.pack(side="left")

     
        log_frame = tk.Frame(self, bg=bg)
        log_frame.pack(fill="both", expand=True, padx=30, pady=(18, 20))

        tk.Label(log_frame, text="Log da execução", bg=bg, fg="#555555",
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", pady=(0, 4))

        text_wrap = tk.Frame(log_frame, bg="#161616",
                             highlightthickness=1, highlightbackground="#2a2a2a")
        text_wrap.pack(fill="both", expand=True)

        self._log = tk.Text(
            text_wrap, bg="#161616", fg="#c0c0c0",
            insertbackground="#c0c0c0", relief="flat", bd=0,
            font=("Consolas", 9), state="disabled", wrap="word",
            padx=12, pady=10, cursor="arrow",
        )
        scroll = ttk.Scrollbar(text_wrap, orient="vertical", command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

        self._log.tag_configure("info",    foreground="#8c8c8c")
        self._log.tag_configure("ok",      foreground="#2ecc71")
        self._log.tag_configure("warn",    foreground="#e8a020")
        self._log.tag_configure("err",     foreground="#e74c3c")
        self._log.tag_configure("heading", foreground="#f2f2f2",
                                font=("Consolas", 9, "bold"))

    def _pick_folder(self):
        d = filedialog.askdirectory(title="Selecione a pasta Boletos")
        if d:
            self._folder.set(d)

    def _go_back(self):
        self._cancel_requested = True
        for b in list(self._browser_ref):
            try:
                b.close()
            except Exception:
                pass
        self._browser_ref.clear()
        self.controller.show_frame("Home")

    def _log_write(self, text: str, tag: str = "info"):
        def _do():
            try:
                self._log.configure(state="normal")
                self._log.insert("end", text + "\n", tag)
                self._log.configure(state="disabled")
                self._log.see("end")
            except Exception:
                pass
        self._ui(_do)

    def _log_clear(self):
        def _do():
            try:
                self._log.configure(state="normal")
                self._log.delete("1.0", "end")
                self._log.configure(state="disabled")
            except Exception:
                pass
        self._ui(_do)

    def _on_start(self):
        pasta = self._folder.get().strip()
        if not pasta:
            messagebox.showwarning("Pasta", "Selecione a pasta Boletos antes de iniciar.")
            return
        if not os.path.isdir(pasta):
            messagebox.showerror("Pasta", "Caminho inválido ou não encontrado.")
            return
        if self._worker_running:
            messagebox.showinfo("Aguarde", "Uma rotina já está em execução.")
            return

        self._cancel_requested = False
        self._start_btn.configure(
            state="disabled", text="⏳  Executando…",
            fg="#555555", cursor="no", bg="#151515",
        )
        self._log_clear()
        self._log_write("Iniciando rotina…", "heading")
        self._log_write(f"Pasta: {pasta}", "info")

        t = threading.Thread(target=self._worker, args=(pasta,), daemon=True)
        self._worker_thread = t
        t.start()

    def _worker(self, pasta: str):
        self._worker_running = True

        def log(msg: str, tag: str = "info"):
            self._log_write(msg, tag)

        def cancelado() -> bool:
            return self._cancel_requested

        try:
            executar_rotina_taxas(pasta, log, cancelado, self._browser_ref)
        except Exception as e:
            if not self._cancel_requested:
                log(f"\nERRO INESPERADO: {e}", "err")
        finally:
            self._worker_running = False
            self._browser_ref.clear()
            self._ui(self._reset_start_btn)

    def _reset_start_btn(self):
        try:
            self._start_btn.configure(
                state="normal", text="▶  Iniciar Envio",
                fg=self.ACCENT, cursor="hand2", bg="#1e2a1a",
            )
        except Exception:
            pass


class HomeFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, style="Root.TFrame")
        self.controller = controller
        self.COLORS = controller.COLORS
        self._build_ui()

    def _build_ui(self):
        inner = ttk.Frame(self, style="Root.TFrame")
        inner.pack(fill="both", expand=True, padx=40, pady=40)

        self.logo_label = ttk.Label(inner, style="Root.TLabel")
        self.logo_label.pack(pady=(0, 14))
        lp = resource_path(LOGO_FILENAME)
        if os.path.exists(lp):
            try:
                img  = tk.PhotoImage(file=lp)
                img2 = img.subsample(2, 2)
                self._logo_ref = img2
                self.logo_label.configure(image=img2)
            except Exception:
                self.logo_label.configure(text="Logo")

        row = ttk.Frame(inner, style="Root.TFrame")
        row.pack(fill="x")
        ttk.Button(
            row, text="Taxas Large",
            style="Taxa.TButton",
            command=lambda: self.controller.show_frame("TaxasLarge"),
        ).pack(fill="x")

        ttk.Label(
            inner,
            text="Rotinas - Mesa de Operação Risco Sacado",
            style="Sub.TLabel",
        ).pack(pady=(14, 0))


class App(tk.Tk):
    COLORS = {
        "bg":     "#121212",
        "panel":  "#1a1a1a",
        "text":   "#f2f2f2",
        "muted":  "#1f1f1f",
        "border": "#2e2e2e",
        "accent": "#e8a020",
    }

    def __init__(self):
        super().__init__()
        self.title("Mesa - Itaú · Taxas Large")
        self.geometry("800x680")
        self.configure(bg=self.COLORS["bg"])
        self._setup_styles()

        container = ttk.Frame(self, style="Root.TFrame")
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.frames: dict[str, ttk.Frame] = {}
        for Cls, name in (
            (HomeFrame,      "Home"),
            (TaxasLargeFrame, "TaxasLarge"),
        ):
            f = Cls(container, self)
            self.frames[name] = f
            f.grid(row=0, column=0, sticky="nsew")

        self.show_frame("Home")

    def _setup_styles(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        c = self.COLORS
        s.configure("Root.TFrame",  background=c["bg"])
        s.configure("Root.TLabel",  background=c["bg"],    foreground=c["text"])
        s.configure("Dark.TFrame",  background=c["panel"])
        s.configure("Dark.TLabel",  background=c["panel"], foreground=c["text"])
        s.configure("Sub.TLabel",   background=c["bg"],    foreground=c["text"])
        s.configure("Dark.TButton", background=c["panel"], foreground=c["text"])
        s.map("Dark.TButton", background=[("active", c["border"])])
        s.configure(
            "Taxa.TButton",
            background="#1e2a1a",
            foreground="#e8a020",
            font=("Segoe UI", 11, "bold"),
        )
        s.map(
            "Taxa.TButton",
            background=[("active", "#253220")],
            foreground=[("active", "#f0b840")],
        )
        s.configure("Dark.TSeparator", background=c["border"])

    def show_frame(self, name: str):
        f = self.frames.get(name)
        if f:
            f.tkraise()
            if hasattr(f, "on_show"):
                try:
                    f.on_show()
                except Exception:
                    pass


if __name__ == "__main__":
    App().mainloop()