from playwright.sync_api import sync_playwright
import time

URL = ""
usuario = ""
senha = ""
CLIENTES_FILE = "clientes.txt"
LOG_FILE = "bmflimites_log.txt"

with open(CLIENTES_FILE, "r", encoding="utf-8") as f:
    raw = f.read().replace("\n", ",")
    CLIENTES = [c.strip() for c in raw.split(",") if c.strip()]

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(URL)

      
        page.fill('input#txtBoxUsuario', usuario)
        page.fill('input#txtBoxSenha', senha)
        page.click('input#btnEntrar')
        time.sleep(2)

        with open(LOG_FILE, "a", encoding="utf-8") as log:
            for cliente in CLIENTES:
                status = "ok"
                try:
                    page.fill('input#txtBoxCodCli_TextBox', cliente)
                    page.keyboard.press('Tab')
                    time.sleep(2)

                    
                    with page.expect_popup() as popup_info:
                        page.click('input#imgLimiteCli')
                    nova_aba = popup_info.value
                    nova_aba.bring_to_front()
                    time.sleep(2)

                    nova_aba.click('table#tab1')
                    time.sleep(1)

                    nova_aba.select_option('select#cboRegraBMF', value='312')
                    time.sleep(1)

                    is_checked = nova_aba.is_checked('input#chkCalculaLimiteFinanceiroBMF')
                    if is_checked:
                        nova_aba.uncheck('input#chkCalculaLimiteFinanceiroBMF')
                    time.sleep(1)

                    nova_aba.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    nova_aba.click('input#btnConfirmar')
                    time.sleep(2)

                    
                    try:
                        with nova_aba.expect_popup() as confirm_popup_info:
                            pass
                        confirm_aba = confirm_popup_info.value
                        confirm_aba.bring_to_front()
                        confirm_aba.fill('input#txtSenha', senha)
                        confirm_aba.click('input#btnOk')
                        time.sleep(2)
                    except Exception:
                      
                        confirm_pages = browser.contexts[0].pages
                        for p in confirm_pages:
                            if p != page and p != nova_aba:
                                try:
                                    p.fill('input#txtSenha', senha)
                                    p.click('input#btnOk')
                                    time.sleep(2)
                                except Exception:
                                    continue

                    
                    nova_aba.close()
                    page.bring_to_front()
                except Exception as e:
                    status = f"erro: {str(e)}"
                log.write(f"Cliente {cliente}: {status}\n")
                log.flush()

        print("Pressione Enter para fechar o navegador...")
        input()
        browser.close()

if __name__ == "__main__":
    main() 
