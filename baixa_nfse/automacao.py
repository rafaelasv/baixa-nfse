import os
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from .config import URL_RECEBIDAS, TIMEOUT_LOGIN


def sanitizar(nome: str) -> str:
    for c in r'\/:*?"<>|':
        nome = nome.replace(c, "_")
    return nome.strip()


def criar_pasta_empresa(pasta_raiz: str, nome: str, cnpj: str) -> str:
    caminho = os.path.join(pasta_raiz, sanitizar(f"{nome}_{cnpj}"))
    os.makedirs(caminho, exist_ok=True)
    return caminho


def configurar_chrome(pasta_download: str) -> webdriver.Chrome:
    opcoes = Options()
    prefs = {
        "download.default_directory": os.path.abspath(pasta_download),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    opcoes.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=opcoes)
    driver.maximize_window()
    return driver


def mudar_pasta_download(driver, pasta: str):
    driver.execute_cdp_cmd(
        "Browser.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": os.path.abspath(pasta)},
    )


def aguardar_download(pasta: str, timeout: int = 120) -> bool:
    inicio = time.time()
    while time.time() - inicio < timeout:
        if not any(f.endswith(".crdownload") for f in os.listdir(pasta)):
            return True
        time.sleep(1)
    return False


def aguardar_login(driver, log_fn) -> bool:
    log_fn("  Aguardando selecao do certificado e login...")
    inicio = time.time()
    while time.time() - inicio < TIMEOUT_LOGIN:
        try:
            src = driver.page_source
            if ("Meus dados" in src or "Rascunhos" in src) and "EmissorNacional" in driver.current_url:
                log_fn("  Login detectado automaticamente!")
                return True
        except Exception:
            pass
        time.sleep(1)
    log_fn("  Timeout: login nao detectado em 120s.")
    return False


def navegar_para_recebidas(driver, log_fn):
    log_fn("  Navegando para Notas Recebidas...")
    driver.get(URL_RECEBIDAS)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "datainicio"))
        )
        log_fn("  Pagina de Notas Recebidas carregada.")
        return True
    except TimeoutException:
        log_fn("  Pagina carregou mas campo de data nao encontrado.")
        return False


def preencher_filtro_data(driver, data_inicio: str, data_fim: str, log_fn) -> bool:
    log_fn(f"  Preenchendo datas: {data_inicio} a {data_fim}...")
    wait = WebDriverWait(driver, 15)
    try:
        campo_ini = wait.until(EC.presence_of_element_located((By.ID, "datainicio")))
        campo_fim = driver.find_element(By.ID, "datafim")

        for campo, valor in [(campo_ini, data_inicio), (campo_fim, data_fim)]:
            driver.execute_script("""
                var campo = arguments[0];
                var valor = arguments[1];
                campo.value = valor;
                campo.dispatchEvent(new Event('input',  {bubbles: true}));
                campo.dispatchEvent(new Event('change', {bubbles: true}));
                campo.dispatchEvent(new Event('blur',   {bubbles: true}));
            """, campo, valor)
            time.sleep(0.3)

        btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn-primary[type='submit']")
        ))
        btn.click()
        log_fn("  Filtro aplicado. Aguardando resultados...")
        time.sleep(3)
        return True

    except Exception as e:
        log_fn(f"  Erro ao preencher datas: {e}")
        return False


def contar_notas(driver) -> int:
    try:
        texto = driver.find_element(
            By.XPATH, "//*[contains(text(),'Total de') and contains(text(),'registro')]"
        ).text
        for p in texto.split():
            if p.isdigit():
                return int(p)
    except Exception:
        pass
    return 0


def baixar_todas_as_notas(driver, tipo: str, log_fn):
    wait = WebDriverWait(driver, 15)
    pagina = 1

    while True:
        log_fn(f"  Pagina {pagina}...")

        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "table tbody tr")
            ))
            time.sleep(1)
        except TimeoutException:
            log_fn("  Nenhuma nota encontrada.")
            break

        menus = driver.find_elements(By.CSS_SELECTOR, "div.menu-suspenso-tabela")
        total = len(menus)
        log_fn(f"  {total} nota(s) nesta pagina.")

        for i in range(total):
            try:
                menus = driver.find_elements(By.CSS_SELECTOR, "div.menu-suspenso-tabela")
                if i >= len(menus):
                    break
                menu = menus[i]

                icone = menu.find_element(By.CSS_SELECTOR, "a.icone-trigger")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", icone)
                time.sleep(0.3)
                icone.click()
                time.sleep(0.6)

                trecho_url = "/Download/NFSe/" if tipo == "XML" else "/Download/DANFSe/"

                wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.popover-content")
                ))
                time.sleep(0.3)

                popover = driver.find_element(By.CSS_SELECTOR, "div.popover.in div.popover-content")
                links = popover.find_elements(By.TAG_NAME, "a")

                clicou = False
                for link in links:
                    href = link.get_attribute("href") or ""
                    if trecho_url in href:
                        link.click()
                        clicou = True
                        time.sleep(2)
                        log_fn(f"    Nota {i+1}/{total}: download iniciado.")
                        break

                if not clicou:
                    hrefs = [l.get_attribute("href") or l.text.strip() for l in links]
                    log_fn(f"    Nota {i+1}/{total}: link '{trecho_url}' nao encontrado.")
                    log_fn(f"    Links disponiveis: {hrefs}")
                    driver.find_element(By.TAG_NAME, "body").click()
                    time.sleep(0.3)

            except Exception as e:
                log_fn(f"    Erro na nota {i+1}: {e}")
                try:
                    driver.find_element(By.TAG_NAME, "body").click()
                    time.sleep(0.3)
                except Exception:
                    pass

        try:
            proximo = driver.find_element(
                By.XPATH,
                "//li[not(contains(@class,'disabled'))]/a[@aria-label='Próxima página'] | "
                "//li[not(contains(@class,'disabled'))]/a[@aria-label='Proxima pagina'] | "
                "//li[not(contains(@class,'disabled'))]/a[text()='›'] | "
                "//li[not(contains(@class,'disabled'))]/a[text()='»']"
            )
            proximo.click()
            pagina += 1
            time.sleep(2)
        except NoSuchElementException:
            log_fn(f"  Fim das paginas (total: {pagina}).")
            break
