"""
Automacao NFS-e - Portal Contribuinte  v3.0
Baixa XMLs ou DANFS-e de notas recebidas para multiplas empresas.

Dependencias:
    pip install selenium openpyxl

Estrutura da planilha (.xlsx):
    Linha 1: cabecalho (ignorado)
    Coluna A: Nome da empresa
    Coluna B: CNPJ
"""

import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES
# ─────────────────────────────────────────────

URL_LOGIN     = "https://www.nfse.gov.br/EmissorNacional/"
URL_RECEBIDAS = "https://www.nfse.gov.br/EmissorNacional/Notas/Recebidas"

TIPO_DOWNLOAD = "xml"   # "xml" ou "pdf"

PASTA_SAIDA_PADRAO = os.path.join(os.path.expanduser("~"), "Downloads", "Downloads_XML")

# Tempo maximo (segundos) esperando o usuario selecionar o certificado
TIMEOUT_LOGIN = 120


# ─────────────────────────────────────────────
#  FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────

def ler_planilha(caminho: str) -> list:
    wb = openpyxl.load_workbook(caminho)
    ws = wb.active
    empresas = []
    for row in ws.iter_rows(min_row=3, values_only=True):  # pula titulo e cabecalho
        nome = str(row[1]).strip() if row[1] else None  # coluna B = Nome
        cnpj = str(row[2]).strip() if row[2] else None  # coluna C = CNPJ
        if nome and cnpj:
            empresas.append({"nome": nome, "cnpj": cnpj})
    return empresas


def sanitizar(nome: str) -> str:
    for c in r'\/:*?"<>|':
        nome = nome.replace(c, "_")
    return nome.strip()


def criar_pasta_empresa(pasta_raiz: str, nome: str, cnpj: str) -> str:
    """Cria e retorna a pasta  PastaRaiz/NomeEmpresa_CNPJ/"""
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


# ─────────────────────────────────────────────
#  AUTOMAÇÃO DO SITE
# ─────────────────────────────────────────────

def aguardar_login(driver, log_fn) -> bool:
    """
    Aguarda automaticamente o login ser concluido.
    Detecta quando a pagina inicial (Home) carregou apos selecionar o certificado.
    Sinal de login OK: elemento 'Meus dados' ou breadcrumb 'Home' aparecer.
    """
    log_fn("  Aguardando selecao do certificado e login...")
    inicio = time.time()
    while time.time() - inicio < TIMEOUT_LOGIN:
        try:
            # Pagina inicial tem "Meus dados" ou "Rascunhos" no corpo
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
    """Navega direto pela URL — sem depender de clicar em menus."""
    log_fn("  Navegando para Notas Recebidas...")
    driver.get(URL_RECEBIDAS)
    # Aguarda o titulo da pagina aparecer
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
    """
    Preenche os campos pelo id exato: datainicio e datafim.
    Usa JavaScript para setar o valor direto — evita problemas com mascara.
    Depois dispara os eventos necessarios para o site reconhecer a mudanca.
    """
    log_fn(f"  Preenchendo datas: {data_inicio} a {data_fim}...")
    wait = WebDriverWait(driver, 15)

    try:
        campo_ini = wait.until(EC.presence_of_element_located((By.ID, "datainicio")))
        campo_fim = driver.find_element(By.ID, "datafim")

        # Seta o valor via JavaScript e dispara evento 'change' e 'input'
        # Isso garante que a mascara e validacao do site reconhecam o valor
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

        # Clica no botao Filtrar
        # O botao e type="submit" class="btn btn-lg btn-primary" dentro de um <form>
        # O texto "Filtrar" fica num <span> filho, por isso usamos o type e class
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
    """
    Percorre todas as paginas e baixa XML ou DANFS-e de cada nota.

    Estrutura real do site (confirmada via DevTools):
      Botao tres pontinhos: <a class="icone-trigger"> dentro de <div class="menu-suspenso-tabela">
      Dropdown:             <div class="list-group menu-content">
      Opcoes:               <a> dentro do dropdown com o texto da acao
    """
    texto_opcao = "Download XML" if tipo == "xml" else "Download DANFS-e"
    wait = WebDriverWait(driver, 15)
    pagina = 1

    while True:
        log_fn(f"  Pagina {pagina}...")

        # Aguarda a tabela
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "table tbody tr")
            ))
            time.sleep(1)
        except TimeoutException:
            log_fn("  Nenhuma nota encontrada.")
            break

        # Conta quantos menus suspensos tem na pagina (= quantas notas)
        menus = driver.find_elements(By.CSS_SELECTOR, "div.menu-suspenso-tabela")
        total = len(menus)
        log_fn(f"  {total} nota(s) nesta pagina.")

        for i in range(total):
            try:
                # Re-busca a cada iteracao (DOM pode mudar)
                menus = driver.find_elements(By.CSS_SELECTOR, "div.menu-suspenso-tabela")
                if i >= len(menus):
                    break
                menu = menus[i]

                # Clica no icone de tres pontinhos para abrir o dropdown
                icone = menu.find_element(By.CSS_SELECTOR, "a.icone-trigger")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", icone)
                time.sleep(0.3)
                icone.click()
                time.sleep(0.6)

                # O dropdown e um popover — procura o link pela URL (mais confiavel que texto)
                # XML:  href contém /Download/NFSe/
                # PDF:  href contém /Download/DANFSe/
                trecho_url = "/Download/NFSe/" if tipo == "xml" else "/Download/DANFSe/"

                # Aguarda o popover aparecer e ficar visivel
                wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.popover-content")
                ))
                time.sleep(0.3)

                # Pega todos os links visiveis no popover ativo
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

        # Verifica proxima pagina
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


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Automacao NFS-e Nacional")
        self.root.geometry("620x560")
        self.root.resizable(False, False)
        self.root.configure(bg="#2b2b2b")

        self.caminho_planilha = tk.StringVar()
        self.pasta_saida      = tk.StringVar(value=PASTA_SAIDA_PADRAO)
        self.data_inicio      = tk.StringVar(value=f"01/01/{time.strftime('%Y')}")
        self.data_fim         = tk.StringVar(value=f"31/01/{time.strftime('%Y')}")
        self.tipo_download    = tk.StringVar(value=TIPO_DOWNLOAD)

        self.empresas = []
        self.empresa_atual_idx = 0
        self.driver   = None
        self.rodando  = False
        self._confirmacao_ok = False

        self._construir_tela()

    def _construir_tela(self):
        BG    = "#2b2b2b"
        BG2   = "#3c3f41"
        FG    = "#ffffff"
        VERDE = "#00cc44"
        AZUL  = "#3d8bcd"

        tk.Label(self.root, text="Automação de Download - NFS-e",
                 bg=BG, fg=VERDE, font=("Consolas", 14, "bold")).pack(pady=(15, 5))

        # Planilha
        f = tk.Frame(self.root, bg=BG2, pady=6, padx=10)
        f.pack(fill="x", padx=15, pady=4)
        tk.Label(f, text="Planilha (.xlsx):", bg=BG2, fg=FG, width=16, anchor="w").pack(side="left")
        tk.Entry(f, textvariable=self.caminho_planilha, width=42,
                 bg="#1e1e1e", fg=FG, insertbackground=FG).pack(side="left", padx=4)
        tk.Button(f, text="...", command=self._sel_planilha,
                  bg=AZUL, fg=FG, width=3).pack(side="left")

        # Pasta de saida
        f2 = tk.Frame(self.root, bg=BG2, pady=6, padx=10)
        f2.pack(fill="x", padx=15, pady=4)
        tk.Label(f2, text="Salvar em:", bg=BG2, fg=FG, width=16, anchor="w").pack(side="left")
        tk.Entry(f2, textvariable=self.pasta_saida, width=42,
                 bg="#1e1e1e", fg=FG, insertbackground=FG).pack(side="left", padx=4)
        tk.Button(f2, text="...", command=self._sel_pasta,
                  bg=AZUL, fg=FG, width=3).pack(side="left")

        # Datas + tipo
        f3 = tk.Frame(self.root, bg=BG2, pady=6, padx=10)
        f3.pack(fill="x", padx=15, pady=4)
        tk.Label(f3, text="Data Início:", bg=BG2, fg=FG).pack(side="left")
        tk.Entry(f3, textvariable=self.data_inicio, width=12,
                 bg="#1e1e1e", fg=FG, insertbackground=FG).pack(side="left", padx=4)
        tk.Label(f3, text="Data Fim:", bg=BG2, fg=FG).pack(side="left", padx=(10, 0))
        tk.Entry(f3, textvariable=self.data_fim, width=12,
                 bg="#1e1e1e", fg=FG, insertbackground=FG).pack(side="left", padx=4)
        tk.Label(f3, text="Tipo:", bg=BG2, fg=FG).pack(side="left", padx=(15, 0))
        ttk.Combobox(f3, textvariable=self.tipo_download,
                     values=["xml", "pdf"], width=5, state="readonly").pack(side="left", padx=4)

        # Botoes
        fb = tk.Frame(self.root, bg=BG, pady=8)
        fb.pack()

        self.btn_iniciar = tk.Button(
            fb, text="INICIAR PROCESSO",
            bg=VERDE, fg="#000000", font=("Consolas", 11, "bold"),
            padx=15, command=self._iniciar
        )
        self.btn_iniciar.pack(side="left", padx=6)

        self.btn_parar = tk.Button(
            fb, text="Parar",
            bg="#cc3333", fg=FG, font=("Consolas", 10, "bold"),
            padx=10, command=self._parar, state="disabled"
        )
        self.btn_parar.pack(side="left", padx=6)

        # Status
        self.lbl_status = tk.Label(self.root, text="Pronto.", bg=BG, fg="#aaaaaa",
                                   font=("Consolas", 9))
        self.lbl_status.pack()

        # Log
        fl = tk.Frame(self.root, bg=BG)
        fl.pack(fill="both", expand=True, padx=15, pady=(0, 8))
        self.txt_log = tk.Text(
            fl, bg="#1a1a1a", fg="#00ff66",
            font=("Consolas", 9), state="disabled", wrap="word"
        )
        scroll = tk.Scrollbar(fl, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt_log.pack(fill="both", expand=True)

        # Rodape
        tk.Label(
            self.root,
            text=f"Versao 3.0  |  Salvará em: {os.path.basename(PASTA_SAIDA_PADRAO)}",
            bg="#1a1a1a", fg="#666666", font=("Consolas", 8)
        ).pack(fill="x", side="bottom")

    def _sel_planilha(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.caminho_planilha.set(p)

    def _sel_pasta(self):
        p = filedialog.askdirectory()
        if p:
            self.pasta_saida.set(p)

    def log(self, msg: str):
        def _a():
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.root.after(0, _a)

    def status(self, msg: str):
        self.root.after(0, lambda: self.lbl_status.config(text=msg))

    def _iniciar(self):
        if not self.caminho_planilha.get():
            messagebox.showerror("Erro", "Selecione a planilha de empresas.")
            return
        try:
            self.empresas = ler_planilha(self.caminho_planilha.get())
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            return
        if not self.empresas:
            messagebox.showerror("Erro", "Nenhuma empresa encontrada.\n"
                                         "Verifique se há dados a partir da linha 2.")
            return

        self.log(f"Planilha carregada: {len(self.empresas)} empresa(s).")
        self.empresa_atual_idx = 0
        self.rodando = True
        self.btn_iniciar.config(state="disabled")
        self.btn_parar.config(state="normal")
        self.driver = configurar_chrome(self.pasta_saida.get())
        Thread(target=self._loop_empresas, daemon=True).start()

    def _loop_empresas(self):
        total = len(self.empresas)

        while self.empresa_atual_idx < total and self.rodando:
            emp = self.empresas[self.empresa_atual_idx]
            idx = self.empresa_atual_idx + 1
            self.status(f"Empresa {idx}/{total}: {emp['nome']}")
            self.log(f"\n-- Empresa {idx}/{total}: {emp['nome']} ({emp['cnpj']}) --")

            # Cria pasta e aponta download para ela
            pasta = criar_pasta_empresa(self.pasta_saida.get(), emp["nome"], emp["cnpj"])
            self.log(f"  Pasta: {pasta}")
            mudar_pasta_download(self.driver, pasta)

            # Abre o site e aguarda login automaticamente
            self.driver.get(URL_LOGIN)
            self.log("  Selecione o certificado digital no navegador.")
            self.log("  O script continuara automaticamente apos o login.")

            login_ok = aguardar_login(self.driver, self.log)
            if not login_ok or not self.rodando:
                self.log("  Login nao detectado. Pulando empresa.")
                self.empresa_atual_idx += 1
                continue

            # Navega direto para Recebidas
            navegar_para_recebidas(self.driver, self.log)

            # Preenche datas e filtra
            filtrou = preencher_filtro_data(
                self.driver,
                self.data_inicio.get(),
                self.data_fim.get(),
                self.log
            )
            if not filtrou:
                self.log("  Falha no filtro. Pulando empresa.")
                self.empresa_atual_idx += 1
                continue

            # Conta e baixa
            total_notas = contar_notas(self.driver)
            self.log(f"  Total de registros: {total_notas if total_notas else 'nao detectado'}")

            baixar_todas_as_notas(self.driver, self.tipo_download.get(), self.log)
            aguardar_download(pasta)
            self.log(f"  Concluido: {emp['nome']}")
            self.empresa_atual_idx += 1

        if self.rodando:
            self.log("\nTodas as empresas foram processadas!")
            self.status("Concluido.")
        self._finalizar()

    def _parar(self):
        self.rodando = False
        self.log("\nInterrompido pelo usuário.")
        self._finalizar()

    def _finalizar(self):
        def _u():
            self.btn_iniciar.config(state="normal")
            self.btn_parar.config(state="disabled")
        self.root.after(0, _u)
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
