"""
Automacao NFS-e - Portal Contribuinte  v3.0
Baixa XMLs ou DANFS-e de notas recebidas para multiplas empresas.

Dependencias:
    pip install selenium openpyxl customtkinter

Estrutura da planilha (.xlsx):
    Linha 1: cabecalho (ignorado)
    Coluna A: Nome da empresa
    Coluna B: CNPJ
"""

import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

import customtkinter as ctk
from PIL import Image

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

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

TIPO_DOWNLOAD = "XML"   # "xml" ou "pdf"

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
    texto_opcao = "Download XML" if tipo == "XML" else "Download DANFS-e"
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
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("NFS-e  ·  Download Automático")
        self.root.geometry("660x620")
        self.root.resizable(False, False)

        self.caminho_planilha = tk.StringVar()
        self.pasta_saida      = tk.StringVar(value=PASTA_SAIDA_PADRAO)
        self.data_inicio      = tk.StringVar(value=f"01/01/{time.strftime('%Y')}")
        self.data_fim         = tk.StringVar(value=f"31/01/{time.strftime('%Y')}")
        self.tipo_download    = tk.StringVar(value=TIPO_DOWNLOAD)

        self.empresas = []
        self.empresa_atual_idx = 0
        self.driver   = None
        self.rodando  = False

        self._construir_tela()

    def _construir_tela(self):
        # ── Cabeçalho ──────────────────────────────────────────
        header = ctk.CTkFrame(self.root, fg_color=("gray90", "#111827"), corner_radius=0)
        header.pack(fill="x")

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nfse-logo.png")
        try:
            pil_img = Image.open(logo_path)
            # mantém proporção baseando na altura desejada de 44px
            h = 44
            w = int(pil_img.width * h / pil_img.height)
            logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(w, h))
            ctk.CTkLabel(header, image=logo_img, text="").pack(side="left", padx=(18, 10), pady=12)
        except Exception:
            ctk.CTkLabel(
                header, text="NFS-e", font=ctk.CTkFont("Segoe UI", 22, "bold"),
                text_color="#3b82f6"
            ).pack(side="left", padx=(20, 4), pady=14)

        ctk.CTkLabel(
            header, text="Automação de Download",
            font=ctk.CTkFont("Segoe UI", 13), text_color=("gray40", "gray60")
        ).pack(side="left", pady=14)
        ctk.CTkLabel(
            header, text="v3.0",
            font=ctk.CTkFont("Segoe UI", 10), text_color=("gray60", "gray50")
        ).pack(side="right", padx=20, pady=14)

        # ── Corpo ──────────────────────────────────────────────
        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(16, 0))

        # Card: Arquivos
        card1 = ctk.CTkFrame(body, corner_radius=10)
        card1.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            card1, text="ARQUIVOS",
            font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=("gray50", "gray50")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 4))

        ctk.CTkLabel(card1, text="Planilha", font=ctk.CTkFont("Segoe UI", 12),
                     anchor="w").grid(row=1, column=0, padx=(14, 8), pady=6, sticky="w")
        ctk.CTkEntry(card1, textvariable=self.caminho_planilha,
                     width=440, placeholder_text="Selecione o arquivo .xlsx…"
                     ).grid(row=1, column=1, pady=6, sticky="ew")
        ctk.CTkButton(card1, text="···", width=40, command=self._sel_planilha,
                      fg_color=("#3b82f6", "#1d4ed8"), hover_color=("#2563eb", "#1e40af")
                      ).grid(row=1, column=2, padx=(6, 14), pady=6)

        ctk.CTkLabel(card1, text="Salvar em", font=ctk.CTkFont("Segoe UI", 12),
                     anchor="w").grid(row=2, column=0, padx=(14, 8), pady=(0, 10), sticky="w")
        ctk.CTkEntry(card1, textvariable=self.pasta_saida,
                     width=440, placeholder_text="Pasta de destino…"
                     ).grid(row=2, column=1, pady=(0, 10), sticky="ew")
        ctk.CTkButton(card1, text="···", width=40, command=self._sel_pasta,
                      fg_color=("#3b82f6", "#1d4ed8"), hover_color=("#2563eb", "#1e40af")
                      ).grid(row=2, column=2, padx=(6, 14), pady=(0, 10))

        card1.columnconfigure(1, weight=1)

        # Card: Período
        card2 = ctk.CTkFrame(body, corner_radius=10)
        card2.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            card2, text="PERÍODO  &  TIPO",
            font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=("gray50", "gray50")
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=14, pady=(10, 4))

        ctk.CTkLabel(card2, text="De", font=ctk.CTkFont("Segoe UI", 12)
                     ).grid(row=1, column=0, padx=(14, 6), pady=(0, 12))
        ctk.CTkEntry(card2, textvariable=self.data_inicio, width=110
                     ).grid(row=1, column=1, pady=(0, 12))
        ctk.CTkLabel(card2, text="Até", font=ctk.CTkFont("Segoe UI", 12)
                     ).grid(row=1, column=2, padx=(14, 6), pady=(0, 12))
        ctk.CTkEntry(card2, textvariable=self.data_fim, width=110
                     ).grid(row=1, column=3, pady=(0, 12))
        ctk.CTkLabel(card2, text="Tipo", font=ctk.CTkFont("Segoe UI", 12)
                     ).grid(row=1, column=4, padx=(20, 6), pady=(0, 12))
        ctk.CTkComboBox(card2, variable=self.tipo_download,
                        values=["XML", "PDF"], width=80, state="readonly"
                        ).grid(row=1, column=5, padx=(0, 14), pady=(0, 12))

        # ── Botões ─────────────────────────────────────────────
        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x", pady=(4, 10))

        self.btn_iniciar = ctk.CTkButton(
            btn_row, text="  Iniciar Processo  ",
            font=ctk.CTkFont("Segoe UI", 13, "bold"),
            height=40, corner_radius=8,
            fg_color=("#16a34a", "#15803d"), hover_color=("#15803d", "#166534"),
            command=self._iniciar
        )
        self.btn_iniciar.pack(side="left", padx=(0, 10))

        self.btn_parar = ctk.CTkButton(
            btn_row, text="Parar",
            font=ctk.CTkFont("Segoe UI", 12),
            height=40, corner_radius=8, width=100,
            fg_color=("#dc2626", "#b91c1c"), hover_color=("#b91c1c", "#991b1b"),
            state="disabled", command=self._parar
        )
        self.btn_parar.pack(side="left")

        # ── Status + barra de progresso ────────────────────────
        self.lbl_status = ctk.CTkLabel(
            body, text="● Pronto",
            font=ctk.CTkFont("Segoe UI", 11), text_color=("gray50", "gray50"),
            anchor="w"
        )
        self.lbl_status.pack(fill="x", pady=(0, 4))

        self.progress = ctk.CTkProgressBar(body, height=4, corner_radius=2)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(0, 10))

        # ── Log ────────────────────────────────────────────────
        self.txt_log = ctk.CTkTextbox(
            body,
            font=ctk.CTkFont("Consolas", 10),
            fg_color=("#f8f9fa", "#0d1117"),
            text_color=("#374151", "#e2e8f0"),
            corner_radius=8,
            state="disabled",
            wrap="word"
        )
        self.txt_log.pack(fill="both", expand=True, pady=(0, 16))

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
        self.root.after(0, lambda: self.lbl_status.configure(text=f"● {msg}"))

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
                                         "Verifique se há dados a partir da linha 3.")
            return

        self.log(f"Planilha carregada: {len(self.empresas)} empresa(s).")
        self.empresa_atual_idx = 0
        self.rodando = True
        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.driver = configurar_chrome(self.pasta_saida.get())
        Thread(target=self._loop_empresas, daemon=True).start()

    def _loop_empresas(self):
        total = len(self.empresas)

        while self.empresa_atual_idx < total and self.rodando:
            emp = self.empresas[self.empresa_atual_idx]
            idx = self.empresa_atual_idx + 1
            self.status(f"Empresa {idx}/{total}: {emp['nome']}")
            self.log(f"\n── Empresa {idx}/{total}: {emp['nome']} ({emp['cnpj']}) ──")

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
            self.status("Concluído.")
        self._finalizar()

    def _parar(self):
        self.rodando = False
        self.log("\nInterrompido pelo usuário.")
        self._finalizar()

    def _finalizar(self):
        def _u():
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(0)
        self.root.after(0, _u)
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None


if __name__ == "__main__":
    app = App()
    app.root.mainloop()
