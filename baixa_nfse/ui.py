import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

import customtkinter as ctk
from PIL import Image

from .config import PASTA_SAIDA_PADRAO, TIPO_DOWNLOAD, URL_LOGIN
from .planilha import ler_planilha
from .automacao import (
    configurar_chrome, mudar_pasta_download, aguardar_download,
    criar_pasta_empresa, aguardar_login, navegar_para_recebidas,
    preencher_filtro_data, contar_notas, baixar_todas_as_notas,
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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

        logo_path = os.path.join(_BASE_DIR, "nfse-logo.png")
        try:
            pil_img = Image.open(logo_path)
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

            pasta = criar_pasta_empresa(self.pasta_saida.get(), emp["nome"], emp["cnpj"])
            self.log(f"  Pasta: {pasta}")
            mudar_pasta_download(self.driver, pasta)

            self.driver.get(URL_LOGIN)
            self.log("  Selecione o certificado digital no navegador.")
            self.log("  O script continuara automaticamente apos o login.")

            login_ok = aguardar_login(self.driver, self.log)
            if not login_ok or not self.rodando:
                self.log("  Login nao detectado. Pulando empresa.")
                self.empresa_atual_idx += 1
                continue

            navegar_para_recebidas(self.driver, self.log)

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
