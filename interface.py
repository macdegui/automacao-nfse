import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import json
import os
from processador import executar_automacao

# Arquivo para salvar configurações
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".nfse_config.json")

# Tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

AZUL_ESCURO   = "#0D1B2A"
AZUL_MEDIO    = "#1B2E45"
AZUL_CARD     = "#162236"
AZUL_BORDA    = "#1E3A5F"
AZUL_ACENTO   = "#1A6FD4"
AZUL_HOVER    = "#1559B0"
VERDE         = "#22C55E"
VERMELHO      = "#EF4444"
AMARELO       = "#F59E0B"
TEXTO         = "#E8F0FE"
TEXTO_MUTED   = "#7A9CC4"
BRANCO        = "#FFFFFF"

ALTURA_COMPACTA  = 460
ALTURA_EXPANDIDA = 700


class NFSeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação NFS-e")
        self.geometry(f"820x{ALTURA_COMPACTA}")
        self.minsize(720, 400)
        self.configure(fg_color=AZUL_ESCURO)
        self.resizable(True, True)

        self._processando = False
        self._log_visivel = False
        self._carregar_config()
        self._build_ui()

    # ── Persistência ──────────────────────────────────────────────────────────
    def _carregar_config(self):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                self._config = json.load(f)
        except Exception:
            self._config = {}

    def _salvar_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ── Header ────────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=AZUL_MEDIO, corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="  ⚙  Automação NFS-e",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=BRANCO
        ).grid(row=0, column=0, sticky="w", padx=24, pady=16)

        ctk.CTkLabel(
            header,
            text="Grupo Guerra · Organização de NFSe",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXTO_MUTED
        ).grid(row=0, column=1, sticky="e", padx=24)

        # ── Body (configurações + botão + progresso) ──────────────────────────
        body = ctk.CTkFrame(self, fg_color=AZUL_ESCURO, corner_radius=0)
        body.grid(row=1, column=0, sticky="ew", padx=0, pady=0)
        body.grid_columnconfigure(0, weight=1)

        # Card de configurações
        card = ctk.CTkFrame(body, fg_color=AZUL_CARD,
                            corner_radius=12,
                            border_width=1, border_color=AZUL_BORDA)
        card.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 8))
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            card, text="CONFIGURAÇÕES",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=TEXTO_MUTED
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=20, pady=(16, 4))

        sep = ctk.CTkFrame(card, fg_color=AZUL_BORDA, height=1)
        sep.grid(row=1, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 16))

        campos = [
            ("pasta_origem",  "📥  Pasta de Origem (Downloads NFSe)",  "Selecionar pasta de origem"),
            ("pasta_destino", "📤  Pasta de Destino (Servidor Z:)",     "Selecionar pasta do servidor"),
            ("config_excel",  "📋  Mapeamento de Empresas (.xlsx)",     "Selecionar arquivo Excel"),
        ]

        self._entries = {}
        for idx, (chave, label, placeholder) in enumerate(campos):
            row = idx * 2 + 2
            ctk.CTkLabel(
                card, text=label,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=TEXTO
            ).grid(row=row, column=0, sticky="w", padx=20, pady=(0, 2))

            entry = ctk.CTkEntry(
                card,
                placeholder_text=placeholder,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                fg_color=AZUL_MEDIO, border_color=AZUL_BORDA,
                text_color=TEXTO, placeholder_text_color=TEXTO_MUTED,
                height=36, corner_radius=8
            )
            entry.grid(row=row + 1, column=0, columnspan=2,
                       sticky="ew", padx=(20, 8), pady=(0, 12))

            valor_salvo = self._config.get(chave, "")
            if valor_salvo:
                entry.insert(0, valor_salvo)

            btn = ctk.CTkButton(
                card,
                text="📁",
                width=40, height=36,
                fg_color=AZUL_MEDIO, hover_color=AZUL_ACENTO,
                border_width=1, border_color=AZUL_BORDA,
                corner_radius=8,
                command=lambda c=chave, e=entry: self._selecionar_caminho(c, e)
            )
            btn.grid(row=row + 1, column=2, sticky="e", padx=(0, 20), pady=(0, 12))

            self._entries[chave] = entry

        ctk.CTkFrame(card, fg_color="transparent", height=8).grid(
            row=len(campos) * 2 + 2, column=0)

        # ── Botão Executar ────────────────────────────────────────────────────
        self._btn_executar = ctk.CTkButton(
            body,
            text="▶   Executar Automação",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color=AZUL_ACENTO, hover_color=AZUL_HOVER,
            height=46, corner_radius=10,
            command=self._iniciar_processamento
        )
        self._btn_executar.grid(row=1, column=0, sticky="ew",
                                padx=24, pady=(8, 4))

        # ── Barra de progresso ────────────────────────────────────────────────
        prog_frame = ctk.CTkFrame(body, fg_color="transparent")
        prog_frame.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 16))
        prog_frame.grid_columnconfigure(0, weight=1)

        self._barra = ctk.CTkProgressBar(
            prog_frame, height=8, corner_radius=4,
            fg_color=AZUL_MEDIO, progress_color=AZUL_ACENTO
        )
        self._barra.grid(row=0, column=0, sticky="ew")
        self._barra.set(0)

        self._label_progresso = ctk.CTkLabel(
            prog_frame, text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXTO_MUTED
        )
        self._label_progresso.grid(row=0, column=1, padx=(12, 0))

        # ── Log (oculto por padrão) ───────────────────────────────────────────
        self._log_frame = ctk.CTkFrame(self, fg_color=AZUL_CARD,
                                       corner_radius=12,
                                       border_width=1, border_color=AZUL_BORDA)
        self._log_frame.grid_columnconfigure(0, weight=1)
        self._log_frame.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(self._log_frame, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_header, text="LOG DE EXECUÇÃO",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=TEXTO_MUTED
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            log_header, text="Limpar",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color="transparent", hover_color=AZUL_MEDIO,
            text_color=TEXTO_MUTED, height=24, width=60,
            command=self._limpar_log
        ).grid(row=0, column=1, sticky="e", padx=(0, 4))

        ctk.CTkButton(
            log_header, text="▲ Recolher",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color="transparent", hover_color=AZUL_MEDIO,
            text_color=TEXTO_MUTED, height=24, width=80,
            command=self._ocultar_log
        ).grid(row=0, column=2, sticky="e")

        self._log_box = ctk.CTkTextbox(
            self._log_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=AZUL_ESCURO, text_color=TEXTO,
            border_width=0, corner_radius=0,
            wrap="word", state="disabled"
        )
        self._log_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        # Log começa oculto
        self._log_frame.grid_remove()

    # ── Log visibilidade ──────────────────────────────────────────────────────
    def _mostrar_log(self):
        if not self._log_visivel:
            self._log_frame.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 16))
            self.geometry(f"{self.winfo_width()}x{ALTURA_EXPANDIDA}")
            self.minsize(720, 560)
            self._log_visivel = True

    def _ocultar_log(self):
        if self._log_visivel:
            self._log_frame.grid_remove()
            self.geometry(f"{self.winfo_width()}x{ALTURA_COMPACTA}")
            self.minsize(720, 400)
            self._log_visivel = False

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _selecionar_caminho(self, chave, entry):
        if chave == 'config_excel':
            caminho = filedialog.askopenfilename(
                title="Selecionar mapeamento de empresas",
                filetypes=[("Excel", "*.xlsx *.xls")]
            )
        else:
            caminho = filedialog.askdirectory(title="Selecionar pasta")

        if caminho:
            entry.delete(0, "end")
            entry.insert(0, caminho)
            self._config[chave] = caminho
            self._salvar_config()

    def _log(self, msg):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _limpar_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _atualizar_progresso(self, valor):
        self._barra.set(valor / 100)
        self._label_progresso.configure(text=f"{valor}%")

    def _iniciar_processamento(self):
        if self._processando:
            return

        pasta_origem = self._entries['pasta_origem'].get().strip()
        pasta_destino = self._entries['pasta_destino'].get().strip()
        config_excel = self._entries['config_excel'].get().strip()

        if not pasta_origem or not pasta_destino:
            messagebox.showwarning(
                "Atenção",
                "Informe a pasta de origem e a pasta de destino."
            )
            return

        self._config['pasta_origem'] = pasta_origem
        self._config['pasta_destino'] = pasta_destino
        self._config['config_excel'] = config_excel
        self._salvar_config()

        self._processando = True
        self._btn_executar.configure(
            state="disabled",
            text="⏳  Processando...",
            fg_color=AZUL_MEDIO
        )
        self._barra.set(0)
        self._label_progresso.configure(text="0%")
        self._limpar_log()
        self._mostrar_log()

        thread = threading.Thread(
            target=self._processar,
            args=(pasta_origem, pasta_destino, config_excel),
            daemon=True
        )
        thread.start()

    def _processar(self, pasta_origem, pasta_destino, config_excel):
        try:
            executar_automacao(
                pasta_origem=pasta_origem,
                pasta_servidor=pasta_destino,
                caminho_config=config_excel,
                callback_log=lambda msg: self.after(0, self._log, msg),
                callback_progresso=lambda v: self.after(0, self._atualizar_progresso, v)
            )
        except Exception as e:
            self.after(0, self._log, f"❌ Erro inesperado: {str(e)}")
        finally:
            self.after(0, self._finalizar_processamento)

    def _finalizar_processamento(self):
        self._processando = False
        self._barra.set(1)
        self._label_progresso.configure(text="100%")
        self._btn_executar.configure(
            state="normal",
            text="▶   Executar Automação",
            fg_color=AZUL_ACENTO
        )


if __name__ == "__main__":
    app = NFSeApp()
    app.mainloop()
