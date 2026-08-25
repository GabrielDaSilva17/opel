import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from config_manager import carregar_config, salvar_config, testar_conexao_gemini, t


class JanelaConfiguracoes(tk.Toplevel):
    """Janela popup em Tkinter puro para configurações do Google AI Studio, Idiomas e Arquivos."""

    def __init__(self, parent, callback_atualizar=None):
        super().__init__(parent)
        self.parent = parent
        self.callback_atualizar = callback_atualizar

        self.config = carregar_config()
        self.idioma_atual = self.config.get("idioma_interface", "pt")

        self.title("⚙️ Configurações - Google AI Studio & Sistema")
        self.geometry("640x520")
        self.minsize(580, 480)
        self.resizable(False, False)

        # Modal
        self.transient(parent)
        self.grab_set()

        self._criar_layout()
        self._carregar_valores()

    def _criar_layout(self):
        # 1. Header
        f_header = ttk.Frame(self, padding=(14, 10))
        f_header.pack(fill="x")

        lbl_tit = ttk.Label(
            f_header,
            text="⚙️ Configurações & Conexão Google AI Studio",
            font=("Segoe UI", 13, "bold"),
            foreground="#0284c7",
        )
        lbl_tit.pack(anchor="w")

        # 2. Notebook de Abas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=4)

        self.tab_api = ttk.Frame(self.notebook, padding=12)
        self.tab_idioma = ttk.Frame(self.notebook, padding=12)
        self.tab_arquivos = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.tab_api, text=" 🔑 Google AI Studio & IA ")
        self.notebook.add(self.tab_idioma, text=" 🌐 Idiomas ")
        self.notebook.add(self.tab_arquivos, text=" 📁 Arquivos & Word ")

        self._montar_tab_api()
        self._montar_tab_idioma()
        self._montar_tab_arquivos()

        # 3. Rodapé com Botões de Ação
        f_footer = ttk.Frame(self, padding=(12, 10))
        f_footer.pack(fill="x")

        ttk.Button(f_footer, text="❌ Cancelar", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(f_footer, text="💾 Salvar Configurações", command=self._salvar_e_fechar).pack(side="right", padx=4)

    # -------------------------------------------------------------
    # ABA 1: GOOGLE AI STUDIO & IA
    # -------------------------------------------------------------
    def _montar_tab_api(self):
        f = self.tab_api

        # Chave API
        lbl_key = ttk.Label(f, text="Chave de API do Google AI Studio (GEMINI_API_KEY):", font=("Segoe UI", 9, "bold"))
        lbl_key.pack(anchor="w", pady=(2, 2))

        f_key = ttk.Frame(f)
        f_key.pack(fill="x", pady=(0, 8))

        self.ent_api_key = ttk.Entry(f_key, show="*")
        self.ent_api_key.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_ver_key = ttk.Button(f_key, text="👁️ Mostrar", width=10, command=self._toggle_ver_key)
        self.btn_ver_key.pack(side="right")

        # Botão de Testar Conexão
        f_teste = ttk.Frame(f)
        f_teste.pack(fill="x", pady=(0, 10))

        self.btn_testar_api = ttk.Button(
            f_teste,
            text="⚡ Testar Conexão com Google AI Studio",
            command=self._testar_api_thread,
        )
        self.btn_testar_api.pack(side="left")

        self.lbl_status_api = ttk.Label(f_teste, text="⚪ Não testado", foreground="gray")
        self.lbl_status_api.pack(side="left", padx=10)

        # Modelo
        ttk.Label(f, text="Modelo de Inteligência Artificial:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        self.var_modelo = tk.StringVar()
        self.opt_modelo = ttk.Combobox(
            f,
            textvariable=self.var_modelo,
            values=["gemini-3.5-flash", "gemini-flash-latest", "gemini-3.7-flash", "gemini-3.5-flash-lite", "gemini-flash-lite-latest", "gemini-2.5-flash"],
            state="readonly",
        )
        self.opt_modelo.pack(fill="x", pady=(0, 10))

        # Temperatura
        f_temp = ttk.Frame(f)
        f_temp.pack(fill="x", pady=(2, 2))
        ttk.Label(f_temp, text="Temperatura de Criatividade (0.0 a 1.0):", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.lbl_temp_val = ttk.Label(f_temp, text="0.70", font=("Segoe UI", 9, "bold"), foreground="#0284c7")
        self.lbl_temp_val.pack(side="right")

        self.scale_temp = ttk.Scale(f, from_=0.0, to=1.0, command=self._on_scale_temp)
        self.scale_temp.pack(fill="x", pady=(0, 6))

        # Origem / Departamento Padrão
        ttk.Label(f, text="Origem / Departamento Padrão:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        self.ent_origem = ttk.Entry(f)
        self.ent_origem.pack(fill="x", pady=(0, 6))

    def _toggle_ver_key(self):
        if self.ent_api_key.cget("show") == "*":
            self.ent_api_key.configure(show="")
            self.btn_ver_key.configure(text="🔒 Ocultar")
        else:
            self.ent_api_key.configure(show="*")
            self.btn_ver_key.configure(text="👁️ Mostrar")

    def _on_scale_temp(self, val):
        self.lbl_temp_val.configure(text=f"{float(val):.2f}")

    def _testar_api_thread(self):
        threading.Thread(target=self._testar_api, daemon=True).start()

    def _testar_api(self):
        chave = self.ent_api_key.get().strip()
        modelo = self.var_modelo.get().strip()
        self.lbl_status_api.configure(text="⏳ Conectando ao Google AI Studio...", foreground="#0284c7")
        self.btn_testar_api.configure(state="disabled")

        sucesso, msg = testar_conexao_gemini(chave, modelo)

        self.btn_testar_api.configure(state="normal")
        if sucesso:
            self.lbl_status_api.configure(text=f"✅ {msg}", foreground="#15803d")
        else:
            self.lbl_status_api.configure(text="❌ Falha na conexão", foreground="#b91c1c")
            messagebox.showerror("Erro de Conexão Google AI Studio", msg)

    # -------------------------------------------------------------
    # ABA 2: IDIOMAS
    # -------------------------------------------------------------
    def _montar_tab_idioma(self):
        f = self.tab_idioma

        ttk.Label(f, text="Idioma da Interface do Aplicativo:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        self.var_idioma_ui = tk.StringVar()
        self.opt_idioma_ui = ttk.Combobox(
            f,
            textvariable=self.var_idioma_ui,
            values=["Português (Brasil)", "Deutsch (Alemanha)", "English (Inglês)"],
            state="readonly",
        )
        self.opt_idioma_ui.pack(fill="x", pady=(0, 14))

        ttk.Label(f, text="Idioma de Redação do Relatório:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        self.var_idioma_rel = tk.StringVar()
        self.opt_idioma_rel = ttk.Combobox(
            f,
            textvariable=self.var_idioma_rel,
            values=["Deutsch (Alemão Técnico Industriemechaniker)", "Português (Técnico Industrial)", "English (Industrial Mechanics)"],
            state="readonly",
        )
        self.opt_idioma_rel.pack(fill="x", pady=(0, 10))

    # -------------------------------------------------------------
    # ABA 3: ARQUIVOS & WORD
    # -------------------------------------------------------------
    def _montar_tab_arquivos(self):
        f = self.tab_arquivos

        # Arquivo Word (.docm)
        ttk.Label(f, text="Modelo Word com Macros (.docm):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        f_word = ttk.Frame(f)
        f_word.pack(fill="x", pady=(0, 8))
        self.ent_arq_word = ttk.Entry(f_word)
        self.ent_arq_word.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(f_word, text="📂 Procurar...", command=self._procurar_arquivo_word).pack(side="right")

        # Pasta de Cópias
        ttk.Label(f, text="Pasta para salvar Cópias Semanais:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        f_copias = ttk.Frame(f)
        f_copias.pack(fill="x", pady=(0, 8))
        self.ent_pasta_copias = ttk.Entry(f_copias)
        self.ent_pasta_copias.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(f_copias, text="📂 Escolher...", command=self._procurar_pasta_copias).pack(side="right")

        # Pasta de Exemplos
        ttk.Label(f, text="Pasta de Exemplos / Aprendizado:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        f_exemplos = ttk.Frame(f)
        f_exemplos.pack(fill="x", pady=(0, 8))
        self.ent_pasta_exemplos = ttk.Entry(f_exemplos)
        self.ent_pasta_exemplos.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(f_exemplos, text="📂 Escolher...", command=self._procurar_pasta_exemplos).pack(side="right")

        # Auto-focar Word
        self.var_auto_focar = tk.BooleanVar(value=True)
        self.chk_focar = ttk.Checkbutton(f, text="Trazer Microsoft Word para o primeiro plano ao injetar dados", variable=self.var_auto_focar)
        self.chk_focar.pack(anchor="w", pady=(6, 2))

    def _procurar_arquivo_word(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo do Word",
            filetypes=[("Word Habilitado para Macro", "*.docm"), ("Documentos Word", "*.docx"), ("Todos os Arquivos", "*.*")],
        )
        if caminho:
            self.ent_arq_word.delete(0, "end")
            self.ent_arq_word.insert(0, Path(caminho).name)

    def _procurar_pasta_copias(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta de cópias")
        if caminho:
            self.ent_pasta_copias.delete(0, "end")
            self.ent_pasta_copias.insert(0, Path(caminho).name)

    def _procurar_pasta_exemplos(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta de exemplos")
        if caminho:
            self.ent_pasta_exemplos.delete(0, "end")
            self.ent_pasta_exemplos.insert(0, Path(caminho).name)

    # -------------------------------------------------------------
    # CARREGAR E SALVAR
    # -------------------------------------------------------------
    def _carregar_valores(self):
        self.ent_api_key.insert(0, self.config.get("gemini_api_key", ""))
        self.var_modelo.set(self.config.get("gemini_modelo", "gemini-3.5-flash"))

        temp = float(self.config.get("temperatura", 0.70))
        self.scale_temp.set(temp)
        self.lbl_temp_val.configure(text=f"{temp:.2f}")

        self.ent_origem.insert(0, self.config.get("origem_padrao", "Oficina Geral / Opel"))

        # Idiomas
        idioma_map_ui = {"pt": "Português (Brasil)", "de": "Deutsch (Alemanha)", "en": "English (Inglês)"}
        self.var_idioma_ui.set(idioma_map_ui.get(self.config.get("idioma_interface", "pt"), "Português (Brasil)"))

        idioma_map_rel = {
            "de": "Deutsch (Alemão Técnico Industriemechaniker)",
            "pt": "Português (Técnico Industrial)",
            "en": "English (Industrial Mechanics)",
        }
        self.var_idioma_rel.set(idioma_map_rel.get(self.config.get("idioma_relatorio", "de"), "Deutsch (Alemão Técnico Industriemechaniker)"))

        # Arquivos
        self.ent_arq_word.insert(0, self.config.get("arquivo_word", "doc ausbildund.docm"))
        self.ent_pasta_copias.insert(0, self.config.get("pasta_copias", "copias"))
        self.ent_pasta_exemplos.insert(0, self.config.get("pasta_exemplos", "exemplos_antigos"))
        self.var_auto_focar.set(self.config.get("auto_focar_word", True))

    def _salvar_e_fechar(self):
        map_ui_inv = {"Português (Brasil)": "pt", "Deutsch (Alemanha)": "de", "English (Inglês)": "en"}
        map_rel_inv = {
            "Deutsch (Alemão Técnico Industriemechaniker)": "de",
            "Português (Técnico Industrial)": "pt",
            "English (Industrial Mechanics)": "en",
        }

        novas_configs = {
            "gemini_api_key": self.ent_api_key.get().strip(),
            "gemini_modelo": self.var_modelo.get().strip(),
            "temperatura": round(float(self.scale_temp.get()), 2),
            "origem_padrao": self.ent_origem.get().strip(),
            "idioma_interface": map_ui_inv.get(self.var_idioma_ui.get(), "pt"),
            "idioma_relatorio": map_rel_inv.get(self.var_idioma_rel.get(), "de"),
            "arquivo_word": self.ent_arq_word.get().strip() or "doc ausbildund.docm",
            "pasta_copias": self.ent_pasta_copias.get().strip() or "copias",
            "pasta_exemplos": self.ent_pasta_exemplos.get().strip() or "exemplos_antigos",
            "auto_focar_word": self.var_auto_focar.get(),
            "tema": "system",
            "cor_destaque": "blue",
        }

        salvar_config(novas_configs)

        if self.callback_atualizar:
            self.callback_atualizar(novas_configs)

        messagebox.showinfo("Configurações", "Configurações salvas com sucesso!")
        self.destroy()
