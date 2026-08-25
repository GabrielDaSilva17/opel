import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path


class JanelaPopupEscrita(tk.Toplevel):
    """Janela Pop-up do Windows dedicada para escrever, editar e gerar o documento."""

    def __init__(self, parent, dados_iniciais: dict, callback_salvar=None, callback_gerar_ia=None, callback_injetar_word=None):
        super().__init__(parent)
        self.parent = parent
        self.callback_salvar = callback_salvar
        self.callback_gerar_ia = callback_gerar_ia
        self.callback_injetar_word = callback_injetar_word

        self.title("📝 Janela Pop-up de Redação do Documento (Ausbildungsnachweis)")
        self.geometry("960x780")
        self.minsize(840, 640)

        self.dias_chaves = ["montag", "dienstag", "mittwoch", "donnerstag", "freitag"]
        self.dados_iniciais = dados_iniciais

        self._criar_layout()
        self._preencher_dados_iniciais()

    def _criar_layout(self):
        # 1. Header do Popup
        f_header = ttk.Frame(self, padding=(12, 8))
        f_header.pack(fill="x")

        lbl_tit = ttk.Label(
            f_header,
            text="📝 Redação do Documento em Janela Pop-up",
            font=("Segoe UI", 14, "bold"),
            foreground="#0284c7",
        )
        lbl_tit.pack(side="left")

        # 2. Informações da Semana
        f_info = ttk.LabelFrame(self, text=" 📅 Identificação do Relatório ", padding=6)
        f_info.pack(fill="x", padx=10, pady=(0, 4))

        f_inputs = ttk.Frame(f_info)
        f_inputs.pack(fill="x")

        ttk.Label(f_inputs, text="vom:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=4, sticky="w")
        self.ent_vom = ttk.Entry(f_inputs, width=12)
        self.ent_vom.grid(row=0, column=1, padx=4, sticky="w")

        ttk.Label(f_inputs, text="bis:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=4, sticky="w")
        self.ent_bis = ttk.Entry(f_inputs, width=12)
        self.ent_bis.grid(row=0, column=3, padx=4, sticky="w")

        ttk.Label(f_inputs, text="Semana (KW):", font=("Segoe UI", 9, "bold")).grid(row=0, column=4, padx=4, sticky="w")
        self.ent_kw = ttk.Entry(f_inputs, width=12)
        self.ent_kw.grid(row=0, column=5, padx=4, sticky="w")

        ttk.Label(f_inputs, text="Origem/Depto:", font=("Segoe UI", 9, "bold")).grid(row=0, column=6, padx=4, sticky="w")
        self.ent_origem = ttk.Entry(f_inputs, width=20)
        self.ent_origem.grid(row=0, column=7, padx=4, sticky="ew")

        # 3. Prompt de Atividades e Chips Rápidos
        f_prompt_box = ttk.LabelFrame(self, text=" 🎯 Assunto / Atividades para a IA Redigir ", padding=6)
        f_prompt_box.pack(fill="x", padx=10, pady=(0, 4))

        f_chips = ttk.Frame(f_prompt_box)
        f_chips.pack(fill="x", pady=(0, 4))

        chips = [
            ("🏫 Escola LF14", "Segunda-feira: Berufsschule (LF14 Definition Projekt, LF15, W/SK). Terça a Sexta: Trabalho na oficina em usinagem no torno e fresa."),
            ("🏥 Doente (AU)", "Terça-feira: Doente / Krankheitsbedingt abwesend (AU). Segunda, Quarta a Sexta: Manutenção preventiva e hidráulica."),
            ("🏖️ Férias (Urlaub)", "Segunda a Quinta: Manutenção de bombas e troca de vedações. Sexta-feira: Urlaub (Folga)."),
            ("🔧 Hidráulica", "Troca de vedações e gaxetas em cilindros hidráulicos, alinhamento de bomba com comparador e teste de estanqueidade."),
            ("⚙️ Torno H7", "Torneamento de buchas e pinos de guia em torno convencional, respeitando tolerâncias H7 e acabamento superficial."),
        ]

        for label_text, chip_text in chips:
            ttk.Button(f_chips, text=label_text, command=lambda t=chip_text: self._aplicar_chip(t)).pack(side="left", padx=2)

        self.txt_prompt = tk.Text(f_prompt_box, height=2, font=("Segoe UI", 10), wrap="word")
        self.txt_prompt.pack(fill="x", pady=2)

        # 4. Barra de Ações Rápidas do Pop-up
        f_acoes = ttk.Frame(self, padding=(10, 2))
        f_acoes.pack(fill="x")

        ttk.Button(f_acoes, text="✨ Gerar com IA (Google AI Studio)", command=self._acionar_gerar_ia).pack(side="left", padx=2)
        ttk.Button(f_acoes, text="💾 Salvar no SQLite", command=self._acionar_salvar).pack(side="left", padx=2)
        ttk.Button(f_acoes, text="📄 Injetar no Word (.docm)", command=self._acionar_injetar_word).pack(side="left", padx=2)
        ttk.Button(f_acoes, text="📋 Copiar Tudo", command=self._copiar_tudo).pack(side="left", padx=2)
        ttk.Button(f_acoes, text="🧹 Limpar", command=self._limpar).pack(side="left", padx=2)

        # 5. Notebook de Edição no Pop-up (Caixas Diárias vs Texto Livre)
        self.notebook_popup = ttk.Notebook(self)
        self.notebook_popup.pack(fill="both", expand=True, padx=10, pady=4)

        self.tab_dias = ttk.Frame(self.notebook_popup, padding=6)
        self.tab_texto_livre = ttk.Frame(self.notebook_popup, padding=6)

        self.notebook_popup.add(self.tab_dias, text=" 📅 Edição por Dia (Segunda a Sexta) ")
        self.notebook_popup.add(self.tab_texto_livre, text=" 📄 Edição em Texto Livre Corrido ")

        self._montar_tab_dias()
        self._montar_tab_texto_livre()

        # 6. Rodapé do Pop-up
        f_footer = ttk.Frame(self, padding=(10, 8))
        f_footer.pack(fill="x")

        ttk.Button(f_footer, text="✔️ Aplicar Alterações e Fechar", command=self._aplicar_e_fechar).pack(side="right", padx=4)
        ttk.Button(f_footer, text="❌ Fechar", command=self.destroy).pack(side="right", padx=4)

    def _montar_tab_dias(self):
        f = self.tab_dias

        f_sync = ttk.Frame(f)
        f_sync.pack(fill="x", pady=(0, 4))
        ttk.Button(f_sync, text="➡️ Sincronizar para Texto Livre", command=self._sincronizar_dias_para_texto).pack(side="left")

        self.entradas_dias = {}
        dias_info = [
            ("Montag (Segunda-feira)", "montag", "#0284c7"),
            ("Dienstag (Terça-feira)", "dienstag", "#4f46e5"),
            ("Mittwoch (Quarta-feira)", "mittwoch", "#7c3aed"),
            ("Donnerstag (Quinta-feira)", "donnerstag", "#db2777"),
            ("Freitag (Sexta-feira)", "freitag", "#059669"),
        ]

        for label_text, dia_key, cor in dias_info:
            f_linha = ttk.Frame(f)
            f_linha.pack(fill="x", pady=2)

            lbl = tk.Label(f_linha, text=label_text, width=22, anchor="w", font=("Segoe UI", 9, "bold"), fg=cor)
            lbl.pack(side="left", padx=(0, 4))

            txt = tk.Text(f_linha, height=2, font=("Segoe UI", 10), wrap="word")
            txt.pack(side="left", fill="x", expand=True)
            self.entradas_dias[dia_key] = txt

    def _montar_tab_texto_livre(self):
        f = self.tab_texto_livre

        f_sync = ttk.Frame(f)
        f_sync.pack(fill="x", pady=(0, 4))
        ttk.Button(f_sync, text="⬅️ Converter deste Texto para as Caixas Diárias", command=self._sincronizar_texto_para_dias).pack(side="left")

        self.txt_livre = scrolledtext.ScrolledText(f, wrap="word", font=("Consolas", 11), height=14)
        self.txt_livre.pack(fill="both", expand=True, pady=4)

    def _preencher_dados_iniciais(self):
        self.ent_vom.insert(0, self.dados_iniciais.get("vom", ""))
        self.ent_bis.insert(0, self.dados_iniciais.get("bis", ""))
        self.ent_kw.insert(0, self.dados_iniciais.get("kw", ""))
        self.ent_origem.insert(0, self.dados_iniciais.get("origem", ""))
        self.txt_prompt.insert("1.0", self.dados_iniciais.get("prompt", ""))

        for dia in self.dias_chaves:
            self.entradas_dias[dia].insert("1.0", self.dados_iniciais.get(dia, ""))

        self._sincronizar_dias_para_texto()

    def _aplicar_chip(self, texto):
        self.txt_prompt.delete("1.0", "end")
        self.txt_prompt.insert("1.0", texto)

    def _limpar(self):
        self.txt_prompt.delete("1.0", "end")
        for dia in self.dias_chaves:
            self.entradas_dias[dia].delete("1.0", "end")
        self.txt_livre.delete("1.0", "end")

    def _sincronizar_dias_para_texto(self):
        linhas = []
        for dia in self.dias_chaves:
            conteudo = self.entradas_dias[dia].get("1.0", "end-1c").strip()
            linhas.append(f"[{dia.capitalize()}]\n{conteudo}\n")
        self.txt_livre.delete("1.0", "end")
        self.txt_livre.insert("1.0", "\n".join(linhas).strip())

    def _sincronizar_texto_para_dias(self):
        import re
        conteudo = self.txt_livre.get("1.0", "end-1c")
        padrao = r"\[(Montag|Dienstag|Mittwoch|Donnerstag|Freitag)\](.*?)(?=\[(?:Montag|Dienstag|Mittwoch|Donnerstag|Freitag)\]|\Z)"
        matches = re.findall(padrao, conteudo, re.DOTALL | re.IGNORECASE)
        if matches:
            for dia_nome, texto in matches:
                chave = dia_nome.lower()
                if chave in self.entradas_dias:
                    self.entradas_dias[chave].delete("1.0", "end")
                    self.entradas_dias[chave].insert("1.0", texto.strip())
            messagebox.showinfo("Sincronizado", "Caixas diárias atualizadas a partir do texto livre!")
        else:
            messagebox.showwarning("Aviso", "Use marcadores como [Montag], [Dienstag], etc. no texto!")

    def _copiar_tudo(self):
        self._sincronizar_dias_para_texto()
        texto = self.txt_livre.get("1.0", "end-1c")
        if texto.strip():
            self.clipboard_clear()
            self.clipboard_append(texto)
            messagebox.showinfo("Copiado", "Texto do relatório copiado para a Área de Transferência!")

    def obter_dados_popup(self) -> dict:
        return {
            "vom": self.ent_vom.get().strip(),
            "bis": self.ent_bis.get().strip(),
            "kw": self.ent_kw.get().strip(),
            "origem": self.ent_origem.get().strip(),
            "prompt": self.txt_prompt.get("1.0", "end-1c").strip(),
            "montag": self.entradas_dias["montag"].get("1.0", "end-1c").strip(),
            "dienstag": self.entradas_dias["dienstag"].get("1.0", "end-1c").strip(),
            "mittwoch": self.entradas_dias["mittwoch"].get("1.0", "end-1c").strip(),
            "donnerstag": self.entradas_dias["donnerstag"].get("1.0", "end-1c").strip(),
            "freitag": self.entradas_dias["freitag"].get("1.0", "end-1c").strip(),
        }

    def definir_dados_popup(self, dados: dict):
        if "montag" in dados:
            for dia in self.dias_chaves:
                self.entradas_dias[dia].delete("1.0", "end")
                self.entradas_dias[dia].insert("1.0", dados.get(dia, ""))
            self._sincronizar_dias_para_texto()

    def _acionar_gerar_ia(self):
        if self.callback_gerar_ia:
            self.callback_gerar_ia(self)

    def _acionar_salvar(self):
        if self.callback_salvar:
            self.callback_salvar(self.obter_dados_popup())
            messagebox.showinfo("SQLite", "Relatório salvo com sucesso no banco de dados SQLite!")

    def _acionar_injetar_word(self):
        if self.callback_injetar_word:
            self.callback_injetar_word(self.obter_dados_popup())

    def _aplicar_e_fechar(self):
        if self.callback_salvar:
            self.callback_salvar(self.obter_dados_popup())
        self.destroy()
