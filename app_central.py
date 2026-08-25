import datetime
import json
import os
import re
import sys
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from pathlib import Path

# Garante inclusão da pasta raiz no sys.path
PASTA_RAIZ = Path(__file__).parent.resolve()
if str(PASTA_RAIZ) not in sys.path:
    sys.path.insert(0, str(PASTA_RAIZ))

from pydantic import BaseModel, Field
from google import genai
from google.genai import types

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

import database as db
from config_manager import carregar_config, salvar_config, t
from config_dialog import JanelaConfiguracoes
from editor_popup import JanelaPopupEscrita
from word_manager import WordController
from extrator_exemplos import (
    extrair_texto_de_docx,
    extrair_texto_de_pdf,
    extrair_texto_de_txt,
    extrair_texto_de_imagem,
)


class AusbildungsnachweisSchema(BaseModel):
    vom: str = Field(description="DD.MM.AAAA")
    bis: str = Field(description="DD.MM.AAAA")
    nr: str = Field(description="KWXX/AAAA")
    montag: str = Field(description="Descrição técnica do dia")
    dienstag: str = Field(description="Descrição técnica do dia")
    mittwoch: str = Field(description="Descrição técnica do dia")
    donnerstag: str = Field(description="Descrição técnica do dia")
    freitag: str = Field(description="Descrição técnica do dia")


class AppCentral(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configurações do config.json
        self.config = carregar_config()
        self.idioma = self.config.get("idioma_interface", "pt")

        self.title(t("titulo_app", self.idioma))
        self.geometry("1080x890")
        self.minsize(940, 740)

        # Configura tema ttk nativo limpo
        self.style = ttk.Style(self)
        try:
            if "vista" in self.style.theme_names():
                self.style.theme_use("vista")
            elif "clam" in self.style.theme_names():
                self.style.theme_use("clam")
        except Exception:
            pass

        if getattr(sys, "frozen", False):
            self.pasta_raiz = Path(sys.executable).parent.resolve()
        else:
            self.pasta_raiz = Path(__file__).parent.resolve()
        self._atualizar_caminhos()

        self.dias_chaves = ["montag", "dienstag", "mittwoch", "donnerstag", "freitag"]
        self.data_selecionada = datetime.date.today()
        self.importando = False
        self.popup_escrita = None

        self._criar_layout_principal()
        self._definir_data(self.data_selecionada)

        # Monitoramento em segundo plano
        self.after(500, self._auto_sincronizar_inicio)
        self.after(1000, self._verificar_status_word_async)

    def _atualizar_caminhos(self):
        """Atualiza os caminhos de arquivos e pastas com base no config.json."""
        self.caminho_doc = self.pasta_raiz / self.config.get("arquivo_word", "doc ausbildund.docm")
        self.caminho_json = self.pasta_raiz / "dados_relatorio.json"
        self.caminho_vbs = self.pasta_raiz / "gerar_copia_e_zerar.vbs"
        self.pasta_copias = self.pasta_raiz / self.config.get("pasta_copias", "copias")
        self.pasta_copias.mkdir(parents=True, exist_ok=True)
        self.pasta_exemplos = self.pasta_raiz / self.config.get("pasta_exemplos", "exemplos_antigos")
        self.pasta_exemplos.mkdir(parents=True, exist_ok=True)

    def _criar_layout_principal(self):
        # 1. Header Superior Nativo Limpo
        f_header = ttk.Frame(self, padding=(12, 8))
        f_header.pack(fill="x", padx=8, pady=(6, 2))

        # Título e Subtítulo
        f_titulos = ttk.Frame(f_header)
        f_titulos.pack(side="left", fill="y")

        lbl_logo = ttk.Label(
            f_titulos,
            text="⚡ Ausbildungsnachweis AI",
            font=("Segoe UI", 16, "bold"),
            foreground="#0369a1",
        )
        lbl_logo.pack(anchor="w")

        lbl_sub = ttk.Label(
            f_titulos,
            text=t("subtitulo_app", self.idioma),
            font=("Segoe UI", 9),
            foreground="#475569",
        )
        lbl_sub.pack(anchor="w")

        # Painel da Direita: Botões e Badges
        f_topo_dir = ttk.Frame(f_header)
        f_topo_dir.pack(side="right", fill="y")

        f_topo_botoes = ttk.Frame(f_topo_dir)
        f_topo_botoes.pack(anchor="e", pady=(0, 4))

        self.btn_popup_topo = ttk.Button(
            f_topo_botoes,
            text="🪟 Escrever em Pop-up",
            command=self._abrir_popup_escrita,
        )
        self.btn_popup_topo.pack(side="left", padx=2)

        self.btn_abrir_word = ttk.Button(
            f_topo_botoes,
            text=t("btn_abrir_word", self.idioma),
            command=self._abrir_word_interativo,
        )
        self.btn_abrir_word.pack(side="left", padx=2)

        self.btn_config = ttk.Button(
            f_topo_botoes,
            text=t("btn_config", self.idioma),
            command=self._abrir_janela_configuracoes,
        )
        self.btn_config.pack(side="left", padx=2)

        f_topo_badges = ttk.Frame(f_topo_dir)
        f_topo_badges.pack(anchor="e")

        self.lbl_badge_db = ttk.Label(f_topo_badges, text="🟢 SQLite Ativo", font=("Segoe UI", 8, "bold"), foreground="#15803d")
        self.lbl_badge_db.pack(side="left", padx=4)

        self.lbl_badge_word = ttk.Label(f_topo_badges, text="⚪ Word...", font=("Segoe UI", 8, "bold"), foreground="#475569")
        self.lbl_badge_word.pack(side="left", padx=4)

        self.lbl_badge_sync = ttk.Label(f_topo_badges, text="📁 Arquivos...", font=("Segoe UI", 8, "bold"), foreground="#0369a1")
        self.lbl_badge_sync.pack(side="left", padx=4)

        # 2. Barra de Status Inferior (criada antes das abas para suportar logs de inicialização)
        f_status = ttk.Frame(self, padding=(8, 4))
        f_status.pack(side="bottom", fill="x", padx=8, pady=(2, 6))

        self.lbl_status = ttk.Label(
            f_status,
            text=t("msg_status_pronto", self.idioma),
            font=("Segoe UI", 9),
            foreground="#334155",
        )
        self.lbl_status.pack(side="left", padx=4)

        self.progress_bar = ttk.Progressbar(f_status, orient="horizontal", mode="indeterminate", length=140)

        # 3. Notebook Principal (Abas)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=4)

        self.tab_gerador = ttk.Frame(self.notebook, padding=6)
        self.tab_texto = ttk.Frame(self.notebook, padding=6)
        self.tab_extrator = ttk.Frame(self.notebook, padding=6)
        self.tab_historico = ttk.Frame(self.notebook, padding=6)

        self.notebook.add(self.tab_gerador, text=" 📝 Editor & Gerador IA ")
        self.notebook.add(self.tab_texto, text=" 📄 Visualizador de Texto Completo ")
        self.notebook.add(self.tab_extrator, text=" ⚡ Gerenciador de Arquivos & RAG ")
        self.notebook.add(self.tab_historico, text=" 🗄️ Histórico SQLite ")

        self._montar_tab_gerador()
        self._montar_tab_texto()
        self._montar_tab_extrator()
        self._montar_tab_historico()

    # -------------------------------------------------------------
    # JANELA POP-UP DEDICADA PARA ESCREVER O DOCUMENTO
    # -------------------------------------------------------------
    def _abrir_popup_escrita(self):
        """Abre uma janela pop-up do Windows dedicada exclusivamente para escrever e editar o documento."""
        dados_iniciais = {
            "vom": self.ent_vom.get().strip(),
            "bis": self.ent_bis.get().strip(),
            "kw": self.ent_kw.get().strip(),
            "origem": self.ent_app_origem.get().strip(),
            "prompt": self.txt_prompt.get("1.0", "end-1c").strip(),
            "montag": self.entradas_dias["montag"].get("1.0", "end-1c").strip(),
            "dienstag": self.entradas_dias["dienstag"].get("1.0", "end-1c").strip(),
            "mittwoch": self.entradas_dias["mittwoch"].get("1.0", "end-1c").strip(),
            "donnerstag": self.entradas_dias["donnerstag"].get("1.0", "end-1c").strip(),
            "freitag": self.entradas_dias["freitag"].get("1.0", "end-1c").strip(),
        }

        def _salvar_callback(dados_pop):
            self.ent_vom.delete(0, "end"); self.ent_vom.insert(0, dados_pop["vom"])
            self.ent_bis.delete(0, "end"); self.ent_bis.insert(0, dados_pop["bis"])
            self.ent_kw.delete(0, "end"); self.ent_kw.insert(0, dados_pop["kw"])
            self.ent_app_origem.delete(0, "end"); self.ent_app_origem.insert(0, dados_pop["origem"])
            self.txt_prompt.delete("1.0", "end"); self.txt_prompt.insert("1.0", dados_pop["prompt"])
            for dia in self.dias_chaves:
                self.entradas_dias[dia].delete("1.0", "end")
                self.entradas_dias[dia].insert("1.0", dados_pop[dia])
            self._sincronizar_dias_para_texto_unico()
            self._salvar_campos_no_banco_silencioso()

        def _gerar_ia_callback(popup_obj):
            dados_pop = popup_obj.obter_dados_popup()
            _salvar_callback(dados_pop)
            self._gerar_com_ia_thread(callback_concluido=popup_obj.definir_dados_popup)

        def _injetar_word_callback(dados_pop):
            _salvar_callback(dados_pop)
            self._injetar_word_thread()

        self.popup_escrita = JanelaPopupEscrita(
            self,
            dados_iniciais=dados_iniciais,
            callback_salvar=_salvar_callback,
            callback_gerar_ia=_gerar_ia_callback,
            callback_injetar_word=_injetar_word_callback,
        )

    # -------------------------------------------------------------
    # ABA 1: GERADOR & CAIXAS DIÁRIAS
    # -------------------------------------------------------------
    def _montar_tab_gerador(self):
        f_top = ttk.Frame(self.tab_gerador)
        f_top.pack(fill="x", pady=(0, 4))

        # 1. Período e Identificação
        f_periodo = ttk.LabelFrame(f_top, text=" 📅 Período e Identificação da Semana ", padding=6)
        f_periodo.pack(fill="x", pady=(0, 4))

        f_nav = ttk.Frame(f_periodo)
        f_nav.pack(fill="x", pady=(0, 4))

        self.btn_sem_ant = ttk.Button(f_nav, text=t("btn_sem_ant", self.idioma), command=self._semana_anterior)
        self.btn_sem_ant.pack(side="left", padx=2)

        self.btn_sem_hoje = ttk.Button(f_nav, text=t("btn_sem_hoje", self.idioma), command=self._semana_atual)
        self.btn_sem_hoje.pack(side="left", padx=2)

        self.btn_sem_prox = ttk.Button(f_nav, text=t("btn_sem_prox", self.idioma), command=self._semana_seguinte)
        self.btn_sem_prox.pack(side="left", padx=2)

        f_inputs = ttk.Frame(f_periodo)
        f_inputs.pack(fill="x")

        ttk.Label(f_inputs, text=t("lbl_vom", self.idioma), font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=4, sticky="w")
        self.ent_vom = ttk.Entry(f_inputs, width=12)
        self.ent_vom.grid(row=0, column=1, padx=4, sticky="w")
        self.ent_vom.bind("<FocusOut>", self._on_vom_digitado)
        self.ent_vom.bind("<Return>", self._on_vom_digitado)

        ttk.Label(f_inputs, text=t("lbl_bis", self.idioma), font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=4, sticky="w")
        self.ent_bis = ttk.Entry(f_inputs, width=12)
        self.ent_bis.grid(row=0, column=3, padx=4, sticky="w")

        ttk.Label(f_inputs, text=t("lbl_kw", self.idioma), font=("Segoe UI", 9, "bold")).grid(row=0, column=4, padx=4, sticky="w")
        self.ent_kw = ttk.Entry(f_inputs, width=12)
        self.ent_kw.grid(row=0, column=5, padx=4, sticky="w")

        ttk.Label(f_inputs, text=t("lbl_origem", self.idioma), font=("Segoe UI", 9, "bold")).grid(row=0, column=6, padx=4, sticky="w")
        self.ent_app_origem = ttk.Entry(f_inputs, width=24)
        self.ent_app_origem.insert(0, self.config.get("origem_padrao", "Oficina Geral / Opel"))
        self.ent_app_origem.grid(row=0, column=7, padx=4, sticky="ew")

        # 2. Atividades Realizadas & Chips Rápidos
        f_prompt = ttk.LabelFrame(f_top, text=" 🎯 Atividades Realizadas na Semana (em tópicos ou resumo) ", padding=6)
        f_prompt.pack(fill="x", pady=(0, 4))

        # Chips de Sugestão
        f_chips = ttk.Frame(f_prompt)
        f_chips.pack(fill="x", pady=(0, 4))

        chips_linha1 = [
            ("🏫 Escola / LF14", "Segunda-feira: Berufsschule (LF14 Definition Projekt, LF15, W/SK). Terça a Sexta: Trabalho na oficina em usinagem no torno e fresa."),
            ("🏥 Doente / Atestado (AU)", "Terça-feira: Doente / Krankheitsbedingt abwesend (AU). Segunda, Quarta a Sexta: Manutenção preventiva e hidráulica."),
            ("🏖️ Férias / Folga (Urlaub)", "Segunda a Quinta: Manutenção de bombas e troca de vedações. Sexta-feira: Urlaub (Folga)."),
        ]
        chips_linha2 = [
            ("🔧 Hidráulica & Bombas", "Troca de vedações e gaxetas em cilindros hidráulicos, alinhamento de bomba com comparador e teste de estanqueidade."),
            ("⚙️ Torno & Usinagem H7", "Torneamento de buchas e pinos de guia em torno convencional, respeitando tolerâncias H7 e acabamento superficial."),
            ("🛠️ Manutenção 5S", "Manutenção preventiva em prensa excêntrica, ajuste de guias lineares e etiquetagem da bancada conforme padrão 5S."),
        ]

        f_cr1 = ttk.Frame(f_chips)
        f_cr1.pack(fill="x", pady=1)
        for label_text, chip_text in chips_linha1:
            ttk.Button(f_cr1, text=label_text, command=lambda t=chip_text: self._aplicar_chip_prompt(t)).pack(side="left", padx=2, expand=True, fill="x")

        f_cr2 = ttk.Frame(f_chips)
        f_cr2.pack(fill="x", pady=1)
        for label_text, chip_text in chips_linha2:
            ttk.Button(f_cr2, text=label_text, command=lambda t=chip_text: self._aplicar_chip_prompt(t)).pack(side="left", padx=2, expand=True, fill="x")

        self.txt_prompt = tk.Text(f_prompt, height=3, font=("Segoe UI", 10), wrap="word")
        self.txt_prompt.pack(fill="x", pady=4)
        self.txt_prompt.insert("1.0", "Segunda-feira: Berufsschule (LF14 Definition Projekt, LF15, W/SK). Terça a Sexta: Trabalho na oficina em usinagem no torno e manutenção hidráulica.")

        # 3. Barra de Ações Principais
        f_actions = ttk.Frame(f_top, padding=2)
        f_actions.pack(fill="x", pady=(0, 4))

        self.btn_gerar = ttk.Button(
            f_actions,
            text="✨ 1. Gerar com IA (Google AI Studio)",
            command=self._gerar_com_ia_thread,
        )
        self.btn_gerar.pack(side="left", padx=2, expand=True, fill="x")

        self.btn_popup_acao = ttk.Button(
            f_actions,
            text="🪟 Abrir Pop-up de Redação",
            command=self._abrir_popup_escrita,
        )
        self.btn_popup_acao.pack(side="left", padx=2, expand=True, fill="x")

        self.btn_limpar = ttk.Button(
            f_actions,
            text=t("btn_limpar", self.idioma),
            command=self._limpar_formulario,
        )
        self.btn_limpar.pack(side="left", padx=2, expand=True, fill="x")

        self.btn_salvar = ttk.Button(
            f_actions,
            text=t("btn_salvar", self.idioma),
            command=self._salvar_campos_no_banco,
        )
        self.btn_salvar.pack(side="left", padx=2, expand=True, fill="x")

        self.btn_injetar = ttk.Button(
            f_actions,
            text=t("btn_injetar", self.idioma),
            command=self._injetar_word_thread,
        )
        self.btn_injetar.pack(side="left", padx=2, expand=True, fill="x")

        self.btn_vbs = ttk.Button(
            f_actions,
            text=t("btn_vbs", self.idioma),
            command=self._executar_vbs,
        )
        self.btn_vbs.pack(side="left", padx=2, expand=True, fill="x")

        # 4. Caixas Diárias (Segunda a Sexta)
        f_dias_container = ttk.LabelFrame(self.tab_gerador, text=" 📅 Preenchimento Diário (Segunda a Sexta) ", padding=6)
        f_dias_container.pack(fill="both", expand=True)

        self.entradas_dias = {}
        dias_info = [
            ("Montag (Segunda-feira)", "montag", "#0284c7"),
            ("Dienstag (Terça-feira)", "dienstag", "#4f46e5"),
            ("Mittwoch (Quarta-feira)", "mittwoch", "#7c3aed"),
            ("Donnerstag (Quinta-feira)", "donnerstag", "#db2777"),
            ("Freitag (Sexta-feira)", "freitag", "#059669"),
        ]

        for label_text, dia_key, cor in dias_info:
            f_linha = ttk.Frame(f_dias_container)
            f_linha.pack(fill="x", pady=2)

            lbl = tk.Label(
                f_linha,
                text=label_text,
                width=22,
                anchor="w",
                font=("Segoe UI", 9, "bold"),
                fg=cor,
            )
            lbl.pack(side="left", padx=(0, 4))

            txt = tk.Text(f_linha, height=2, font=("Segoe UI", 10), wrap="word")
            txt.pack(side="left", fill="x", expand=True)
            self.entradas_dias[dia_key] = txt

    # -------------------------------------------------------------
    # ABA 2: VISUALIZADOR DE TEXTO COMPLETO (ABA WINDOWS DEDICADA)
    # -------------------------------------------------------------
    def _montar_tab_texto(self):
        f = self.tab_texto

        # Barra de Ferramentas
        f_toolbar = ttk.Frame(f, padding=(0, 4))
        f_toolbar.pack(fill="x")

        ttk.Button(f_toolbar, text="🪟 Abrir em Janela Pop-up", command=self._abrir_popup_escrita).pack(side="left", padx=2)
        ttk.Button(f_toolbar, text="📋 Copiar Tudo (Clipboard)", command=self._copiar_texto_para_clipboard).pack(side="left", padx=2)
        ttk.Button(f_toolbar, text="➡️ Sincronizar das Caixas Diárias para cá", command=self._sincronizar_dias_para_texto_unico).pack(side="left", padx=2)
        ttk.Button(f_toolbar, text="⬅️ Converter deste Texto para as Caixas Diárias", command=self._sincronizar_texto_unico_para_dias).pack(side="left", padx=2)
        ttk.Button(f_toolbar, text="💾 Salvar em Arquivo .txt...", command=self._salvar_texto_em_arquivo).pack(side="left", padx=2)
        ttk.Button(f_toolbar, text="🧹 Limpar Texto", command=lambda: self.txt_completo.delete("1.0", "end")).pack(side="right", padx=2)

        ttk.Label(
            f,
            text="💡 Você pode colar ou redigir o relatório todo de uma vez usando marcadores [Montag], [Dienstag], etc.:",
            font=("Segoe UI", 9, "italic"),
            foreground="#475569",
        ).pack(anchor="w", pady=(4, 2))

        # Editor Amplo ScrolledText
        self.txt_completo = scrolledtext.ScrolledText(
            f,
            wrap="word",
            font=("Consolas", 11),
            padx=8,
            pady=8,
        )
        self.txt_completo.pack(fill="both", expand=True, pady=(2, 4))

    def _copiar_texto_para_clipboard(self):
        texto = self.txt_completo.get("1.0", "end-1c")
        if texto.strip():
            self.clipboard_clear()
            self.clipboard_append(texto)
            self._set_status("Texto completo copiado para a área de transferência!")
            messagebox.showinfo("Copiado", "O texto completo do relatório foi copiado para a Área de Transferência!")

    def _salvar_texto_em_arquivo(self):
        texto = self.txt_completo.get("1.0", "end-1c")
        if not texto.strip():
            messagebox.showwarning("Aviso", "O texto está vazio.")
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar Relatório em Texto",
            defaultextension=".txt",
            filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*")],
            initialfile=f"Relatorio_{self.ent_kw.get().strip().replace('/', '_')}.txt",
        )
        if caminho:
            Path(caminho).write_text(texto, encoding="utf-8")
            self._set_status(f"Relatório exportado para {Path(caminho).name}")
            messagebox.showinfo("Exportado", f"Arquivo salvo com sucesso em:\n{caminho}")

    # -------------------------------------------------------------
    # ABA 3: GERENCIADOR DE ARQUIVOS & APRENDIZADO RAG
    # -------------------------------------------------------------
    def _montar_tab_extrator(self):
        f = self.tab_extrator

        # 1. Barra de Ações de Arquivos
        f_top_acoes = ttk.LabelFrame(f, text=" 📁 Gerenciamento de Arquivos de Aprendizado da IA ", padding=8)
        f_top_acoes.pack(fill="x", pady=(0, 4))

        f_botoes = ttk.Frame(f_top_acoes)
        f_botoes.pack(fill="x")

        ttk.Button(f_botoes, text="➕ Adicionar Arquivos para Aprendizado...", command=self._adicionar_arquivos_dialog).pack(side="left", padx=2)
        ttk.Button(f_botoes, text="✅ Ativar / Desativar Selecionado", command=self._toggle_ativo_arquivo_selecionado).pack(side="left", padx=2)
        ttk.Button(f_botoes, text="⚡ Extrair Texto de Todos Agora", command=self._sincronizar_manual_thread).pack(side="left", padx=2)
        ttk.Button(f_botoes, text="🗑️ Remover Arquivo", command=self._remover_arquivo_selecionado).pack(side="left", padx=2)
        ttk.Button(f_botoes, text="📂 Abrir Pasta no Explorer", command=self._abrir_pasta_no_explorer).pack(side="right", padx=2)

        # 2. Tabela de Arquivos
        f_tabela = ttk.Frame(f)
        f_tabela.pack(fill="both", expand=True, pady=4)

        cols = ("Ativo", "Arquivo", "Formato", "Status no SQLite", "Tamanho")
        self.tree_arquivos = ttk.Treeview(f_tabela, columns=cols, show="headings", height=8)
        self.tree_arquivos.heading("Ativo", text="Aprendizado IA")
        self.tree_arquivos.heading("Arquivo", text="Nome do Arquivo")
        self.tree_arquivos.heading("Formato", text="Tipo")
        self.tree_arquivos.heading("Status no SQLite", text="Status")
        self.tree_arquivos.heading("Tamanho", text="Tamanho")

        self.tree_arquivos.column("Ativo", width=120, anchor="center")
        self.tree_arquivos.column("Arquivo", width=340, anchor="w")
        self.tree_arquivos.column("Formato", width=80, anchor="center")
        self.tree_arquivos.column("Status no SQLite", width=180, anchor="center")
        self.tree_arquivos.column("Tamanho", width=90, anchor="center")

        self.tree_arquivos.pack(fill="both", expand=True, side="left")
        scroll_arq = ttk.Scrollbar(f_tabela, orient="vertical", command=self.tree_arquivos.yview)
        self.tree_arquivos.configure(yscrollcommand=scroll_arq.set)
        scroll_arq.pack(side="right", fill="y")

        self.tree_arquivos.bind("<<TreeviewSelect>>", self._on_arquivo_tabela_selecionado)

        # 3. Log / Visualizador do Conteúdo Extraído
        f_log = ttk.LabelFrame(f, text=" 🧠 Conteúdo Extraído do Arquivo Selecionado (Base RAG) ", padding=6)
        f_log.pack(fill="both", expand=True, pady=(4, 0))

        self.txt_extraido = scrolledtext.ScrolledText(f_log, height=7, font=("Consolas", 10), wrap="word")
        self.txt_extraido.pack(fill="both", expand=True)

        self._atualizar_tabela_arquivos()

    def _adicionar_arquivos_dialog(self):
        """Permite ao usuário selecionar novos documentos/fotos para importar para a base."""
        arquivos = filedialog.askopenfilenames(
            title="Selecione os arquivos para aprendizado da IA",
            filetypes=[
                ("Documentos e Imagens", "*.docx *.docm *.pdf *.txt *.png *.jpg *.jpeg *.webp *.bmp"),
                ("Documentos Word", "*.docx *.docm"),
                ("PDFs", "*.pdf"),
                ("Imagens / Prints", "*.png *.jpg *.jpeg *.webp *.bmp"),
                ("Texto", "*.txt"),
                ("Todos os Arquivos", "*.*"),
            ],
        )
        if arquivos:
            copiados = 0
            for arq_orig in arquivos:
                p_orig = Path(arq_orig)
                p_dest = self.pasta_exemplos / p_orig.name
                if not p_dest.exists():
                    shutil.copy2(p_orig, p_dest)
                    copiados += 1

            self._atualizar_tabela_arquivos()
            self._set_status(f"{copiados} arquivo(s) adicionado(s) à pasta de aprendizado!")
            self._sincronizar_manual_thread()

    def _toggle_ativo_arquivo_selecionado(self):
        """Ativa ou desativa um arquivo selecionado na base de aprendizado da IA."""
        item = self.tree_arquivos.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um arquivo na tabela para ativar/desativar.")
            return

        nome_arq = self.tree_arquivos.item(item[0], "values")[1]
        novo_status = db.alternar_status_conhecimento(nome_arq)
        self._atualizar_tabela_arquivos()
        status_txt = "ATIVADO (incluído no aprendizado da IA)" if novo_status else "DESATIVADO (ignorado pela IA)"
        self._set_status(f"Arquivo '{nome_arq}' agora está {status_txt}.")

    def _remover_arquivo_selecionado(self):
        """Remove o arquivo selecionado da pasta de exemplos e do banco de conhecimento."""
        item = self.tree_arquivos.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um arquivo na tabela para remover.")
            return

        nome_arq = self.tree_arquivos.item(item[0], "values")[1]
        confirmar = messagebox.askyesno("Remover Arquivo", f"Deseja realmente remover o arquivo '{nome_arq}' do aprendizado e do disco?")
        if confirmar:
            p_arq = self.pasta_exemplos / nome_arq
            if p_arq.exists():
                try:
                    p_arq.unlink()
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível excluir o arquivo: {e}")
                    return

            db.excluir_conhecimento_arquivo(nome_arq)
            self._atualizar_tabela_arquivos()
            self.txt_extraido.delete("1.0", "end")
            self._set_status(f"Arquivo '{nome_arq}' removido com sucesso.")

    def _on_arquivo_tabela_selecionado(self, event=None):
        item = self.tree_arquivos.selection()
        if not item:
            return
        nome_arq = self.tree_arquivos.item(item[0], "values")[1]
        conteudo = db.obter_conteudo_por_arquivo(nome_arq)
        self.txt_extraido.delete("1.0", "end")
        if conteudo:
            self.txt_extraido.insert("1.0", f"=== CONTEÚDO EXTRAÍDO DE '{nome_arq}' ===\n\n{conteudo}")
        else:
            self.txt_extraido.insert("1.0", f"Arquivo '{nome_arq}' ainda não foi processado ou está pendente de extração.")

    def _abrir_pasta_no_explorer(self):
        os.startfile(str(self.pasta_exemplos))

    def _atualizar_tabela_arquivos(self):
        for item in self.tree_arquivos.get_children():
            self.tree_arquivos.delete(item)

        status_ativos = db.obter_status_ativo_arquivos()
        importados = set(status_ativos.keys())

        extensoes = ("*.docx", "*.docm", "*.pdf", "*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp", "*.txt")
        arquivos = []
        for ext in extensoes:
            arquivos.extend(
                [
                    a for a in self.pasta_exemplos.glob(ext)
                    if not a.name.startswith("~$")
                    and not a.name.lower().startswith("readme")
                    and not a.name.endswith(".png.txt")
                    and not a.name.endswith(".jpg.txt")
                    and not a.name.endswith(".jpeg.txt")
                ]
            )

        ativos_count = 0
        for arq in sorted(arquivos, key=lambda x: x.name.lower()):
            esta_importado = arq.name in importados
            esta_ativo = status_ativos.get(arq.name, True)
            if esta_ativo:
                ativos_count += 1

            ativo_txt = "✅ ATIVO" if esta_ativo else "❌ Inativo"
            status_txt = "✅ Salvo no SQLite" if esta_importado else "🆕 Novo (Pendente)"
            tam_kb = f"{max(1, arq.stat().st_size // 1024)} KB"

            self.tree_arquivos.insert("", "end", values=(ativo_txt, arq.name, arq.suffix.upper(), status_txt, tam_kb))

        self.lbl_badge_sync.configure(text=f"📁 {ativos_count}/{len(arquivos)} no Aprendizado")

    def _auto_sincronizar_inicio(self):
        threading.Thread(target=self._sincronizar_arquivos_pendentes, daemon=True).start()

    def _sincronizar_manual_thread(self):
        if self.importando:
            messagebox.showinfo("Aviso", "A sincronização já está em andamento!")
            return
        threading.Thread(target=self._sincronizar_arquivos_pendentes, daemon=True).start()

    def _sincronizar_arquivos_pendentes_sync(self, api_key: str):
        client = None
        if api_key:
            try:
                client = genai.Client(api_key=api_key)
            except Exception:
                pass

        importados = db.listar_nomes_arquivos_importados()
        extensoes = ("*.docx", "*.docm", "*.pdf", "*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp", "*.txt")
        arquivos = []
        for ext in extensoes:
            arquivos.extend(
                [
                    a for a in self.pasta_exemplos.glob(ext)
                    if not a.name.startswith("~$")
                    and not a.name.lower().startswith("readme")
                    and not a.name.endswith(".png.txt")
                    and not a.name.endswith(".jpg.txt")
                    and not a.name.endswith(".jpeg.txt")
                ]
            )

        novos = [a for a in arquivos if a.name not in importados]
        for arq in novos:
            self._processar_arquivo_individual(arq, client)

    def _sincronizar_arquivos_pendentes(self):
        if self.importando:
            return
        self.importando = True
        self._set_status("Varrendo pasta e extraindo texto de novos arquivos...")

        api_key = self.config.get("gemini_api_key") or os.getenv("GEMINI_API_KEY", "")
        client = None
        if api_key:
            try:
                client = genai.Client(api_key=api_key)
            except Exception:
                pass

        importados = db.listar_nomes_arquivos_importados()
        extensoes = ("*.docx", "*.docm", "*.pdf", "*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp", "*.txt")
        arquivos = []
        for ext in extensoes:
            arquivos.extend(
                [a for a in self.pasta_exemplos.glob(ext) if not a.name.startswith("~$") and not a.name.lower().startswith("readme")]
            )

        novos = [a for a in arquivos if a.name not in importados]
        if not novos:
            self._set_status(f"Todos os {len(arquivos)} arquivos estão sincronizados.")
            self._atualizar_tabela_arquivos()
            self.importando = False
            return

        processados = 0
        for arq in novos:
            self._set_status(f"Extraindo texto: {arq.name}...")
            sucesso = self._processar_arquivo_individual(arq, client)
            if sucesso:
                processados += 1

        self._atualizar_tabela_arquivos()
        self._set_status(f"Sincronização concluída! {processados} novo(s) arquivo(s) salvos no SQLite.")
        self.importando = False

    def _processar_arquivo_individual(self, arq: Path, client: genai.Client) -> bool:
        ext = arq.suffix.lower()
        texto_extraido = ""

        try:
            if ext in (".docx", ".docm"):
                frases = extrair_texto_de_docx(arq)
                texto_extraido = "\n".join(frases)
            elif ext == ".pdf":
                frases = extrair_texto_de_pdf(arq)
                texto_extraido = "\n".join(frases)
            elif ext == ".txt":
                frases = extrair_texto_de_txt(arq)
                texto_extraido = "\n".join(frases)
            elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
                if client:
                    frases = extrair_texto_de_imagem(arq, client)
                    texto_extraido = "\n".join(frases)
                else:
                    return False

            if texto_extraido.strip():
                db.salvar_conhecimento("Auto-Import / pasta", arq.name, texto_extraido.strip(), ativo=1)
                return True
        except Exception as e:
            print(f"Erro ao processar {arq.name}: {e}")

        return False

    # -------------------------------------------------------------
    # ABA 4: HISTÓRICO SQLITE
    # -------------------------------------------------------------
    # -------------------------------------------------------------
    # ABA 4: HISTÓRICO SQLITE
    # -------------------------------------------------------------
    def _montar_tab_historico(self):
        f = self.tab_historico

        # 1. Barra de Busca e Filtro
        f_busca = ttk.LabelFrame(f, text=" 🔍 Pesquisar no Histórico SQLite (por ID, KW, Período, Origem ou Atividades) ", padding=6)
        f_busca.pack(fill="x", pady=(0, 4))

        f_linha_busca = ttk.Frame(f_busca)
        f_linha_busca.pack(fill="x")

        ttk.Label(f_linha_busca, text="Buscar (ID / KW / Texto):", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))

        self.ent_busca_hist = ttk.Entry(f_linha_busca, font=("Segoe UI", 10), width=28)
        self.ent_busca_hist.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.ent_busca_hist.bind("<KeyRelease>", lambda e: self._pesquisar_historico())
        self.ent_busca_hist.bind("<Return>", lambda e: self._pesquisar_historico())

        ttk.Button(f_linha_busca, text="🔍 Buscar", command=self._pesquisar_historico).pack(side="left", padx=2)
        ttk.Button(f_linha_busca, text="🧹 Limpar Filtro", command=self._limpar_busca_historico).pack(side="left", padx=2)

        self.lbl_busca_count = ttk.Label(f_linha_busca, text="", font=("Segoe UI", 9, "italic"), foreground="#0369a1")
        self.lbl_busca_count.pack(side="left", padx=6)

        # 2. Barra de Ferramentas / Ações da Tabela
        f_topo_hist = ttk.Frame(f, padding=(0, 4))
        f_topo_hist.pack(fill="x")

        ttk.Label(
            f_topo_hist,
            text="💡 Selecione um relatório para carregar no Editor ou clique 2x na linha:",
            font=("Segoe UI", 9, "bold"),
            foreground="#0284c7",
        ).pack(side="left")

        ttk.Button(f_topo_hist, text="📂 Carregar no Editor", command=self._carregar_relatorio_selecionado).pack(side="right", padx=2)
        ttk.Button(f_topo_hist, text="🗑️ Excluir Selecionado", command=self._excluir_relatorio_selecionado).pack(side="right", padx=2)
        ttk.Button(f_topo_hist, text="🔄 Atualizar Lista", command=self._carregar_lista_historico).pack(side="right", padx=2)

        f_tabela_hist = ttk.Frame(f)
        f_tabela_hist.pack(fill="both", expand=True, pady=4)

        cols = ("ID", "Data Registro", "KW", "vom (Início)", "bis (Fim)", "Origem/Depto")
        self.tree_hist = ttk.Treeview(f_tabela_hist, columns=cols, show="headings", height=15)
        for col in cols:
            self.tree_hist.heading(col, text=col)
            self.tree_hist.column(col, width=130, anchor="center")

        self.tree_hist.pack(fill="both", expand=True, side="left")
        scroll = ttk.Scrollbar(f_tabela_hist, orient="vertical", command=self.tree_hist.yview)
        self.tree_hist.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")

        self.tree_hist.bind("<Double-1>", self._carregar_relatorio_selecionado)
        self._carregar_lista_historico()

    def _pesquisar_historico(self):
        termo = self.ent_busca_hist.get().strip()
        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)
        registros = db.pesquisar_relatorios(termo)
        for row in registros:
            self.tree_hist.insert("", "end", values=row)

        total_relatorios = len(db.listar_relatorios())
        if termo:
            self.lbl_busca_count.configure(text=f"Exibindo: {len(registros)} de {total_relatorios}")
            self._set_status(f"Filtro '{termo}': {len(registros)} relatório(s) encontrado(s).")
        else:
            self.lbl_busca_count.configure(text=f"Total: {total_relatorios} relatórios")
            self._set_status(f"Exibindo todos os {total_relatorios} relatórios salvos.")

    def _limpar_busca_historico(self):
        self.ent_busca_hist.delete(0, "end")
        self._pesquisar_historico()

    def _carregar_lista_historico(self):
        self._pesquisar_historico()
        total_relatorios = len(db.listar_relatorios())
        if total_relatorios > 0:
            self.lbl_badge_db.configure(text=f"🟢 SQLite ({total_relatorios} Registros)")
        else:
            self.lbl_badge_db.configure(text="🟢 SQLite Ativo")

    def _carregar_relatorio_selecionado(self, event=None):
        item = self.tree_hist.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um relatório na tabela para carregar.")
            return
        relatorio_id = self.tree_hist.item(item[0], "values")[0]
        dados = db.obter_relatorio_por_id(relatorio_id)
        if dados:
            self.ent_vom.delete(0, "end"); self.ent_vom.insert(0, dados[2] or "")
            self.ent_bis.delete(0, "end"); self.ent_bis.insert(0, dados[3] or "")
            self.ent_kw.delete(0, "end"); self.ent_kw.insert(0, dados[4] or "")
            self.ent_app_origem.delete(0, "end"); self.ent_app_origem.insert(0, dados[5] or "")
            self.txt_prompt.delete("1.0", "end"); self.txt_prompt.insert("1.0", dados[6] or "")

            for idx, dia in enumerate(self.dias_chaves):
                self.entradas_dias[dia].delete("1.0", "end")
                self.entradas_dias[dia].insert("1.0", dados[7 + idx] or "")

            self._sincronizar_dias_para_texto_unico()
            self.notebook.select(0)
            self._set_status(f"Relatório #{relatorio_id} ({dados[4]}) carregado no editor!")
            messagebox.showinfo("Carregado", f"Relatório #{relatorio_id} ({dados[4]}) carregado com sucesso na aba do Editor!")

    def _excluir_relatorio_selecionado(self):
        item = self.tree_hist.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um relatório na tabela para excluir.")
            return
        relatorio_id = self.tree_hist.item(item[0], "values")[0]
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir o relatório #{relatorio_id} do banco de dados?"):
            db.excluir_relatorio(relatorio_id)
            self._carregar_lista_historico()
            self._set_status(f"Relatório #{relatorio_id} excluído do SQLite.")

    # -------------------------------------------------------------
    # LÓGICA DE SINCRONIZAÇÃO ENTRE CAIXAS E TEXTO ÚNICO
    # -------------------------------------------------------------
    def _sincronizar_dias_para_texto_unico(self):
        linhas = []
        for dia in self.dias_chaves:
            conteudo = self.entradas_dias[dia].get("1.0", "end-1c").strip()
            linhas.append(f"[{dia.capitalize()}]\n{conteudo}\n")

        texto_completo = "\n".join(linhas).strip()
        self.txt_completo.delete("1.0", "end")
        self.txt_completo.insert("1.0", texto_completo)
        self._set_status("Sincronizado: Caixas diárias -> Visualizador de Texto Completo.")

    def _sincronizar_texto_unico_para_dias(self):
        conteudo_total = self.txt_completo.get("1.0", "end-1c")

        padrao = r"\[(Montag|Dienstag|Mittwoch|Donnerstag|Freitag)\](.*?)(?=\[(?:Montag|Dienstag|Mittwoch|Donnerstag|Freitag)\]|\Z)"
        matches = re.findall(padrao, conteudo_total, re.DOTALL | re.IGNORECASE)

        if matches:
            for dia_nome, texto in matches:
                chave = dia_nome.lower()
                if chave in self.entradas_dias:
                    self.entradas_dias[chave].delete("1.0", "end")
                    self.entradas_dias[chave].insert("1.0", texto.strip())
            self._set_status("Sincronizado: Texto Completo -> Caixas diárias.")
            messagebox.showinfo("Sincronizado", "O texto foi convertido e preenchido nas 5 caixas diárias com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Use marcadores como [Montag], [Dienstag], [Mittwoch], [Donnerstag], [Freitag] no texto!")

    def _obter_dados_dias_atuais(self) -> dict:
        return {dia: self.entradas_dias[dia].get("1.0", "end-1c").strip() for dia in self.dias_chaves}

    def _aplicar_chip_prompt(self, texto: str):
        self.txt_prompt.delete("1.0", "end")
        self.txt_prompt.insert("1.0", texto)
        self._set_status("Sugestão técnica aplicada ao campo de atividades.")

    def _limpar_formulario(self):
        confirmar = messagebox.askyesno("Limpar Formulário", t("msg_limpar_confirma", self.idioma))
        if not confirmar:
            return

        self.txt_prompt.delete("1.0", "end")
        for dia in self.dias_chaves:
            self.entradas_dias[dia].delete("1.0", "end")

        self.txt_completo.delete("1.0", "end")
        self._definir_data(datetime.date.today())
        self._set_status(t("msg_limpo_sucesso", self.idioma))

    # -------------------------------------------------------------
    # GERENCIAMENTO DE DATAS & AUTO-KW
    # -------------------------------------------------------------
    def _definir_data(self, data_ref: datetime.date):
        self.data_selecionada = data_ref
        segunda = data_ref - datetime.timedelta(days=data_ref.weekday())
        sexta = segunda + datetime.timedelta(days=4)
        ano, semana_iso, _ = segunda.isocalendar()
        kw_str = f"KW{semana_iso:02d}/{ano}"

        self.ent_vom.delete(0, "end"); self.ent_vom.insert(0, segunda.strftime("%d.%m.%Y"))
        self.ent_bis.delete(0, "end"); self.ent_bis.insert(0, sexta.strftime("%d.%m.%Y"))
        self.ent_kw.delete(0, "end"); self.ent_kw.insert(0, kw_str)

        self._set_status(f"Período ativo: {kw_str} ({segunda.strftime('%d.%m.%Y')} até {sexta.strftime('%d.%m.%Y')})")

    def _semana_anterior(self):
        self._definir_data(self.data_selecionada - datetime.timedelta(days=7))

    def _semana_atual(self):
        self._definir_data(datetime.date.today())

    def _semana_seguinte(self):
        self._definir_data(self.data_selecionada + datetime.timedelta(days=7))

    def _on_vom_digitado(self, event=None):
        texto = self.ent_vom.get().strip()
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt = datetime.datetime.strptime(texto, fmt).date()
                self._definir_data(dt)
                return
            except ValueError:
                pass

    def _set_status(self, msg: str):
        if hasattr(self, "lbl_status") and self.lbl_status is not None:
            try:
                self.lbl_status.configure(text=msg)
                self.update_idletasks()
            except Exception:
                pass

    # -------------------------------------------------------------
    # CONTROLE E VERIFICAÇÃO DO WORD
    # -------------------------------------------------------------
    def _verificar_status_word_async(self):
        def _check():
            rodando = WordController.verificar_word_em_execucao()
            if rodando:
                self.lbl_badge_word.configure(text="🔵 Word Aberto", foreground="#0284c7")
            else:
                self.lbl_badge_word.configure(text="⚪ Word Fechado", foreground="#475569")
        threading.Thread(target=_check, daemon=True).start()

    def _abrir_word_interativo(self):
        self._set_status("Abrindo Microsoft Word em primeiro plano...")
        sucesso, msg = WordController.abrir_documento_para_edicao(self.caminho_doc)
        if sucesso:
            self._set_status("Word aberto com sucesso em primeiro plano.")
            self._verificar_status_word_async()
            messagebox.showinfo("Microsoft Word", "O documento do Word foi aberto em primeiro plano!\nVocê pode verificar macros e conceder permissões se necessário.")
        else:
            self._set_status("Falha ao abrir Word.")
            messagebox.showerror("Erro ao Abrir Word", msg)

    # -------------------------------------------------------------
    # POPUP DE CONFIGURAÇÕES
    # -------------------------------------------------------------
    def _abrir_janela_configuracoes(self):
        JanelaConfiguracoes(self, callback_atualizar=self._aplicar_novas_configuracoes)

    def _aplicar_novas_configuracoes(self, novas_configs: dict):
        self.config = novas_configs
        self.idioma = self.config.get("idioma_interface", "pt")
        self._atualizar_caminhos()

        self.title(t("titulo_app", self.idioma))
        self.btn_abrir_word.configure(text=t("btn_abrir_word", self.idioma))
        self.btn_config.configure(text=t("btn_config", self.idioma))

        self.btn_sem_ant.configure(text=t("btn_sem_ant", self.idioma))
        self.btn_sem_hoje.configure(text=t("btn_sem_hoje", self.idioma))
        self.btn_sem_prox.configure(text=t("btn_sem_prox", self.idioma))

        self.btn_limpar.configure(text=t("btn_limpar", self.idioma))
        self.btn_salvar.configure(text=t("btn_salvar", self.idioma))
        self.btn_injetar.configure(text=t("btn_injetar", self.idioma))
        self.btn_vbs.configure(text=t("btn_vbs", self.idioma))

        self._verificar_status_word_async()
        self._atualizar_tabela_arquivos()
        self._carregar_lista_historico()
        self._set_status("Configurações atualizadas com sucesso!")

    # -------------------------------------------------------------
    # GERAÇÃO COM GOOGLE AI STUDIO (100% DINÂMICA COM FALLBACK)
    # -------------------------------------------------------------
    def _gerar_com_ia_thread(self, callback_concluido=None):
        threading.Thread(target=self._gerar_com_ia, args=(callback_concluido,), daemon=True).start()

    def _gerar_com_ia(self, callback_concluido=None):
        api_key = self.config.get("gemini_api_key") or os.getenv("GEMINI_API_KEY", "")
        if not api_key:
            messagebox.showerror("Erro", "Chave do Google AI Studio não configurada. Abra as Configurações (⚙️) para definir sua chave.")
            return

        modelo = self.config.get("gemini_modelo", "gemini-3.5-flash")
        temp = float(self.config.get("temperatura", 0.70))
        idioma_rel = self.config.get("idioma_relatorio", "de")

        instrucao_idioma = "Redija as frases exclusivamente em alemão técnico profissional para Industriemechaniker (Mecânica Industrial).\n"
        if idioma_rel == "pt":
            instrucao_idioma = "Redija as frases em português técnico industrial claro e objetivo.\n"
        elif idioma_rel == "en":
            instrucao_idioma = "Redija as frases em inglês técnico industrial (Industrial Mechanics).\n"

        self.btn_gerar.configure(state="disabled")
        self.progress_bar.pack(side="right", padx=8, pady=2)
        self.progress_bar.start()
        self._set_status(f"Consultando Google AI Studio ({modelo})...")

        try:
            try:
                self._sincronizar_arquivos_pendentes_sync(api_key)
            except Exception:
                pass

            client = genai.Client(api_key=api_key)

            # Vocabulário dos arquivos ATIVOS
            termos_apoio = []
            try:
                conhecimentos = db.obter_textos_conhecimento()
                for c in conhecimentos[:10]:
                    for l in c.splitlines():
                        l_clean = l.strip()
                        if len(l_clean) > 25 and not any(ig in l_clean.lower() for ig in ["abteilung", "name:", "unterschrift", "datum", "vom:"]):
                            termos_apoio.append(l_clean)
            except Exception:
                pass

            bloco_vocabulario = ""
            if termos_apoio:
                termos_unicos = list(dict.fromkeys(termos_apoio))[:8]
                bloco_vocabulario = "\nVocabulário de apoio (use apenas como referência terminológica, NÃO copie textualmente):\n- " + "\n- ".join(termos_unicos) + "\n"

            sys_inst = (
                "Você é uma Inteligência Artificial de alta capacidade atuando como redator técnico "
                "de relatórios semanais de formação profissional (Ausbildungsnachweis) para Industriemechaniker (Mecânica Industrial / Opel).\n\n"
                "SUA MISSÃO PRINCIPAL:\n"
                "Gerar descrições diárias 100% DINÂMICAS, AUTÊNTICAS e PERSONALIZADAS para cada dia (Montag a Freitag), "
                "seguindo FIELMENTE o que o usuário solicitou nas atividades da semana.\n\n"
                "DIRETRIZES OBRIGATÓRIAS POR SITUAÇÃO:\n"
                "1. ESCOLA TÉCNICA / AULAS (Berufsschule / LF14 / LF15 / W/SK / Projeto / Foto):\n"
                "   - Quando o usuário mencionar escola, aulas, matérias ou foto/LF14:\n"
                "   - Redija a descrição teórica correspondente, ex: 'Berufsschule: Vertiefung in LF14 (Definition und Phasen von Projekten, Arbeitsblatt Projekt), LF15 sowie Wirtschafts- und Sozialkunde (W/SK).'\n"
                "2. DOENÇA / ATESTADO MÉDICO (Krank / AU / Atestado / Falta médica):\n"
                "   - Quando o usuário mencionar doença ou atestado em algum dia:\n"
                "   - Redija exatamente: 'Krankheitsbedingt abwesend (Arbeitsunfähigkeitsbescheinigung liegt vor).'\n"
                "3. FÉRIAS / FOLGA (Urlaub / Freizeitausgleich / Feiertag):\n"
                "   - Redija 'Urlaub.' ou 'Gesetzlicher Feiertag.'\n"
                "4. TRABALHO / OFICINA / FÁBRICA (Betrieb / Werkstatt / Usinagem / Manutenção / Testes / Torno / Fresa / Hidráulica / Elétrica / Montagem):\n"
                "   - Crie tarefas técnicas reais, específicas e detalhadas para Industriemechaniker com base no que o usuário pediu.\n"
                "   - Use verbos substantivados no início das frases (Fertigung, Demontage, Montage, Ausrichtung, Überprüfung, Wartung, Kalibrierung, Fehlerbehebung).\n"
                "5. PROIBIÇÃO:\n"
                "   - NUNCA retorne textos fixos repetidos do banco. Crie SEMPRE respostas novas e adaptadas ao que o usuário digitou!\n\n"
                f"{instrucao_idioma}"
                f"{bloco_vocabulario}"
            )

            prompt = f"""
            Gere o preenchimento diário do relatório semanal (Montag até Freitag):
            - Período: vom {self.ent_vom.get().strip()} bis {self.ent_bis.get().strip()} ({self.ent_kw.get().strip()})
            - Origem/Departamento: {self.ent_app_origem.get().strip()}
            - ATIVIDADES SOLICITADAS PELO APRENDIZ (Interprete com máxima fidelidade):
            \"\"\"
            {self.txt_prompt.get('1.0', 'end-1c').strip()}
            \"\"\"

            INSTRUÇÕES:
            - Distribua as situações nos dias apropriados (Montag a Freitag) conforme informado pelo usuário.
            - Se o usuário pediu escola na segunda, doente na terça e oficina no restante, preencha exatamente essa sequência!
            """

            temp_ajustada = max(0.65, temp)
            modelos_fallback = [modelo, "gemini-3.5-flash", "gemini-flash-latest", "gemini-3.7-flash", "gemini-3.5-flash-lite", "gemini-flash-lite-latest", "gemini-3.1-flash-lite"]
            modelos_unicos = list(dict.fromkeys(modelos_fallback))

            dados = None
            ultimo_erro = None
            modelo_usado = modelo

            for mod in modelos_unicos:
                try:
                    self._set_status(f"Consultando Google AI Studio ({mod})...")
                    res = client.models.generate_content(
                        model=mod,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            system_instruction=sys_inst,
                            response_mime_type="application/json",
                            response_schema=AusbildungsnachweisSchema,
                            temperature=temp_ajustada,
                            automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
                        ),
                    )
                    if res and res.text:
                        dados = json.loads(res.text)
                        modelo_usado = mod
                        break
                except Exception as err:
                    ultimo_erro = err

            if not dados:
                raise ultimo_erro or Exception("Não foi possível gerar com os modelos disponíveis.")

            for dia in self.dias_chaves:
                self.entradas_dias[dia].delete("1.0", "end")
                self.entradas_dias[dia].insert("1.0", dados.get(dia, ""))

            self._sincronizar_dias_para_texto_unico()

            # Salva automaticamente no banco de dados SQLite
            try:
                db.salvar_ou_atualizar_relatorio(
                    vom=self.ent_vom.get().strip(),
                    bis=self.ent_bis.get().strip(),
                    kw=self.ent_kw.get().strip(),
                    origem_app=self.ent_app_origem.get().strip(),
                    prompt=self.txt_prompt.get("1.0", "end-1c").strip(),
                    m=dados.get("montag", ""),
                    d=dados.get("dienstag", ""),
                    mi=dados.get("mittwoch", ""),
                    don=dados.get("donnerstag", ""),
                    f=dados.get("freitag", ""),
                )
                self._carregar_lista_historico()
            except Exception:
                pass

            if callback_concluido:
                callback_concluido(dados)

            self._set_status(f"Relatório gerado e salvo no SQLite via Google AI Studio ({modelo_usado})!")
        except Exception as e:
            messagebox.showerror("Erro Google AI Studio", f"Falha na geração: {e}")
            self._set_status("Erro ao gerar com IA.")
        finally:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.btn_gerar.configure(state="normal")

    # -------------------------------------------------------------
    # SALVAR NO SQLITE
    # -------------------------------------------------------------
    def _salvar_campos_no_banco_silencioso(self):
        dados_dias = self._obter_dados_dias_atuais()
        db.salvar_ou_atualizar_relatorio(
            self.ent_vom.get().strip(),
            self.ent_bis.get().strip(),
            self.ent_kw.get().strip(),
            self.ent_app_origem.get().strip(),
            self.txt_prompt.get("1.0", "end-1c").strip(),
            dados_dias["montag"],
            dados_dias["dienstag"],
            dados_dias["mittwoch"],
            dados_dias["donnerstag"],
            dados_dias["freitag"],
        )
        self._carregar_lista_historico()

    def _salvar_campos_no_banco(self):
        dados_dias = self._obter_dados_dias_atuais()
        novo_id = db.salvar_ou_atualizar_relatorio(
            self.ent_vom.get().strip(),
            self.ent_bis.get().strip(),
            self.ent_kw.get().strip(),
            self.ent_app_origem.get().strip(),
            self.txt_prompt.get("1.0", "end-1c").strip(),
            dados_dias["montag"],
            dados_dias["dienstag"],
            dados_dias["mittwoch"],
            dados_dias["donnerstag"],
            dados_dias["freitag"],
        )
        self._carregar_lista_historico()
        self._set_status(f"Relatório salvo no SQLite com ID #{novo_id}!")
        messagebox.showinfo("SQLite", f"Relatório da {self.ent_kw.get().strip()} salvo com sucesso no banco de dados (ID #{novo_id})!")

    # -------------------------------------------------------------
    # INJEÇÃO NO WORD COM WORDCONTROLLER
    # -------------------------------------------------------------
    def _injetar_word_thread(self):
        threading.Thread(target=self._injetar_word, daemon=True).start()

    def _injetar_word(self):
        if not self.caminho_doc.exists():
            messagebox.showerror("Erro", f"Arquivo Word não encontrado:\n{self.caminho_doc}")
            return

        dados_dias = self._obter_dados_dias_atuais()
        dados = {
            "vom": self.ent_vom.get().strip(),
            "bis": self.ent_bis.get().strip(),
            "nr": self.ent_kw.get().strip(),
            "montag": dados_dias["montag"],
            "dienstag": dados_dias["dienstag"],
            "mittwoch": dados_dias["mittwoch"],
            "donnerstag": dados_dias["donnerstag"],
            "freitag": dados_dias["freitag"],
        }

        self._set_status("Verificando Word e executando macro VBA...")
        focar = self.config.get("auto_focar_word", True)

        sucesso, msg = WordController.injetar_dados_e_executar_macro(
            caminho_doc=self.caminho_doc,
            caminho_json=self.caminho_json,
            dados=dados,
            nome_macro="CarregarDadosDoJSON",
            trazer_para_frente=focar,
        )

        self._verificar_status_word_async()

        if sucesso:
            self._set_status("Documento Word atualizado e salvo com sucesso!")
            messagebox.showinfo("Sucesso Word", "Todas as caixas do Word foram preenchidas e salvas com sucesso!")
        else:
            self._set_status("Erro ao injetar no Word.")
            messagebox.showerror("Erro Word", msg)

    # -------------------------------------------------------------
    # BACKUP E LIMPEZA VBS
    # -------------------------------------------------------------
    def _executar_vbs(self):
        if not self.caminho_vbs.exists():
            messagebox.showerror("Erro", f"Arquivo VBScript não encontrado:\n{self.caminho_vbs}")
            return
        confirmar = messagebox.askyesno("Confirmar", "Deseja salvar uma cópia em /copias e zerar o documento principal?")
        if confirmar:
            subprocess.run(["wscript.exe", str(self.caminho_vbs)])
            self._set_status("Cópia criada e documento original zerado!")


if __name__ == "__main__":
    app = AppCentral()
    app.mainloop()
