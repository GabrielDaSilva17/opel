import datetime
import json
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import pythoncom
import win32com.client as win32

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

# Importa o extrator multimodal
from extrator_exemplos import carregar_base_multimodal

NOME_ARQUIVO_WORD = "doc ausbildund.docm"
CHAVE_API_PADRAO = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6IS3ZmzQm8J5zjABuinQ-75pFu0XPYye5CbRPKdrTQg7Q")


class AusbildungsnachweisSchema(BaseModel):
    vom: str = Field(description="Data inicial DD.MM.AAAA")
    bis: str = Field(description="Data final DD.MM.AAAA")
    nr: str = Field(description="Número da semana KWXX/AAAA")
    montag: str = Field(description="Segunda-feira em alemão técnico baseado nos exemplos")
    dienstag: str = Field(description="Terça-feira em alemão técnico baseado nos exemplos")
    mittwoch: str = Field(description="Quarta-feira em alemão técnico baseado nos exemplos")
    donnerstag: str = Field(description="Quinta-feira em alemão técnico baseado nos exemplos")
    freitag: str = Field(description="Sexta-feira em alemão técnico baseado nos exemplos")


class AppGeradorRelatorio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador Inteligente de Ausbildungsnachweis (Multimodal RAG & Auto-KW)")
        self.geometry("900x860")
        self.minsize(760, 680)
        self.resizable(False, False)

        self.pasta_raiz = Path(__file__).parent.resolve()
        self.caminho_doc = self.pasta_raiz / NOME_ARQUIVO_WORD
        self.caminho_json = self.pasta_raiz / "dados_relatorio.json"
        self.caminho_vbs = self.pasta_raiz / "gerar_copia_e_zerar.vbs"

        self.dados_gerados = {}
        self.data_selecionada = datetime.date.today()

        self._criar_layout()
        self._definir_data(self.data_selecionada)

    def _criar_layout(self):
        # 1. Configurações e Base de Conhecimento Multimodal
        frame_config = ttk.LabelFrame(self, text=" Configurações e Base Multimodal ", padding=10)
        frame_config.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_config, text="Gemini API Key:").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_api_key = ttk.Entry(frame_config, width=54, show="*")
        self.ent_api_key.insert(0, CHAVE_API_PADRAO)
        self.ent_api_key.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        self.lbl_exemplos = ttk.Label(frame_config, text="Status dos Exemplos: verificando...", foreground="blue")
        self.lbl_exemplos.grid(row=1, column=0, columnspan=2, sticky="w", pady=4)
        self._atualizar_status_exemplos()

        # 2. Dados da Semana com Seleção Inteligente de Datas e KW Automático
        frame_dados = ttk.LabelFrame(self, text=" Período e Assunto da Semana (com Auto-KW) ", padding=10)
        frame_dados.pack(fill="x", padx=10, pady=5)

        # 2.1 Linha de Navegação Rápida de Semana e Calendário Interativo
        frame_nav = ttk.Frame(frame_dados)
        frame_nav.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(0, 8))

        btn_sem_ant = ttk.Button(frame_nav, text="◀ Semana Anterior", command=self._semana_anterior)
        btn_sem_ant.pack(side="left", padx=2)

        btn_sem_atual = ttk.Button(frame_nav, text="📅 Esta Semana (Hoje)", command=self._semana_atual)
        btn_sem_atual.pack(side="left", padx=2)

        btn_sem_prox = ttk.Button(frame_nav, text="Semana Seguinte ▶", command=self._semana_seguinte)
        btn_sem_prox.pack(side="left", padx=2)

        ttk.Label(frame_nav, text=" |  Escolher no Calendário:").pack(side="left", padx=(10, 5))
        if HAS_TKCALENDAR:
            self.cal_entry = DateEntry(
                frame_nav,
                width=12,
                background="darkblue",
                foreground="white",
                borderwidth=2,
                date_pattern="dd.mm.yyyy",
            )
            self.cal_entry.pack(side="left", padx=2)
            self.cal_entry.bind("<<DateEntrySelected>>", self._on_calendario_selecionado)
        else:
            self.cal_entry = None

        # 2.2 Campos de Data Início (vom), Data Fim (bis) e Semana (KW)
        ttk.Label(frame_dados, text="Data Início (vom):").grid(row=1, column=0, sticky="w", pady=4)
        self.ent_vom = ttk.Entry(frame_dados, width=14)
        self.ent_vom.grid(row=1, column=1, sticky="w", padx=5, pady=4)
        self.ent_vom.bind("<FocusOut>", self._on_vom_digitado)
        self.ent_vom.bind("<Return>", self._on_vom_digitado)

        ttk.Label(frame_dados, text="Data Fim (bis):").grid(row=1, column=2, sticky="w", pady=4)
        self.ent_bis = ttk.Entry(frame_dados, width=14)
        self.ent_bis.grid(row=1, column=3, sticky="w", padx=5, pady=4)

        ttk.Label(frame_dados, text="Semana (Nr.):").grid(row=1, column=4, sticky="w", pady=4)
        self.ent_kw = ttk.Entry(frame_dados, width=14)
        self.ent_kw.grid(row=1, column=5, sticky="w", padx=5, pady=4)

        # 2.3 Resumo das Atividades
        ttk.Label(frame_dados, text="Resumo das Atividades Feitas na Semana (em português ou tópicos):").grid(
            row=2, column=0, columnspan=6, sticky="w", pady=(8, 2)
        )
        self.txt_tema = tk.Text(frame_dados, height=3, wrap="word", font=("Arial", 10))
        self.txt_tema.insert(
            "1.0",
            "Troca de selos em cilindros hidráulicos, alinhamento com relógio comparador e confecção de buchas no torno.",
        )
        self.txt_tema.grid(row=3, column=0, columnspan=6, sticky="ew", pady=2)

        # 3. Botões de Ação
        frame_botoes = ttk.Frame(self, padding=5)
        frame_botoes.pack(fill="x", padx=10, pady=5)

        self.btn_gerar = ttk.Button(
            frame_botoes, text="1. Analisar Base & Gerar com IA", command=self._iniciar_thread_geracao
        )
        self.btn_gerar.pack(side="left", padx=3, expand=True, fill="x")

        self.btn_injetar = ttk.Button(
            frame_botoes, text="2. Injetar no Word (.docm)", command=self._iniciar_thread_injecao, state="disabled"
        )
        self.btn_injetar.pack(side="left", padx=3, expand=True, fill="x")

        self.btn_zerar = ttk.Button(
            frame_botoes, text="3. Salvar Cópia & Zerar Modelo", command=self._iniciar_thread_backup_zerar
        )
        self.btn_zerar.pack(side="left", padx=3, expand=True, fill="x")

        # 4. Preview
        frame_preview = ttk.LabelFrame(self, text=" Pré-visualização Adaptada ao seu Estilo Técnico ", padding=10)
        frame_preview.pack(fill="both", expand=True, padx=10, pady=5)

        self.txt_preview = scrolledtext.ScrolledText(frame_preview, wrap="word", font=("Consolas", 10))
        self.txt_preview.pack(fill="both", expand=True)

        self.lbl_status = ttk.Label(self, text="Pronto.", relief="sunken", anchor="w", padding=3)
        self.lbl_status.pack(fill="x", side="bottom")

    # -------------------------------------------------------------
    # Gerenciamento Inteligente de Datas e Cálculo Automático do KW
    # -------------------------------------------------------------
    def _definir_data(self, data_ref: datetime.date):
        """Calcula automaticamente Segunda (vom), Sexta (bis) e o número da semana (KW)."""
        self.data_selecionada = data_ref

        # Segunda-feira da semana
        segunda = data_ref - datetime.timedelta(days=data_ref.weekday())
        # Sexta-feira da semana
        sexta = segunda + datetime.timedelta(days=4)

        # Cálculo da Semana ISO (KWxx/AAAA)
        ano, semana_iso, _ = segunda.isocalendar()
        kw_str = f"KW{semana_iso:02d}/{ano}"

        # Atualiza campos de entrada
        self.ent_vom.delete(0, "end")
        self.ent_vom.insert(0, segunda.strftime("%d.%m.%Y"))

        self.ent_bis.delete(0, "end")
        self.ent_bis.insert(0, sexta.strftime("%d.%m.%Y"))

        self.ent_kw.delete(0, "end")
        self.ent_kw.insert(0, kw_str)

        # Sincroniza o calendário se existir
        if self.cal_entry:
            try:
                self.cal_entry.set_date(data_ref)
            except Exception:
                pass

        self._set_status(f"Período selecionado: {kw_str} ({segunda.strftime('%d.%m.%Y')} até {sexta.strftime('%d.%m.%Y')})")

    def _semana_anterior(self):
        nova_data = self.data_selecionada - datetime.timedelta(days=7)
        self._definir_data(nova_data)

    def _semana_atual(self):
        self._definir_data(datetime.date.today())

    def _semana_seguinte(self):
        nova_data = self.data_selecionada + datetime.timedelta(days=7)
        self._definir_data(nova_data)

    def _on_calendario_selecionado(self, event=None):
        if self.cal_entry:
            dt = self.cal_entry.get_date()
            self._definir_data(dt)

    def _on_vom_digitado(self, event=None):
        texto = self.ent_vom.get().strip()
        formatos = ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")
        for fmt in formatos:
            try:
                dt = datetime.datetime.strptime(texto, fmt).date()
                self._definir_data(dt)
                return
            except ValueError:
                pass

    def _atualizar_status_exemplos(self):
        pasta = self.pasta_raiz / "exemplos_antigos"
        if not pasta.exists():
            pasta.mkdir(parents=True, exist_ok=True)
            arquivos = []
        else:
            extensoes = ("*.docx", "*.docm", "*.pdf", "*.png", "*.jpg", "*.jpeg", "*.webp", "*.txt")
            arquivos = []
            for ext in extensoes:
                arquivos.extend(
                    [a for a in pasta.glob(ext) if not a.name.startswith("~$") and not a.name.lower().startswith("readme")]
                )

        if arquivos:
            self.lbl_exemplos.config(
                text=f"📁 Base Multimodal: {len(arquivos)} arquivo(s) (Word, PDF, Fotos/TXT) em /exemplos_antigos",
                foreground="darkgreen",
            )
        else:
            self.lbl_exemplos.config(
                text="📁 Base Multimodal: Nenhum arquivo em /exemplos_antigos (adicione .docx, .pdf, fotos para Few-Shot)",
                foreground="#b36b00",
            )

    def _set_status(self, msg: str):
        self.lbl_status.config(text=msg)
        self.update_idletasks()

    def _iniciar_thread_geracao(self):
        threading.Thread(target=self._gerar_conteudo_ia, daemon=True).start()

    def _gerar_conteudo_ia(self):
        api_key = self.ent_api_key.get().strip()
        if not api_key:
            messagebox.showerror("Erro", "Insira a chave da API do Gemini.")
            return

        self.btn_gerar.config(state="disabled")
        self.btn_injetar.config(state="disabled")
        self._set_status("Processando base multimodal (Word, PDFs, Fotos) e consultando o Gemini...")

        try:
            exemplos_contexto = carregar_base_multimodal("exemplos_antigos", api_key=api_key)

            client = genai.Client(api_key=api_key)

            system_instruction = (
                "Você é um assistente técnico especializado na redação de relatórios semanais "
                "(Ausbildungsnachweis) para Industriemechaniker (Mecânica Industrial).\n"
                "Sua missão principal é IMITAR O ESTILO, a terminologia alemã e a concisão técnica "
                "dos exemplos fornecidos abaixo, mantendo total rigor gramatical e vocabulário industrial real.\n"
                f"{exemplos_contexto}"
            )

            prompt = f"""
            Gere as descrições diárias (Montag bis Freitag) baseando-se nas atividades informadas:
            - Período: vom {self.ent_vom.get().strip()} bis {self.ent_bis.get().strip()} ({self.ent_kw.get().strip()})
            - Assunto: {self.txt_tema.get("1.0", "end-1c").strip()}

            Regras:
            1. Mantenha o mesmo tom e padrão de formulação dos exemplos anteriores.
            2. Forneça sentenças claras com foco em métodos, ferramentas e segurança.
            """

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json",
                    response_schema=AusbildungsnachweisSchema,
                    temperature=0.25,
                ),
            )

            self.dados_gerados = json.loads(response.text)

            with open(self.caminho_json, "w", encoding="utf-8") as f:
                json.dump(self.dados_gerados, f, ensure_ascii=False, indent=2)

            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("end", json.dumps(self.dados_gerados, ensure_ascii=False, indent=2))

            self._set_status("Sucesso! Relatório gerado com base no seu estilo e gravado no JSON.")
            self.btn_injetar.config(state="normal")
            self._atualizar_status_exemplos()

        except Exception as e:
            messagebox.showerror("Erro na Geração", f"Falha:\n{e}")
            self._set_status("Erro ao processar.")
        finally:
            self.btn_gerar.config(state="normal")

    def _iniciar_thread_injecao(self):
        threading.Thread(target=self._injetar_no_word, daemon=True).start()

    def _injetar_no_word(self):
        if not self.caminho_doc.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{self.caminho_doc}")
            return

        self.btn_injetar.config(state="disabled")
        self._set_status("Conectando ao Word e executando macro...")

        # Inicializa o subsistema COM para a thread secundária
        pythoncom.CoInitialize()
        try:
            # 1. Tenta conectar a uma instância do Word já aberta ou inicia uma nova
            try:
                word = win32.GetActiveObject("Word.Application")
            except Exception:
                try:
                    word = win32.Dispatch("Word.Application")
                except Exception:
                    word = win32.gencache.EnsureDispatch("Word.Application")

            word.Visible = True

            # 2. Verifica se o documento já está aberto no Word
            doc = None
            for d in word.Documents:
                try:
                    if d.FullName.lower() == str(self.caminho_doc).lower() or d.Name.lower() == self.caminho_doc.name.lower():
                        doc = d
                        break
                except Exception:
                    pass

            if doc is None:
                doc = word.Documents.Open(str(self.caminho_doc))

            # 3. Executa a macro VBA que carrega o dados_relatorio.json
            word.Run("CarregarDadosDoJSON")
            doc.Save()

            self._set_status("Relatório atualizado e salvo com sucesso no Word!")
            messagebox.showinfo("Concluído", "Word preenchido e salvo com sucesso!")
        except Exception as e:
            mensagem_erro = str(e)
            if "permissão" in mensagem_erro.lower() or "readonly" in mensagem_erro.lower() or "-2147352567" in mensagem_erro:
                mensagem_erro += (
                    "\n\n👉 Dica: O arquivo está bloqueado por outro processo do Word aberto em segundo plano.\n"
                    "Feche todas as janelas do Word (ou encerre o processo WINWORD.EXE no Gerenciador de Tarefas) e tente novamente."
                )
            messagebox.showerror("Erro no Word", f"Falha na automação do Word:\n{mensagem_erro}")
            self._set_status("Erro no Word.")
        finally:
            pythoncom.CoUninitialize()
            self.btn_injetar.config(state="normal")

    def _iniciar_thread_backup_zerar(self):
        confirmar = messagebox.askyesno(
            "Confirmar Backup e Limpeza",
            "Deseja criar uma cópia de backup do relatório atual na pasta 'copias/' e zerar o documento principal para a próxima semana?",
        )
        if confirmar:
            threading.Thread(target=self._executar_backup_zerar, daemon=True).start()

    def _executar_backup_zerar(self):
        if not self.caminho_vbs.exists():
            messagebox.showerror("Erro", f"Arquivo VBScript não encontrado:\n{self.caminho_vbs}")
            return

        self._set_status("Criando cópia em /copias e zerando documento principal...")
        try:
            subprocess.run(["wscript.exe", str(self.caminho_vbs)])
            self._set_status("Backup realizado em /copias e documento principal limpo!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao executar script de backup:\n{e}")
            self._set_status("Erro no processo de backup.")


if __name__ == "__main__":
    app = AppGeradorRelatorio()
    app.mainloop()
