import json
import os
import subprocess
import time
from pathlib import Path
import pythoncom
import win32com.client as win32


class WordController:
    """Gerencia a automação segura do Microsoft Word com controle de processos e permissões."""

    @staticmethod
    def verificar_word_em_execucao() -> bool:
        """Verifica se o Microsoft Word está aberto e respondendo."""
        pythoncom.CoInitialize()
        try:
            word = win32.GetActiveObject("Word.Application")
            return word is not None
        except Exception:
            return False
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def limpar_travas_temporarias(pasta: Path = None):
        """Remove arquivos temporários residuais de travamento (~$*) se o Word não estiver em execução."""
        if pasta is None:
            pasta = Path(__file__).parent
        try:
            if not WordController.verificar_word_em_execucao():
                for arq in pasta.glob("~*"):
                    if arq.is_file():
                        try:
                            arq.unlink()
                        except Exception:
                            pass
        except Exception:
            pass

    @staticmethod
    def finalizar_processos_fantasmas_word():
        """Finaliza instâncias invisíveis ou travadas do WINWORD.EXE em segundo plano."""
        try:
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, text=True)
            time.sleep(0.5)
            WordController.limpar_travas_temporarias()
        except Exception as e:
            print(f"Aviso ao finalizar processos do Word: {e}")

    @staticmethod
    def verificar_arquivo_bloqueado(caminho_arquivo: Path) -> tuple[bool, str]:
        """
        Verifica se o arquivo existe e se está bloqueado para escrita por outro processo.
        Retorna (bloqueado: bool, mensagem: str).
        """
        if not caminho_arquivo.exists():
            return True, f"Arquivo não encontrado: {caminho_arquivo}"

        # Limpa travas se o Word não estiver rodando
        WordController.limpar_travas_temporarias(caminho_arquivo.parent)

        # Tenta abrir o arquivo para verificar integridade de acesso
        try:
            with open(caminho_arquivo, "r+b"):
                pass
            return False, "Arquivo livre para edição."
        except PermissionError:
            return True, "O arquivo está atualmente em uso pelo Word ou bloqueado por permissões do Windows."
        except Exception as e:
            return False, f"Aviso de leitura: {e}"

    @staticmethod
    def abrir_ou_obter_word(caminho_arquivo: Path, trazer_para_frente: bool = True):
        """
        Obtém uma instância ativa do Word ou abre uma nova visível em primeiro plano.
        Retorna (word_app, documento).
        """
        caminho_str = str(caminho_arquivo.resolve())
        word = None

        # 1. Tenta obter o Word ativo ou inicia novo processo
        try:
            word = win32.GetActiveObject("Word.Application")
        except Exception:
            try:
                word = win32.Dispatch("Word.Application")
            except Exception:
                word = win32.gencache.EnsureDispatch("Word.Application")

        # Garante que o Word está sempre VISÍVEL para que permissões e alertas de macros possam ser respondidos
        word.Visible = True
        word.DisplayAlerts = -1  # wdAlertsAll: permite que o usuário veja avisos de macro e modo de exibição protegida

        # 2. Localiza o documento se já estiver aberto
        doc = None
        for d in word.Documents:
            try:
                if d.FullName.lower() == caminho_str.lower() or d.Name.lower() == caminho_arquivo.name.lower():
                    doc = d
                    break
            except Exception:
                pass

        # 3. Se não estiver aberto, abre o documento no Word visível
        if doc is None:
            doc = word.Documents.Open(caminho_str)

        # 4. Traz para primeiro plano se solicitado para permitir que o usuário veja o pedido de permissões
        if trazer_para_frente:
            try:
                word.Activate()
                if doc:
                    doc.Activate()
            except Exception:
                pass

        return word, doc

    @classmethod
    def abrir_documento_para_edicao(cls, caminho_arquivo: Path) -> tuple[bool, str]:
        """Abre o documento do Word em primeiro plano para o usuário visualizar ou conceder permissões."""
        if not caminho_arquivo.exists():
            return False, f"Arquivo Word não encontrado:\n{caminho_arquivo}"

        # Se houver travas residuais sem Word rodando, limpa
        cls.limpar_travas_temporarias(caminho_arquivo.parent)

        pythoncom.CoInitialize()
        try:
            word, doc = cls.abrir_ou_obter_word(caminho_arquivo, trazer_para_frente=True)
            return True, "Word aberto com sucesso em primeiro plano."
        except Exception as e:
            return False, f"Erro ao abrir Word:\n{str(e)}"
        finally:
            pythoncom.CoUninitialize()

    @classmethod
    def injetar_dados_e_executar_macro(
        cls,
        caminho_doc: Path,
        caminho_json: Path,
        dados: dict,
        nome_macro: str = "CarregarDadosDoJSON",
        trazer_para_frente: bool = True,
    ) -> tuple[bool, str]:
        """
        Salva o arquivo JSON de suporte, garante que o Word está aberto em primeiro plano,
        e dispara a macro VBA com tratamento seguro de erros e permissões.
        """
        if not caminho_doc.exists():
            return False, f"Documento Word não encontrado em:\n{caminho_doc}"

        # 1. Salva os dados no JSON
        try:
            with open(caminho_json, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)
        except Exception as e:
            return False, f"Falha ao salvar dados intermediários em JSON:\n{e}"

        # 2. Conecta ao Word e executa
        pythoncom.CoInitialize()
        try:
            word, doc = cls.abrir_ou_obter_word(caminho_doc, trazer_para_frente=trazer_para_frente)

            # Executa a macro VBA contida no documento .docm
            try:
                word.Run(nome_macro)
            except Exception as macro_err:
                msg_err = str(macro_err)
                if "macro" in msg_err.lower() or "-2147352567" in msg_err:
                    return False, (
                        "Aviso de Segurança / Macro do Word:\n\n"
                        "O Word pode estar solicitando 'Habilitar Conteúdo / Macros' na barra amarela superior.\n"
                        "Por favor, clique em 'Habilitar Conteúdo' no Word e tente injetar novamente."
                    )
                raise macro_err

            # Salva o documento após preenchimento
            if doc:
                doc.Save()

            return True, "Documento Word atualizado e salvo com sucesso via VBA!"

        except Exception as e:
            erro_str = str(e)
            if "permissão" in erro_str.lower() or "readonly" in erro_str.lower() or "bloqueado" in erro_str.lower():
                return False, (
                    "O documento está aberto como somente-leitura ou bloqueado por outra janela.\n"
                    "Dica: Feche janelas duplicadas do Word e tente novamente."
                )
            return False, f"Falha na automação do Word:\n{erro_str}"
        finally:
            pythoncom.CoUninitialize()
