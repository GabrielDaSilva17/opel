import json
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

if getattr(sys, "frozen", False):
    CONFIG_FILE = Path(sys.executable).parent / "config.json"
    ENV_FILE = Path(sys.executable).parent / ".env"
else:
    CONFIG_FILE = Path(__file__).parent / "config.json"
    ENV_FILE = Path(__file__).parent / ".env"

# Dicionário padrão de configurações
CONFIG_PADRAO = {
    "idioma_interface": "pt",
    "idioma_relatorio": "de",
    "gemini_api_key": "AQ.Ab8RN6IS3ZmzQm8J5zjABuinQ-75pFu0XPYye5CbRPKdrTQg7Q",
    "gemini_modelo": "gemini-3.5-flash",
    "temperatura": 0.7,
    "arquivo_word": "doc ausbildund.docm",
    "pasta_copias": "copias",
    "pasta_exemplos": "exemplos_antigos",
    "tema": "dark",
    "cor_destaque": "blue",
    "auto_focar_word": True,
    "origem_padrao": "Oficina Geral / Opel",
}

# Textos para suporte a múltiplos idiomas (i18n)
TEXTOS_I18N = {
    "pt": {
        "titulo_app": "Ausbildungsnachweis AI - Central de Relatórios & RAG Local",
        "subtitulo_app": "Central Inteligente com RAG Local, Gemini Vision & SQLite | Opel Mecânica Industrial",
        "badge_db": "🟢 SQLite Ativo",
        "badge_sync": "📁 Sincronizando...",
        "badge_word_pronto": "🔵 Word Aberto",
        "badge_word_fechado": "⚪ Word Fechado",
        "btn_config": "⚙️ Configurações",
        "btn_abrir_word": "👁️ Abrir Word",
        "tab_gerador": " 📝 Editor & Gerador IA ",
        "tab_extrator": " ⚡ Auto-Importador de Arquivos ",
        "tab_historico": " 🗄️ Histórico SQLite ",
        "card_periodo": "📅 Período e Identificação da Semana",
        "btn_sem_ant": "◀ Semana Anterior",
        "btn_sem_hoje": "📅 Esta Semana (Hoje)",
        "btn_sem_prox": "Semana Seguinte ▶",
        "lbl_vom": "vom (Início):",
        "lbl_bis": "bis (Fim):",
        "lbl_kw": "Semana (KW):",
        "lbl_origem": "Origem / App:",
        "card_prompt": "🎯 Atividades Realizadas na Semana (em tópicos ou resumo):",
        "btn_gerar": "✨ 1. Gerar com IA (Gemini)",
        "btn_limpar": "🧹 Limpar Formulário",
        "btn_salvar": "💾 2. Salvar no SQLite",
        "btn_injetar": "📄 3. Injetar no Word (.docm)",
        "btn_vbs": "🔄 4. Copiar & Zerar (VBS)",
        "tab_modo_dia": " 📅 Modo 1: Edição Dia a Dia ",
        "tab_modo_livre": " 📄 Modo 2: Edição em Texto Único Livre ",
        "btn_sync_para_livre": "➡️ Sincronizar para o Modo Texto Único",
        "btn_sync_para_dias": "⬅️ Converter Texto Único para as Caixas Diárias",
        "msg_status_pronto": "Sistema pronto para gerar relatórios.",
        "msg_limpar_confirma": "Deseja realmente limpar todos os campos do formulário?",
        "msg_limpo_sucesso": "Formulário limpo com sucesso!",
    },
    "de": {
        "titulo_app": "Ausbildungsnachweis AI - Intelligente Berichts-Zentrale",
        "subtitulo_app": "Intelligente Zentrale mit lokalem RAG, Gemini Vision & SQLite | Opel Industriemechaniker",
        "badge_db": "🟢 SQLite Aktiv",
        "badge_sync": "📁 Synchronisiere...",
        "badge_word_pronto": "🔵 Word Geöffnet",
        "badge_word_fechado": "⚪ Word Geschlossen",
        "btn_config": "⚙️ Einstellungen",
        "btn_abrir_word": "👁️ Word Öffnen",
        "tab_gerador": " 📝 Editor & KI-Generator ",
        "tab_extrator": " ⚡ Datei-Auto-Importeur ",
        "tab_historico": " 🗄️ SQLite-Verlauf ",
        "card_periodo": "📅 Zeitraum und Wochenidentifikation",
        "btn_sem_ant": "◀ Vorherige Woche",
        "btn_sem_hoje": "📅 Diese Woche (Heute)",
        "btn_sem_prox": "Nächste Woche ▶",
        "lbl_vom": "vom (Start):",
        "lbl_bis": "bis (Ende):",
        "lbl_kw": "Woche (KW):",
        "lbl_origem": "Herkunft / Abt.:",
        "card_prompt": "🎯 Durchgeführte Tätigkeiten der Woche (Stichpunkte):",
        "btn_gerar": "✨ 1. Mit KI Generieren (Gemini)",
        "btn_limpar": "🧹 Formular Leeren",
        "btn_salvar": "💾 2. In SQLite Speichern",
        "btn_injetar": "📄 3. In Word Übertragen (.docm)",
        "btn_vbs": "🔄 4. Kopieren & Zurücksetzen",
        "tab_modo_dia": " 📅 Modus 1: Tagesansicht ",
        "tab_modo_livre": " 📄 Modus 2: Freitextansicht ",
        "btn_sync_para_livre": "➡️ In Freitext synchronisieren",
        "btn_sync_para_dias": "⬅️ In Tagesfelder umwandeln",
        "msg_status_pronto": "System bereit zur Berichterstellung.",
        "msg_limpar_confirma": "Möchten Sie wirklich alle Felder des Formulars leeren?",
        "msg_limpo_sucesso": "Formular erfolgreich zurückgesetzt!",
    },
    "en": {
        "titulo_app": "Ausbildungsnachweis AI - Report Center & Local RAG",
        "subtitulo_app": "Smart Center with Local RAG, Gemini Vision & SQLite | Opel Industrial Mechanics",
        "badge_db": "🟢 SQLite Active",
        "badge_sync": "📁 Syncing...",
        "badge_word_pronto": "🔵 Word Open",
        "badge_word_fechado": "⚪ Word Closed",
        "btn_config": "⚙️ Settings",
        "btn_abrir_word": "👁️ Open Word",
        "tab_gerador": " 📝 Editor & AI Generator ",
        "tab_extrator": " ⚡ Auto File Importer ",
        "tab_historico": " 🗄️ SQLite History ",
        "card_periodo": "📅 Period & Week Information",
        "btn_sem_ant": "◀ Previous Week",
        "btn_sem_hoje": "📅 This Week (Today)",
        "btn_sem_prox": "Next Week ▶",
        "lbl_vom": "vom (Start):",
        "lbl_bis": "bis (End):",
        "lbl_kw": "Week (KW):",
        "lbl_origem": "Source / Dept:",
        "card_prompt": "🎯 Activities performed this week (topics/summary):",
        "btn_gerar": "✨ 1. Generate with AI (Gemini)",
        "btn_limpar": "🧹 Clear Form",
        "btn_salvar": "💾 2. Save to SQLite",
        "btn_injetar": "📄 3. Inject into Word (.docm)",
        "btn_vbs": "🔄 4. Copy & Reset (VBS)",
        "tab_modo_dia": " 📅 Mode 1: Daily Breakdown ",
        "tab_modo_livre": " 📄 Mode 2: Full Free Text ",
        "btn_sync_para_livre": "➡️ Sync to Full Text",
        "btn_sync_para_dias": "⬅️ Convert Full Text to Daily Boxes",
        "msg_status_pronto": "System ready to generate reports.",
        "msg_limpar_confirma": "Do you really want to clear all form fields?",
        "msg_limpo_sucesso": "Form cleared successfully!",
    },
}


def carregar_config() -> dict:
    """Carrega as configurações do arquivo config.json ou cria com valores padrão."""
    load_dotenv(override=False)

    config = dict(CONFIG_PADRAO)

    # Lê variáveis do .env se existirem
    env_api_key = os.getenv("GEMINI_API_KEY")
    if env_api_key:
        config["gemini_api_key"] = env_api_key

    env_arquivo = os.getenv("ARQUIVO_WORD")
    if env_arquivo:
        config["arquivo_word"] = env_arquivo

    env_copias = os.getenv("PASTA_COPIAS")
    if env_copias:
        config["pasta_copias"] = env_copias

    # Se o arquivo config.json existir, mescla seus dados
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                dados_arquivo = json.load(f)
                config.update(dados_arquivo)
        except Exception as e:
            print(f"Aviso ao ler config.json: {e}")
    else:
        salvar_config(config)

    return config


def salvar_config(novas_configs: dict):
    """Salva o dicionário de configurações no arquivo config.json e sincroniza com .env."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(novas_configs, f, ensure_ascii=False, indent=2)

        # Sincroniza o .env para compatibilidade
        linhas_env = [
            f"GEMINI_API_KEY={novas_configs.get('gemini_api_key', '')}",
            f"ARQUIVO_WORD={novas_configs.get('arquivo_word', 'doc ausbildund.docm')}",
            f"PASTA_COPIAS={novas_configs.get('pasta_copias', 'copias')}",
        ]
        with open(ENV_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas_env) + "\n")

    except Exception as e:
        print(f"Erro ao salvar configurações: {e}")
        raise e


def t(chave: str, idioma: str = "pt") -> str:
    """Retorna a tradução da chave para o idioma solicitado."""
    dicionario = TEXTOS_I18N.get(idioma, TEXTOS_I18N["pt"])
    return dicionario.get(chave, TEXTOS_I18N["pt"].get(chave, chave))


def testar_conexao_gemini(api_key: str, modelo: str = "gemini-flash-latest") -> tuple[bool, str]:
    """Testa a conectividade da chave de API com o Google Gemini com fallback automático."""
    if not api_key or not api_key.strip():
        return False, "Chave de API não informada."

    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key.strip())

        candidatos = [modelo, "gemini-flash-latest", "gemini-3.7-flash", "gemini-3.5-flash", "gemini-flash-lite-latest"]
        modelos_unicos = list(dict.fromkeys(candidatos))

        ultimo_erro = ""
        for mod in modelos_unicos:
            try:
                res = client.models.generate_content(
                    model=mod,
                    contents="Responda em uma palavra: OK",
                    config=types.GenerateContentConfig(
                        automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True)
                    ),
                )
                if res and res.text:
                    return True, f"Conexão OK com {mod}!"
            except Exception as mod_err:
                ultimo_erro = str(mod_err)

        return False, f"Erro ao conectar com Gemini: {ultimo_erro}"
    except Exception as e:
        return False, f"Erro ao conectar com Gemini: {str(e)}"
