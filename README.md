# Ausbildungsnachweis AI 📋🤖

Sistema inteligente para geração e gerenciamento automatizado de relatórios de formação (*Ausbildungsnachweis*), integrado com a API do Google Gemini e automação de documentos Word.

---

## 🚀 Funcionalidades

- **Central Inteligente:** Interface gráfica amigável para consulta, criação e edição de relatórios.
- **Integração com IA (Google Gemini):** Sugestão e geração de textos profissionais em alemão com base em atividades realizadas.
- **Extração de Exemplos:** Importação e vetorização de relatórios anteriores em PDF e Word para aprendizado de estilo.
- **Automação Word (COM):** Preenchimento automático direto no modelo `.docm` mantendo formatação e estrutura.
- **Banco de Dados SQLite Local:** Armazenamento seguro e organizado do histórico de relatórios.

---

## 🛠️ Pré-requisitos

- **Python 3.10+** (recomendado Python 3.11 ou superior)
- **Microsoft Word** (para automação via biblioteca `pywin32` no Windows)
- **Chave de API do Google Gemini**

---

## 📦 Instalação

1. Clone o repositório ou baixe a pasta do projeto.
2. Crie e ative um ambiente virtual:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```
3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
4. Configure as variáveis de ambiente:
   - Copie o arquivo `.env.example` para `.env`:
     ```bash
     copy .env.example .env
     ```
   - Abra o `.env` e insira sua chave da API do Gemini (`GEMINI_API_KEY`).

---

## ▶️ Como Executar

Para iniciar a aplicação:

```bash
python main.py
```

---

## 📂 Estrutura do Projeto

```text
├── app_central.py          # Interface principal do sistema
├── app_relatorio_gui.py    # Interface do gerador de relatórios
├── config_dialog.py        # Diálogo de configurações
├── config_manager.py       # Gerenciamento de credenciais e parâmetros
├── database.py             # Operações no banco de dados SQLite
├── editor_popup.py         # Janela de edição e ajuste fino de relatórios
├── extrator_exemplos.py    # Extrator de relatórios antigos (PDF/DOCX)
├── word_manager.py         # Automação do Microsoft Word via win32com
├── main.py                 # Ponto de entrada da aplicação
├── requirements.txt        # Dependências do projeto
├── .env.example            # Exemplo de configuração de ambiente
└── .gitignore              # Arquivos ignorados pelo Git
```

---

## 🔒 Segurança

Arquivos que contêm dados pessoais, chaves de API (`.env`), banco de dados local (`relatorios.db`) e arquivos gerados (`copias/`) estão configurados no `.gitignore` e não serão sincronizados com o repositório público.
