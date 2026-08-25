import sqlite3
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    DB_NAME = Path(sys.executable).parent / "relatorios.db"
else:
    DB_NAME = Path(__file__).parent / "relatorios.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Tabela principal de relatórios semanais
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS relatorios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
        vom TEXT,
        bis TEXT,
        kw TEXT,
        origem_app TEXT,
        prompt_usado TEXT,
        montag TEXT,
        dienstag TEXT,
        mittwoch TEXT,
        donnerstag TEXT,
        freitag TEXT,
        observacoes TEXT
    )
    """)

    # Tabela de base de conhecimento (extraída de fotos, prints e documentos)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS base_conhecimento (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP,
        origem_app TEXT,
        arquivo_origem TEXT,
        conteudo_extraido TEXT,
        ativo INTEGER DEFAULT 1
    )
    """)

    # Garante que a coluna ativo existe se a tabela já foi criada antes
    try:
        cursor.execute("ALTER TABLE base_conhecimento ADD COLUMN ativo INTEGER DEFAULT 1")
    except sqlite3.OperationalError:
        pass

    conn.commit()
    conn.close()


def salvar_relatorio(vom, bis, kw, origem_app, prompt, m, d, mi, don, f, obs=""):
    """Salva ou atualiza um relatório no banco SQLite."""
    return salvar_ou_atualizar_relatorio(vom, bis, kw, origem_app, prompt, m, d, mi, don, f, obs)


def salvar_ou_atualizar_relatorio(vom, bis, kw, origem_app, prompt, m, d, mi, don, f, obs=""):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM relatorios WHERE kw = ? AND vom = ?", (kw, vom))
    existente = cursor.fetchone()
    if existente:
        cursor.execute("""
        UPDATE relatorios 
        SET bis = ?, origem_app = ?, prompt_usado = ?, 
            montag = ?, dienstag = ?, mittwoch = ?, donnerstag = ?, freitag = ?, 
            observacoes = ?, data_registro = CURRENT_TIMESTAMP
        WHERE id = ?
        """, (bis, origem_app, prompt, m, d, mi, don, f, obs, existente[0]))
        novo_id = existente[0]
    else:
        cursor.execute("""
        INSERT INTO relatorios (
            vom, bis, kw, origem_app, prompt_usado, 
            montag, dienstag, mittwoch, donnerstag, freitag, observacoes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (vom, bis, kw, origem_app, prompt, m, d, mi, don, f, obs))
        novo_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return novo_id


def salvar_conhecimento(origem_app, arquivo_origem, conteudo, ativo=1):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Se já existir registro para este arquivo, atualiza
    cursor.execute("SELECT id FROM base_conhecimento WHERE arquivo_origem = ?", (arquivo_origem,))
    linha = cursor.fetchone()
    if linha:
        cursor.execute("""
        UPDATE base_conhecimento 
        SET conteudo_extraido = ?, origem_app = ?, data_importacao = CURRENT_TIMESTAMP, ativo = ?
        WHERE id = ?
        """, (conteudo, origem_app, ativo, linha[0]))
        novo_id = linha[0]
    else:
        cursor.execute("""
        INSERT INTO base_conhecimento (origem_app, arquivo_origem, conteudo_extraido, ativo)
        VALUES (?, ?, ?, ?)
        """, (origem_app, arquivo_origem, conteudo, ativo))
        novo_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return novo_id


def arquivo_ja_importado(arquivo_origem: str) -> bool:
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM base_conhecimento WHERE arquivo_origem = ? LIMIT 1", (arquivo_origem,))
    existe = cursor.fetchone() is not None
    conn.close()
    return existe


def listar_nomes_arquivos_importados() -> set[str]:
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT arquivo_origem FROM base_conhecimento WHERE arquivo_origem IS NOT NULL")
    linhas = cursor.fetchall()
    conn.close()
    return {linha[0] for linha in linhas if linha[0]}


def obter_status_ativo_arquivos() -> dict[str, bool]:
    """Retorna um dicionário {nome_arquivo: ativo_bool}."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT arquivo_origem, ativo FROM base_conhecimento WHERE arquivo_origem IS NOT NULL")
    linhas = cursor.fetchall()
    conn.close()
    return {linha[0]: bool(linha[1]) for linha in linhas if linha[0]}


def alternar_status_conhecimento(arquivo_origem: str) -> bool:
    """Alterna o status de ativo/inativo de um arquivo para o aprendizado da IA."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT ativo FROM base_conhecimento WHERE arquivo_origem = ?", (arquivo_origem,))
    linha = cursor.fetchone()
    novo_status = 1
    if linha:
        novo_status = 0 if linha[0] == 1 else 1
        cursor.execute("UPDATE base_conhecimento SET ativo = ? WHERE arquivo_origem = ?", (novo_status, arquivo_origem))
    conn.commit()
    conn.close()
    return bool(novo_status)


def excluir_conhecimento_arquivo(arquivo_origem: str):
    """Remove um arquivo do banco de conhecimento."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM base_conhecimento WHERE arquivo_origem = ?", (arquivo_origem,))
    conn.commit()
    conn.close()


def obter_conteudo_por_arquivo(arquivo_origem: str) -> str | None:
    """Retorna o texto extraído de um arquivo salvo no SQLite."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT conteudo_extraido FROM base_conhecimento WHERE arquivo_origem = ? ORDER BY id DESC LIMIT 1", (arquivo_origem,))
    linha = cursor.fetchone()
    conn.close()
    return linha[0] if linha else None


def listar_conhecimento():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, data_importacao, origem_app, arquivo_origem, substr(conteudo_extraido, 1, 80), ativo FROM base_conhecimento ORDER BY id DESC")
    dados = cursor.fetchall()
    conn.close()
    return dados


def listar_relatorios():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, data_registro, kw, vom, bis, origem_app FROM relatorios ORDER BY id DESC")
    dados = cursor.fetchall()
    conn.close()
    return dados


def pesquisar_relatorios(termo: str = ""):
    """
    Pesquisa relatórios pelo ID, KW, período (vom/bis), departamento ou texto do prompt/dias.
    """
    termo = (termo or "").strip()
    if not termo:
        return listar_relatorios()

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    if termo.isdigit():
        termo_num = int(termo)
        cursor.execute(
            """
            SELECT id, data_registro, kw, vom, bis, origem_app 
            FROM relatorios 
            WHERE id = ? OR kw LIKE ? OR kw LIKE ?
            ORDER BY id DESC
            """,
            (termo_num, f"%KW{termo_num:02d}%", f"%{termo}%")
        )
    else:
        padrao = f"%{termo}%"
        cursor.execute(
            """
            SELECT id, data_registro, kw, vom, bis, origem_app 
            FROM relatorios 
            WHERE kw LIKE ? 
               OR vom LIKE ? 
               OR bis LIKE ? 
               OR origem_app LIKE ? 
               OR prompt_usado LIKE ? 
               OR montag LIKE ? 
               OR dienstag LIKE ? 
               OR mittwoch LIKE ? 
               OR donnerstag LIKE ? 
               OR freitag LIKE ?
            ORDER BY id DESC
            """,
            (padrao, padrao, padrao, padrao, padrao, padrao, padrao, padrao, padrao, padrao)
        )

    dados = cursor.fetchall()
    conn.close()
    return dados


def obter_relatorio_por_id(relatorio_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM relatorios WHERE id = ?", (relatorio_id,))
    dado = cursor.fetchone()
    conn.close()
    return dado


def excluir_relatorio(relatorio_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM relatorios WHERE id = ?", (relatorio_id,))
    conn.commit()
    conn.close()


def obter_textos_conhecimento():
    """Retorna apenas os textos de arquivos que estão marcados como ATIVOS para aprendizado."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT conteudo_extraido FROM base_conhecimento WHERE ativo = 1 ORDER BY id DESC LIMIT 30")
    dados = cursor.fetchall()
    conn.close()
    return [d[0] for d in dados if d[0]]


# Inicializa o banco automaticamente ao importar
init_db()
