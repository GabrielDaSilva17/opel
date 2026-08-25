"""Ponto de entrada principal - Inicia a Central Inteligente de Relatórios (SQLite + Gemini)."""
import sys
from pathlib import Path

# Garante que a pasta raiz do projeto está no sys.path
PASTA_RAIZ = Path(__file__).parent.resolve()
if str(PASTA_RAIZ) not in sys.path:
    sys.path.insert(0, str(PASTA_RAIZ))

from app_central import AppCentral

if __name__ == "__main__":
    app = AppCentral()
    app.mainloop()
