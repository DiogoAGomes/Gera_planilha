import os
import sys
from streamlit.web import cli as stcli


def resource_path(relative_path):
    """Resolve o caminho do arquivo tanto em execução normal (python launcher.py)
    quanto quando empacotado pelo PyInstaller em modo --onefile."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    app_path = resource_path("gera_planilha2.py")
    sys.argv = [
        "streamlit", "run", app_path,
        "--global.developmentMode=false",
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())
