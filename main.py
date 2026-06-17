import sys

from PySide6.QtWidgets import QApplication

from app.core.settings import ensure_directories
from app.ui.main_window import MainWindow


def main() -> int:
    ensure_directories()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
