# main.py
from PySide6.QtWidgets import QApplication
import sys
from views.main_janela import MainJanela

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = MainJanela()
    janela.show()
    sys.exit(app.exec())