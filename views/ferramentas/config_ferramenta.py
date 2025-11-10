# views/ferramentas/config_ferramenta.py
from PySide6.QtWidgets import QDialog, QFormLayout, QLineEdit, QPushButton

class ConfigFerramenta(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configurações")
        self.setGeometry(200, 200, 300, 150)

        layout = QFormLayout()
        self.input_delay = QLineEdit("2")
        layout.addRow("Delay entre mensagens (segundos):", self.input_delay)

        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(self.accept)
        layout.addWidget(btn_ok)

        self.setLayout(layout)

    def get_delay(self):
        try:
            return float(self.input_delay.text())
        except:
            return 2.0