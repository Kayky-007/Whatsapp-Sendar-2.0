# views/ferramentas/add_contato_ferramenta.py
from PySide6.QtWidgets import QDialog, QFormLayout, QLineEdit, QPushButton

class AddContatoFerramenta(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Adicionar Contato Manualmente")
        self.setGeometry(200, 200, 300, 200)

        layout = QFormLayout()
        self.input_nome = QLineEdit()
        self.input_numero = QLineEdit()
        self.input_numero.setPlaceholderText("+55DDDnúmero")
        layout.addRow("Nome:", self.input_nome)
        layout.addRow("Número:", self.input_numero)

        btn_ok = QPushButton("Adicionar")
        btn_ok.clicked.connect(self.accept)
        layout.addWidget(btn_ok)

        self.setLayout(layout)

    def get_contato(self):
        return self.input_nome.text().strip(), self.input_numero.text().strip()