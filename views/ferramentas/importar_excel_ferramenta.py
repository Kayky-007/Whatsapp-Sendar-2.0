# views/ferramentas/importar_excel_ferramenta.py
from PySide6.QtWidgets import QDialog, QFormLayout, QLineEdit, QPushButton, QFileDialog, QHBoxLayout

class ImportarExcelFerramenta(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importar Contatos do Excel")
        self.setGeometry(200, 200, 400, 200)

        layout = QFormLayout()

        self.input_arquivo = QLineEdit()
        self.input_arquivo.setPlaceholderText("Selecione o arquivo .xlsx")
        self.input_arquivo.setReadOnly(True)
        btn_selecionar = QPushButton("Selecionar Arquivo")
        btn_selecionar.clicked.connect(self._selecionar_arquivo)
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.input_arquivo)
        file_layout.addWidget(btn_selecionar)
        layout.addRow("Arquivo Excel:", file_layout)

        self.input_col_nome = QLineEdit()
        self.input_col_nome.setPlaceholderText("Ex.: Nome (opcional)")
        layout.addRow("Coluna com Nomes:", self.input_col_nome)

        self.input_col_numero = QLineEdit()
        self.input_col_numero.setPlaceholderText("Ex.: Número")
        layout.addRow("Coluna com Números:", self.input_col_numero)

        btn_ok = QPushButton("Importar")
        btn_ok.clicked.connect(self.accept)
        layout.addWidget(btn_ok)

        self.setLayout(layout)

    def _selecionar_arquivo(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Excel", "", "Arquivos Excel (*.xlsx)")
        if caminho:
            self.input_arquivo.setText(caminho)

    def get_dados(self):
        return (
            self.input_arquivo.text().strip(),
            self.input_col_nome.text().strip(),
            self.input_col_numero.text().strip()
        )