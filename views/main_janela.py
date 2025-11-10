# views/main_janela.py
from PySide6.QtWidgets import (
    QMainWindow, QTableWidget, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QTextEdit, QLabel, QMessageBox
)
from PySide6.QtCore import Qt
import webbrowser
from controllers.main_controller import MainController

class MainJanela(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhatsApp Sender")
        self.setGeometry(100, 100, 1200, 600)
        
        # 1. MONTAR A INTERFACE PRIMEIRO
        self._montar_interface()
        
        # 2. SÓ DEPOIS CRIAR O CONTROLLER
        self.controller = MainController(self)
        
        # 3. Conectar sinais (opcional, já no controller)
        self._conectar_sinais()

    def _montar_interface(self):
        # === TABELA CONTATOS ===
        self.tabela_contatos = QTableWidget(0, 3)
        self.tabela_contatos.setHorizontalHeaderLabels(["Nome", "Número", "Status"])
        self.tabela_contatos.setMinimumWidth(300)

        btn_excel = QPushButton("Importar Excel")
        btn_excel.setObjectName("btn_importar_excel")

        btn_csv = QPushButton("Importar Números")
        btn_csv.setObjectName("btn_importar_csv")

        btn_manual = QPushButton("Adicionar Manualmente")
        btn_manual.setObjectName("btn_adicionar_manual")

        btn_limpar = QPushButton("Limpar")
        btn_limpar.setObjectName("btn_limpar_contatos")

        btn_apagar = QPushButton("Apagar")
        btn_apagar.setObjectName("btn_apagar_linha")

        botoes_layout = QHBoxLayout()
        for btn in [btn_excel, btn_csv, btn_manual, btn_limpar, btn_apagar]:
            botoes_layout.addWidget(btn)

        contatos_layout = QVBoxLayout()
        contatos_layout.addLayout(botoes_layout)
        contatos_layout.addWidget(self.tabela_contatos)

        # === MENSAGEM ===
        self.lbl_mensagem = QLabel("Mensagem (enviada solta)")
        self.txt_mensagem = QTextEdit()
        self.txt_mensagem.setPlaceholderText("Digite sua mensagem aqui (ex.: 'Oi {nome}')...")

        btn_nome_msg = QPushButton("Nome")
        btn_nome_msg.setObjectName("btn_inserir_nome_msg")

        btn_negrito = QPushButton("Negrito")
        btn_negrito.setObjectName("btn_negrito")

        btn_italico = QPushButton("Itálico")
        btn_italico.setObjectName("btn_italico")

        btn_emoji = QPushButton("Escolher Emoji")
        btn_emoji.setObjectName("btn_emoji")

        format_layout = QHBoxLayout()
        for btn in [btn_nome_msg, btn_negrito, btn_italico, btn_emoji]:
            format_layout.addWidget(btn)

        self.tabela_anexos = QTableWidget(0, 3)
        self.tabela_anexos.setHorizontalHeaderLabels(["Nome Completo", "Tipo", "Legenda"])
        self.tabela_anexos.setMinimumHeight(100)

        btn_nome_legenda = QPushButton("Nome")
        btn_nome_legenda.setObjectName("btn_inserir_nome_legenda")

        btn_anexar = QPushButton("Anexar Arquivos/Imagens")
        btn_anexar.setObjectName("btn_anexar")

        btn_limpar_anexos = QPushButton("Limpar")
        btn_limpar_anexos.setObjectName("btn_limpar_anexos")

        anexos_layout = QHBoxLayout()
        for btn in [btn_nome_legenda, btn_anexar, btn_limpar_anexos]:
            anexos_layout.addWidget(btn)

        btn_config = QPushButton("Configurações")
        btn_config.setObjectName("btn_config")

        self.btn_enviar = QPushButton("Enviar")
        self.btn_enviar.setObjectName("btn_enviar")

        self.btn_pausar = QPushButton("Pausar")
        self.btn_pausar.setObjectName("btn_pausar")
        self.btn_pausar.setEnabled(False)

        self.btn_parar = QPushButton("Parar")
        self.btn_parar.setObjectName("btn_parar")
        self.btn_parar.setEnabled(False)

        envio_layout = QHBoxLayout()
        for btn in [btn_config, self.btn_enviar, self.btn_pausar, self.btn_parar]:
            envio_layout.addWidget(btn)

        msg_layout = QVBoxLayout()
        msg_layout.addWidget(self.lbl_mensagem)
        msg_layout.addWidget(self.txt_mensagem)
        msg_layout.addLayout(format_layout)
        msg_layout.addWidget(self.tabela_anexos)
        msg_layout.addLayout(anexos_layout)
        msg_layout.addLayout(envio_layout)

        # === REGISTROS ===
        self.lbl_registros = QLabel("Registros de Envios")
        self.tabela_registros = QTableWidget(0, 2)
        self.tabela_registros.setHorizontalHeaderLabels(["Número", "Status"])
        btn_limpar_reg = QPushButton("Limpar Registros")
        btn_limpar_reg.setObjectName("btn_limpar_registros")

        reg_layout = QVBoxLayout()
        reg_layout.addWidget(self.lbl_registros)
        reg_layout.addWidget(self.tabela_registros)
        reg_layout.addWidget(btn_limpar_reg)

        # === LAYOUT PRINCIPAL ===
        main_layout = QHBoxLayout()
        main_layout.addLayout(contatos_layout)
        main_layout.addLayout(msg_layout)
        main_layout.addLayout(reg_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def _conectar_sinais(self):
        pass  # Já conectado no controller