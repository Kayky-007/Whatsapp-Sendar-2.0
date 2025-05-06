from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox,
    QTextEdit, QLineEdit, QDialog, QFormLayout, QLabel
)
from PySide6.QtCore import Qt, QThread, Signal
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import os
import webbrowser
import threading

class ConfigDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configurações")
        self.setGeometry(200, 200, 300, 150)

        layout = QFormLayout()
        self.delay_input = QLineEdit("2")  # Delay padrão: 2 segundos
        layout.addRow("Delay entre mensagens (segundos):", self.delay_input)
        
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        layout.addWidget(self.ok_button)

        self.setLayout(layout)

    def get_delay(self):
        return float(self.delay_input.text())

class AddContactDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Adicionar Contato Manualmente")
        self.setGeometry(200, 200, 300, 200)

        layout = QFormLayout()
        self.nome_input = QLineEdit()
        self.numero_input = QLineEdit()
        self.numero_input.setPlaceholderText("+55DDDnúmero")
        layout.addRow("Nome:", self.nome_input)
        layout.addRow("Número:", self.numero_input)

        self.ok_button = QPushButton("Adicionar")
        self.ok_button.clicked.connect(self.accept)
        layout.addWidget(self.ok_button)

        self.setLayout(layout)

    def get_contact(self):
        return self.nome_input.text(), self.numero_input.text()

class ImportExcelDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importar Contatos do Excel")
        self.setGeometry(200, 200, 400, 200)

        layout = QFormLayout()
        
        # Campo para selecionar o arquivo Excel
        self.file_path_input = QLineEdit()
        self.file_path_input.setPlaceholderText("Selecione o arquivo .xlsx")
        self.file_path_input.setReadOnly(True)
        self.file_select_button = QPushButton("Selecionar Arquivo")
        self.file_select_button.clicked.connect(self.select_file)
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_input)
        file_layout.addWidget(self.file_select_button)
        layout.addRow("Arquivo Excel:", file_layout)

        # Campo para o nome da coluna de nomes (opcional)
        self.nome_coluna_input = QLineEdit()
        self.nome_coluna_input.setPlaceholderText("Ex.: Nome (opcional)")
        layout.addRow("Coluna com Nomes:", self.nome_coluna_input)

        # Campo para o nome da coluna de números
        self.numero_coluna_input = QLineEdit()
        self.numero_coluna_input.setPlaceholderText("Ex.: Número")
        layout.addRow("Coluna com Números:", self.numero_coluna_input)

        # Botão de confirmação
        self.ok_button = QPushButton("Importar")
        self.ok_button.clicked.connect(self.accept)
        layout.addWidget(self.ok_button)

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.file_path_input.setText(file_path)

    def get_excel_info(self):
        return (
            self.file_path_input.text(),
            self.nome_coluna_input.text(),
            self.numero_coluna_input.text()
        )

class SenderThread(QThread):
    update_status = Signal(int, str)  # Sinal para atualizar o status na tabela de contatos
    add_log = Signal(str, str)  # Sinal para adicionar um registro na tabela de registros
    finished = Signal()  # Sinal para indicar que o envio terminou
    update_pause_button = Signal(bool)  # Sinal para atualizar o texto do botão "Pausar"

    def __init__(self, driver, mensagem, anexos, tabela, tabela_anexos, delay, parent=None):
        super().__init__(parent)
        self.driver = driver
        self.mensagem = mensagem
        self.anexos = anexos
        self.tabela = tabela
        self.tabela_anexos = tabela_anexos
        self.delay = delay
        self.pausado = False
        self.parar = False
        self.lock = threading.Lock()  # Lock para sincronizar acesso às variáveis pausado/parar

    def run(self):
        for row in range(self.tabela.rowCount()):
            with self.lock:
                if self.parar:
                    break

            while True:
                with self.lock:
                    if not self.pausado:
                        break
                    if self.parar:
                        break
                time.sleep(0.1)

            # Obter o número e o nome do contato
            numero = self.tabela.item(row, 1).text()
            nome_item = self.tabela.item(row, 0)
            nome = nome_item.text() if nome_item and nome_item.text() else ""  # String vazia se nome estiver vazio
            status = ""

            # Passo 1: Enviar mensagem solta, se houver
            if self.mensagem:
                mensagem_personalizada = self.mensagem.replace("{nome}", nome)
                status = self.enviar_mensagem_por_numero(numero, mensagem=mensagem_personalizada)

            # Passo 2: Para cada anexo, carregar a legenda (se houver) e enviar o anexo
            if self.anexos:
                for i, arquivo in enumerate(self.anexos):
                    with self.lock:
                        if self.parar:
                            break

                    while True:
                        with self.lock:
                            if not self.pausado:
                                break
                            if self.parar:
                                break
                        time.sleep(0.1)

                    # Obter a legenda e substituir o marcador {nome}
                    legenda_item = self.tabela_anexos.item(i, 2)
                    legenda = legenda_item.text() if legenda_item and legenda_item.text() else None
                    if legenda:
                        legenda_personalizada = legenda.replace("{nome}", nome)
                    else:
                        legenda_personalizada = None

                    # Enviar o anexo com a legenda carregada via URL
                    anexo_status = self.enviar_mensagem_por_numero(numero, arquivo=arquivo, legenda=legenda_personalizada)
                    status = f"{status}, {anexo_status}" if status else anexo_status

            self.update_status.emit(row, status)
            self.add_log.emit(numero, status)

        self.finished.emit()

    def enviar_mensagem_por_numero(self, numero, mensagem=None, arquivo=None, legenda=None):
        try:
            # Formatar número (remover caracteres indesejados, exceto o +)
            numero = numero.replace(" ", "").replace("-", "")
            
            # Se houver mensagem solta, enviar como texto
            if mensagem:
                mensagem_encoded = mensagem.replace("\n", "%0A")  # Codifica quebras de linha
                url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_encoded}"
                self.driver.get(url)
                while len(self.driver.find_elements(By.ID, 'side')) < 1:
                    time.sleep(1)
                time.sleep(2)

                # Clica no botão de enviar mensagem
                botao_enviar_mensagem = self.driver.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span/div/div[2]/div[2]/button/span")
                botao_enviar_mensagem.click()
                time.sleep(self.delay)
                return "Enviado (texto)"

            # Se houver arquivo, carregar a legenda (se fornecida) via URL e enviar o anexo
            if arquivo:
                # Carregar a legenda via URL, se houver
                if legenda:
                    legenda_encoded = legenda.replace("\n", "%0A")  # Codifica quebras de linha
                    url = f"https://web.whatsapp.com/send?phone={numero}&text={legenda_encoded}"
                else:
                    url = f"https://web.whatsapp.com/send?phone={numero}"
                
                self.driver.get(url)
                while len(self.driver.find_elements(By.ID, 'side')) < 1:
                    time.sleep(1)
                time.sleep(2)

                # Clica no botão de anexar
                botao_anexar = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button/span')
                botao_anexar.click()
                time.sleep(1)

                # Faz o upload do arquivo
                campo_upload = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input')
                campo_upload.send_keys(arquivo)
                time.sleep(3)

                # Clica no botão de enviar (o WhatsApp usará o texto carregado via URL como legenda)
                botao_enviar = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span')
                botao_enviar.click()
                time.sleep(self.delay)
                return "Enviado (com anexo)" + (f" com legenda" if legenda else "")

        except Exception as e:
            return f"Erro: {str(e)}"

    def pausar(self):
        with self.lock:
            self.pausado = not self.pausado
            self.update_pause_button.emit(self.pausado)

    def parar(self):
        with self.lock:
            self.parar = True
            self.pausado = False
            self.update_pause_button.emit(False)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhatsApp Sender")
        self.setGeometry(100, 100, 1200, 600)

        # Seção Esquerda: Tabela de contatos
        self.tabela = QTableWidget()
        self.tabela.setRowCount(0)
        self.tabela.setColumnCount(3)
        self.tabela.setHorizontalHeaderLabels(["Nome", "Número", "Status"])
        self.tabela.setMinimumWidth(300)

        self.botao_importar_excel = QPushButton("Importar Excel")
        self.botao_importar_excel.clicked.connect(self.importar_excel)
        self.botao_importar = QPushButton("Importar Números")
        self.botao_importar.clicked.connect(self.importar_contatos)
        self.botao_adicionar_manual = QPushButton("Adicionar Manualmente")
        self.botao_adicionar_manual.clicked.connect(self.adicionar_manual)
        self.botao_limpar = QPushButton("Limpar")
        self.botao_limpar.clicked.connect(self.limpar_tabela)
        self.botao_apagar = QPushButton("Apagar")
        self.botao_apagar.clicked.connect(self.apagar_linha)

        botoes_contatos_layout = QHBoxLayout()
        botoes_contatos_layout.addWidget(self.botao_importar_excel)
        botoes_contatos_layout.addWidget(self.botao_importar)
        botoes_contatos_layout.addWidget(self.botao_adicionar_manual)
        botoes_contatos_layout.addWidget(self.botao_limpar)
        botoes_contatos_layout.addWidget(self.botao_apagar)

        contatos_layout = QVBoxLayout()
        contatos_layout.addLayout(botoes_contatos_layout)
        contatos_layout.addWidget(self.tabela)

        # Seção Central: Mensagem
        self.mensagem_label = QLabel("Mensagem (enviada solta)")
        self.mensagem_texto = QTextEdit()
        self.mensagem_texto.setPlaceholderText("Digite sua mensagem aqui (ex.: 'Oi {nome}')...")
        self.mensagem_texto.setMinimumWidth(400)

        self.botao_nome = QPushButton("Nome")
        self.botao_nome.clicked.connect(self.inserir_marcador_nome)
        self.botao_negrito = QPushButton("Negrito")
        self.botao_negrito.clicked.connect(self.aplicar_negrito)
        self.botao_italico = QPushButton("Itálico")
        self.botao_italico.clicked.connect(self.aplicar_italico)
        self.botao_emoji = QPushButton("Escolher Emoji")
        self.botao_emoji.clicked.connect(self.abrir_site_emoji)

        formatacao_layout = QHBoxLayout()
        formatacao_layout.addWidget(self.botao_nome)
        formatacao_layout.addWidget(self.botao_negrito)
        formatacao_layout.addWidget(self.botao_italico)
        formatacao_layout.addWidget(self.botao_emoji)

        self.tabela_anexos = QTableWidget()
        self.tabela_anexos.setRowCount(0)
        self.tabela_anexos.setColumnCount(3)
        self.tabela_anexos.setHorizontalHeaderLabels(["Nome Completo", "Tipo", "Legenda (para anexos)"])
        self.tabela_anexos.setMinimumHeight(100)

        self.botao_nome_legenda = QPushButton("Nome")
        self.botao_nome_legenda.clicked.connect(self.inserir_marcador_nome_legenda)
        self.botao_anexar = QPushButton("Anexar Arquivos & Imagens")
        self.botao_anexar.clicked.connect(self.anexar_arquivo)
        self.botao_limpar_anexos = QPushButton("Limpar")
        self.botao_limpar_anexos.clicked.connect(self.limpar_anexos)

        anexos_layout = QHBoxLayout()
        anexos_layout.addWidget(self.botao_nome_legenda)
        anexos_layout.addWidget(self.botao_anexar)
        anexos_layout.addWidget(self.botao_limpar_anexos)

        self.botao_config = QPushButton("Configurações")
        self.botao_config.clicked.connect(self.abrir_configuracoes)
        self.botao_enviar = QPushButton("Enviar")
        self.botao_enviar.clicked.connect(self.enviar_mensagens)
        self.botao_pausar = QPushButton("Pausar")
        self.botao_pausar.clicked.connect(self.pausar_envio)
        self.botao_pausar.setEnabled(False)  # Desativado inicialmente
        self.botao_parar = QPushButton("Parar")
        self.botao_parar.clicked.connect(self.parar_envio)
        self.botao_parar.setEnabled(False)  # Desativado inicialmente

        config_enviar_layout = QHBoxLayout()
        config_enviar_layout.addWidget(self.botao_config)
        config_enviar_layout.addWidget(self.botao_enviar)
        config_enviar_layout.addWidget(self.botao_pausar)
        config_enviar_layout.addWidget(self.botao_parar)

        mensagem_layout = QVBoxLayout()
        mensagem_layout.addWidget(self.mensagem_label)
        mensagem_layout.addWidget(self.mensagem_texto)
        mensagem_layout.addLayout(formatacao_layout)
        mensagem_layout.addWidget(self.tabela_anexos)
        mensagem_layout.addLayout(anexos_layout)
        mensagem_layout.addLayout(config_enviar_layout)

        # Seção Direita: Registros de envio
        self.registros_label = QLabel("Registros de Envios")
        self.tabela_registros = QTableWidget()
        self.tabela_registros.setRowCount(0)
        self.tabela_registros.setColumnCount(2)
        self.tabela_registros.setHorizontalHeaderLabels(["Número", "Status"])
        self.tabela_registros.setMinimumWidth(300)

        self.botao_limpar_registros = QPushButton("Limpar Registros")
        self.botao_limpar_registros.clicked.connect(self.limpar_registros)

        registros_layout = QVBoxLayout()
        registros_layout.addWidget(self.registros_label)
        registros_layout.addWidget(self.tabela_registros)
        registros_layout.addWidget(self.botao_limpar_registros)

        # Layout principal (horizontal)
        main_layout = QHBoxLayout()
        main_layout.addLayout(contatos_layout)  # Esquerda
        main_layout.addLayout(mensagem_layout)  # Centro
        main_layout.addLayout(registros_layout)  # Direita

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Variáveis de controle
        self.delay = 2  # Delay padrão em segundos
        self.anexos = []  # Lista para armazenar caminhos dos anexos
        self.sender_thread = None  # Thread de envio
        self.thread_lock = threading.Lock()  # Lock para proteger acesso a self.sender_thread

    def adicionar_linha(self, nome, numero, status="Aguardando"):
        row_position = self.tabela.rowCount()
        self.tabela.insertRow(row_position)
        self.tabela.setItem(row_position, 0, QTableWidgetItem(str(nome)))
        self.tabela.setItem(row_position, 1, QTableWidgetItem(str(numero)))
        self.tabela.setItem(row_position, 2, QTableWidgetItem(status))

    def adicionar_manual(self):
        dialog = AddContactDialog()
        if dialog.exec():
            nome, numero = dialog.get_contact()
            if numero.startswith("+55"):
                self.adicionar_linha(nome, numero)
            else:
                QMessageBox.warning(self, "Aviso", "O número deve começar com +55DDD.")

    def limpar_tabela(self):
        self.tabela.setRowCount(0)

    def apagar_linha(self):
        selected_row = self.tabela.currentRow()
        if selected_row >= 0:
            self.tabela.removeRow(selected_row)
        else:
            QMessageBox.warning(self, "Aviso", "Selecione uma linha para apagar.")

    def importar_excel(self):
        dialog = ImportExcelDialog()
        if dialog.exec():
            file_path, nome_coluna, numero_coluna = dialog.get_excel_info()

            # Validar entradas
            if not file_path:
                QMessageBox.warning(self, "Aviso", "Selecione um arquivo Excel.")
                return
            if not numero_coluna:
                QMessageBox.warning(self, "Aviso", "Especifique o nome da coluna para números.")
                return

            try:
                # Ler o arquivo Excel
                df = pd.read_excel(file_path)

                # Verificar se a coluna de números especificada existe
                if numero_coluna not in df.columns:
                    QMessageBox.critical(self, "Erro", f"A coluna '{numero_coluna}' não foi encontrada no arquivo Excel.")
                    return

                # Preencher a tabela com os dados do Excel
                for _, row in df.iterrows():
                    # Obter o número
                    numero = row[numero_coluna]
                    if pd.notna(numero):
                        # Obter o nome, se a coluna foi especificada
                        if nome_coluna and nome_coluna in df.columns:
                            nome = row[nome_coluna] if pd.notna(row[nome_coluna]) else ""
                        else:
                            nome = ""  # Nome vazio se a coluna de nomes não for especificada
                        self.adicionar_linha(nome, str(numero))

            except Exception as e:
                QMessageBox.critical(self, "Erro ao importar", str(e))

    def importar_contatos(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo", "", "Arquivos CSV/TXT (*.csv *.txt)")
        if file_path:
            try:
                sep = ";" if file_path.endswith(".txt") else ","
                df = pd.read_csv(file_path, sep=sep, header=None)

                if df.shape[1] >= 2:
                    for _, row in df.iterrows():
                        nome = row[0]
                        numero = row[1]
                        if pd.notna(numero):
                            self.adicionar_linha(nome, numero)
                else:
                    for _, row in df.iterrows():
                        numero = row[0]
                        if pd.notna(numero):
                            self.adicionar_linha("", numero)

            except Exception as e:
                QMessageBox.critical(self, "Erro ao importar", str(e))

    def inserir_marcador_nome(self):
        cursor = self.mensagem_texto.textCursor()
        cursor.insertText("{nome}")
        self.mensagem_texto.setTextCursor(cursor)

    def inserir_marcador_nome_legenda(self):
        selected_row = self.tabela_anexos.currentRow()
        if selected_row >= 0:
            legenda_item = self.tabela_anexos.item(selected_row, 2)
            current_text = legenda_item.text() if legenda_item else ""
            new_text = current_text + "{nome}"
            self.tabela_anexos.setItem(selected_row, 2, QTableWidgetItem(new_text))
        else:
            QMessageBox.warning(self, "Aviso", "Selecione uma linha na tabela de anexos para adicionar o marcador {nome}.")

    def aplicar_negrito(self):
        cursor = self.mensagem_texto.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.insertText(f"*{selected_text}*")
        else:
            self.mensagem_texto.insertPlainText("*texto*")

    def aplicar_italico(self):
        cursor = self.mensagem_texto.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.insertText(f"_{selected_text}_")
        else:
            self.mensagem_texto.insertPlainText("_texto_")

    def abrir_site_emoji(self):
        webbrowser.open("https://emojipedia.org/")

    def anexar_arquivo(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo", "", "Todos os arquivos (*.*)")
        if file_path:
            row_position = self.tabela_anexos.rowCount()
            self.tabela_anexos.insertRow(row_position)
            self.tabela_anexos.setItem(row_position, 0, QTableWidgetItem(file_path))
            self.tabela_anexos.setItem(row_position, 1, QTableWidgetItem("Arquivo"))
            self.tabela_anexos.setItem(row_position, 2, QTableWidgetItem(""))
            self.anexos.append(file_path)

    def limpar_anexos(self):
        self.tabela_anexos.setRowCount(0)
        self.anexos.clear()

    def abrir_configuracoes(self):
        dialog = ConfigDialog()
        if dialog.exec():
            self.delay = dialog.get_delay()

    def pausar_envio(self):
        with self.thread_lock:
            if self.sender_thread and self.sender_thread.isRunning():
                self.sender_thread.pausar()
            else:
                print("Pausar: Thread não está em execução ou não existe.")

    def parar_envio(self):
        with self.thread_lock:
            # Fechar o navegador para interromper o envio
            if hasattr(self, 'driver'):
                try:
                    self.driver.quit()
                except Exception as e:
                    print(f"Erro ao fechar o navegador: {e}")
                finally:
                    delattr(self, 'driver')  # Remove a referência ao driver

            # Garantir que a thread seja parada e finalizada
            if self.sender_thread and self.sender_thread.isRunning():
                self.sender_thread.parar()
                self.sender_thread.wait()  # Aguarda a thread terminar

            # Resetar o estado da interface
            self.botao_enviar.setEnabled(True)
            self.botao_pausar.setEnabled(False)
            self.botao_parar.setEnabled(False)
            self.botao_pausar.setText("Pausar")
            self.sender_thread = None

    def atualizar_texto_pausar(self, pausado):
        self.botao_pausar.setText("Retomar" if pausado else "Pausar")

    def enviar_mensagens(self):
        if not hasattr(self, 'driver'):
            # Configuração inicial do Selenium
            chromedriver_path = './chromedriver.exe'
            service = Service(chromedriver_path)
            self.driver = webdriver.Chrome(service=service)
            self.driver.get("https://web.whatsapp.com/")
            QMessageBox.information(self, "Instrução", "Escaneie o QR Code para entrar no WhatsApp Web e clique em OK.")
            while len(self.driver.find_elements(By.ID, 'side')) < 1:
                time.sleep(1)
            time.sleep(2)

        mensagem = self.mensagem_texto.toPlainText()
        if not mensagem and not self.anexos:
            QMessageBox.warning(self, "Aviso", "Digite uma mensagem ou anexe um arquivo para enviar.")
            return

        if self.tabela.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Importe números para enviar mensagens.")
            return

        # Desativar o botão "Enviar" e ativar "Pausar" e "Parar"
        self.botao_enviar.setEnabled(False)
        self.botao_pausar.setEnabled(True)
        self.botao_parar.setEnabled(True)

        # Iniciar a thread de envio
        with self.thread_lock:
            self.sender_thread = SenderThread(
                self.driver, mensagem, self.anexos, self.tabela, self.tabela_anexos, self.delay, self
            )
            self.sender_thread.update_status.connect(self.atualizar_status)
            self.sender_thread.add_log.connect(self.adicionar_registro)
            self.sender_thread.finished.connect(self.envio_finalizado)
            self.sender_thread.update_pause_button.connect(self.atualizar_texto_pausar)
            self.sender_thread.start()

    def atualizar_status(self, row, status):
        self.tabela.setItem(row, 2, QTableWidgetItem(status))

    def adicionar_registro(self, numero, status):
        row_position = self.tabela_registros.rowCount()
        self.tabela_registros.insertRow(row_position)
        self.tabela_registros.setItem(row_position, 0, QTableWidgetItem(numero))
        self.tabela_registros.setItem(row_position, 1, QTableWidgetItem(status))

    def envio_finalizado(self):
        with self.thread_lock:
            self.botao_enviar.setEnabled(True)
            self.botao_pausar.setEnabled(False)
            self.botao_parar.setEnabled(False)
            self.botao_pausar.setText("Pausar")
            self.sender_thread = None

    def limpar_registros(self):
        self.tabela_registros.setRowCount(0)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Confirmação', "Tem certeza que deseja fechar o sistema? Isso fechará o navegador também.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            with self.thread_lock:
                if self.sender_thread and self.sender_thread.isRunning():
                    self.sender_thread.parar()
                    self.sender_thread.wait()  # Aguarda a thread terminar
            if hasattr(self, 'driver'):
                self.driver.quit()
                delattr(self, 'driver')  # Remove a referência ao driver
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())