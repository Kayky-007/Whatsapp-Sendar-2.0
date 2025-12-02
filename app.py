# app.py
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox,
    QTextEdit, QLineEdit, QDialog, QFormLayout, QLabel, QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import webbrowser
import threading

# ============================= ESTILOS =============================
DARK_STYLESHEET = """
QMainWindow, QWidget { background-color: #1e1e1e; color: #ffffff; }
QPushButton { background-color: #2d2d2d; color: #ffffff; border: 1px solid #444; padding: 6px; font-weight: bold; border-radius: 6px; }
QPushButton:hover { background-color: #3d3d3d; }
QPushButton:disabled { background-color: #252525; color: #666; }
QTableWidget { background-color: #252525; gridline-color: #333; color: #ddd; }
QHeaderView::section { background-color: #2d2d2d; color: #fff; padding: 4px; border: 1px solid #444; }
QTextEdit, QLineEdit { background-color: #2d2d2d; color: #fff; border: 1px solid #444; }
QLabel { color: #ccc; }
QProgressBar { border: 1px solid #444; text-align: center; background: #252525; }
QProgressBar::chunk { background-color: #0a0; }
"""

LIGHT_STYLESHEET = """
QMainWindow, QWidget { background-color: #f5f5f5; color: #000000; }
QPushButton { background-color: #e0e0e0; color: #000; border: 1px solid #ccc; padding: 6px; font-weight: bold; border-radius: 6px; }
QPushButton:hover { background-color: #d0d0d0; }
QPushButton:disabled { background-color: #f0f0f0; color: #888; }
QTableWidget { background-color: #ffffff; gridline-color: #ddd; color: #000; }
QHeaderView::section { background-color: #f0f0f0; color: #000; padding: 4px; border: 1px solid #ccc; }
QTextEdit, QLineEdit { background-color: #ffffff; color: #000; border: 1px solid #ccc; }
QLabel { color: #333; }
QProgressBar { border: 1px solid #ccc; text-align: center; background: #f0f0f0; }
QProgressBar::chunk { background-color: #0a0; }
"""

# ============================= DIALOGS =============================
class ConfigDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configurações")
        self.setGeometry(200, 200, 300, 150)
        layout = QFormLayout()
        self.delay_input = QLineEdit("2")
        layout.addRow("Delay (segundos):", self.delay_input)
        ok = QPushButton("OK")
        ok.clicked.connect(self.accept)
        layout.addWidget(ok)
        self.setLayout(layout)
    def get_delay(self):
        try: return float(self.delay_input.text())
        except: return 2

class AddContactDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Adicionar Contato")
        self.setGeometry(200, 200, 300, 200)
        layout = QFormLayout()
        self.nome = QLineEdit()
        self.numero = QLineEdit()
        self.numero.setPlaceholderText("+55DDDnúmero")
        layout.addRow("Nome:", self.nome)
        layout.addRow("Número:", self.numero)
        ok = QPushButton("Adicionar")
        ok.clicked.connect(self.accept)
        layout.addWidget(ok)
        self.setLayout(layout)
    def get_contact(self):
        return self.nome.text(), self.numero.text()

class ImportExcelDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importar Excel")
        self.setGeometry(200, 200, 400, 200)
        layout = QFormLayout()
        self.file = QLineEdit()
        self.file.setReadOnly(True)
        btn_file = QPushButton("Selecionar")
        btn_file.clicked.connect(self.selecionar)
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file)
        file_layout.addWidget(btn_file)
        layout.addRow("Arquivo:", file_layout)
        self.col_nome = QLineEdit()
        self.col_num = QLineEdit()
        layout.addRow("Coluna Nome:", self.col_nome)
        layout.addRow("Coluna Número:", self.col_num)
        ok = QPushButton("Importar")
        ok.clicked.connect(self.accept)
        layout.addWidget(ok)
        self.setLayout(layout)
    def selecionar(self):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "Excel (*.xlsx)")
        if path: self.file.setText(path)
    def get_excel_info(self):
        return self.file.text(), self.col_nome.text(), self.col_num.text()

# ============================= THREAD =============================
class SenderThread(QThread):
    update_status = Signal(int, str)
    add_log = Signal(str, str)
    finished = Signal()
    update_pause = Signal(bool)
    progress = Signal(int, int)  # atual, total

    def __init__(self, driver, msg, anexos, tabela, tabela_anexos, delay):
        super().__init__()
        self.driver = driver
        self.msg = msg
        self.anexos = anexos
        self.tabela = tabela
        self.tabela_anexos = tabela_anexos
        self.delay = delay
        self.pausado = False
        self.parar = False
        self.lock = threading.Lock()

    def run(self):
        total = self.tabela.rowCount()
        for row in range(total):
            with self.lock:
                if self.parar: break

            while True:
                with self.lock:
                    if not self.pausado or self.parar: break
                time.sleep(0.1)

            numero = self.tabela.item(row, 1).text()
            nome = self.tabela.item(row, 0).text() if self.tabela.item(row, 0) else ""
            status = ""

            if self.msg:
                msg_p = self.msg.replace("{nome}", nome)
                s = self.enviar(numero, mensagem=msg_p)
                status += s

            if self.anexos:
                for i, arq in enumerate(self.anexos):
                    with self.lock:
                        if self.parar: break
                    while True:
                        with self.lock:
                            if not self.pausado or self.parar: break
                        time.sleep(0.1)
                    legenda = ""
                    item = self.tabela_anexos.item(i, 2)
                    if item: legenda = item.text().replace("{nome}", nome)
                    s = self.enviar(numero, arquivo=arq, legenda=legenda)
                    status += (", " if status else "") + s

            self.update_status.emit(row, status.strip())
            self.add_log.emit(numero, status.strip())
            self.progress.emit(row + 1, total)  # ATUALIZA PROGRESSO

        self.finished.emit()

    def enviar(self, numero, mensagem=None, arquivo=None, legenda=None):
        try:
            numero = numero.replace(" ", "").replace("-", "")
            wait = WebDriverWait(self.driver, 15)

            if mensagem and not arquivo:
                url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem.replace(chr(10), '%0A')}"
                self.driver.get(url)
                wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)
                xpath = '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/div/span/button/div/div/div[1]/span'
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                btn.click()
                print(f"[OK] Texto: {numero}")
                time.sleep(self.delay)
                return "Texto OK"

            if arquivo:
                url = f"https://web.whatsapp.com/send?phone={numero}"
                if legenda:
                    url += f"&text={legenda.replace(chr(10), '%0A')}"
                self.driver.get(url)
                wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)

                xpath_plus = '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[1]/div/span/button/div/div/div[1]/span'
                btn_plus = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_plus)))
                btn_plus.click()
                time.sleep(1)

                ext = os.path.splitext(arquivo)[1].lower()
                input_xpath = (
                    '//*[@id="app"]/div/div/span[6]/div/ul/div/div/div[2]/li/div/input'
                    if ext in {'.jpg','.jpeg','.png','.gif','.mp4','.mov'}
                    else '//*[@id="app"]/div/div/span[6]/div/ul/div/div/div[1]/li/div/input'
                )
                campo = wait.until(EC.presence_of_element_located((By.XPATH, input_xpath)))
                campo.send_keys(arquivo)
                time.sleep(3)

                xpath_send_anexo = '//*[@id="app"]/div/div/div[3]/div/div[3]/div[2]/div/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span'
                btn_send_anexo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_send_anexo)))
                btn_send_anexo.click()
                time.sleep(self.delay)
                print(f"[OK] Anexo enviado: {numero}")
                return "Arquivo OK"

        except Exception as e:
            print(f"[ERRO] {numero}: {e}")
            return f"Erro"

    def pausar(self):
        with self.lock:
            self.pausado = not self.pausado
            self.update_pause.emit(self.pausado)

    def parar(self):
        with self.lock:
            self.parar = True
            self.pausado = False
            self.update_pause.emit(False)

# ============================= MAIN WINDOW =============================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhatsApp Sender - KAYKY EDITION")
        self.setGeometry(100, 100, 1400, 700)
        self.setStyleSheet(DARK_STYLESHEET)
        self.tema_escuro = True

        # === PROGRESSO ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        # === TABELA CONTATOS ===
        self.tabela = QTableWidget(0, 3)
        self.tabela.setHorizontalHeaderLabels(["Nome", "Número", "Status"])
        self.tabela.setMinimumWidth(350)

        btns = QHBoxLayout()
        btns.addWidget(self.btn("Importar Excel", self.importar_excel, "#0a7641"))
        btns.addWidget(self.btn("Importar CSV", self.importar_txt, "#182b5f"))
        btns.addWidget(self.btn("Adicionar Manual", self.add_manual, "#685454"))
        btns.addWidget(self.btn("Limpar", self.limpar))
        btns.addWidget(self.btn("Apagar", self.apagar_linha))

        left = QVBoxLayout()
        left.addLayout(btns)
        left.addWidget(self.tabela)

        # === MENSAGEM ===
        self.txt_msg = QTextEdit()
        self.txt_msg.setPlaceholderText("Oi {nome}, tudo bem?")
        btns_msg = QHBoxLayout()
        btns_msg.addWidget(self.btn("Nome", self.ins_nome))
        btns_msg.addWidget(self.btn("Negrito", self.negrito))
        btns_msg.addWidget(self.btn("Itálico", self.italico))
        btns_msg.addWidget(self.btn("Emoji", lambda: webbrowser.open("https://emojipedia.org")))

        self.tabela_anexos = QTableWidget(0, 3)
        self.tabela_anexos.setHorizontalHeaderLabels(["Arquivo", "Tipo", "Legenda"])
        btns_anex = QHBoxLayout()
        btns_anex.addWidget(self.btn("Nome", self.ins_nome_anex))
        btns_anex.addWidget(self.btn("Anexar", self.anexar))
        btns_anex.addWidget(self.btn("Limpar", self.limpar_anexos))

        btns_ctrl = QHBoxLayout()
        btns_ctrl.addWidget(self.btn("Config", self.config))
        self.btn_enviar = self.btn("ENVIAR", self.enviar, "green")
        btns_ctrl.addWidget(self.btn_enviar)
        self.btn_pausar = self.btn("Pausar", self.pausar)
        self.btn_parar = self.btn("Parar", self.parar)
        btns_ctrl.addWidget(self.btn_pausar)
        btns_ctrl.addWidget(self.btn_parar)
        self.btn_pausar.setEnabled(False)
        self.btn_parar.setEnabled(False)
        self.btn_tema = self.btn("Tema Claro", self.trocar_tema, "#333")
        btns_ctrl.addWidget(self.btn_tema)

        center = QVBoxLayout()
        center.addWidget(QLabel("Mensagem"))
        center.addWidget(self.txt_msg)
        center.addLayout(btns_msg)
        center.addWidget(self.tabela_anexos)
        center.addLayout(btns_anex)
        center.addLayout(btns_ctrl)
        center.addWidget(QLabel("Progresso:"))
        center.addWidget(self.progress_bar)

        # === REGISTROS ===
        self.tabela_log = QTableWidget(0, 2)
        self.tabela_log.setHorizontalHeaderLabels(["Número", "Status"])
        right = QVBoxLayout()
        right.addWidget(QLabel("Registros"))
        right.addWidget(self.tabela_log)
        right.addWidget(self.btn("Limpar Logs", self.limpar_logs))

        # === MAIN LAYOUT ===
        main = QHBoxLayout()
        main.addLayout(left, 1)
        main.addLayout(center, 2)
        main.addLayout(right, 1)

        container = QWidget()
        container.setLayout(main)
        self.setCentralWidget(container)

        # === VARIAVEIS ===
        self.delay = 2
        self.anexos = []
        self.driver = None
        self.thread = None
        self.lock = threading.Lock()

    def btn(self, text, func, color=None):
        b = QPushButton(text)
        b.clicked.connect(func)
        if color: b.setStyleSheet(f"background-color: {color}; color: white; font-weight: bold; border-radius: 6px;")
        return b

    # === TEMA ===
    def trocar_tema(self):
        self.tema_escuro = not self.tema_escuro
        self.setStyleSheet(DARK_STYLESHEET if self.tema_escuro else LIGHT_STYLESHEET)
        self.btn_tema.setText("Tema Claro" if self.tema_escuro else "Tema Escuro")

    # === PROGRESSO ===
    def atualizar_progresso(self, atual, total):
        if total > 0:
            percent = int((atual / total) * 100)
            self.progress_bar.setValue(percent)
            self.progress_bar.setFormat(f"Enviando... {atual}/{total} ({percent}%)")

    # === CONTATOS ===
    def add_manual(self):
        d = AddContactDialog()
        if d.exec():
            nome, num = d.get_contact()
            if num.startswith("+55"):
                self.add_row(nome, num)
            else:
                QMessageBox.warning(self, "Erro", "Número deve começar com +55")

    def add_row(self, nome, num, status="Aguardando"):
        r = self.tabela.rowCount()
        self.tabela.insertRow(r)
        self.tabela.setItem(r, 0, QTableWidgetItem(nome))
        self.tabela.setItem(r, 1, QTableWidgetItem(num))
        self.tabela.setItem(r, 2, QTableWidgetItem(status))

    def limpar(self): self.tabela.setRowCount(0)
    def apagar_linha(self):
        r = self.tabela.currentRow()
        if r >= 0: self.tabela.removeRow(r)

    def importar_excel(self):
        d = ImportExcelDialog()
        if d.exec():
            path, col_n, col_num = d.get_excel_info()
            if not path or not col_num:
                return
            try:
                df = pd.read_excel(path)
                if col_num not in df.columns:
                    raise ValueError("Coluna de número não encontrada")

                for _, row in df.iterrows():
                    num_raw = row[col_num]
                    nome = str(row[col_n]) if col_n and pd.notna(row[col_n]) else ""

                    if pd.notna(num_raw):
                        num = str(int(float(num_raw)))
                    else:
                        num = ""

                    if num:
                        self.add_row(nome, num)
                        
                print(f"[OK] {len(df)} contatos importados do Excel")
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))

    def importar_txt(self):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "TXT/CSV (*.txt *.csv)")
        if not path: return
        sep = ";" if path.endswith(".txt") else ","
        df = pd.read_csv(path, sep=sep, header=None)
        for _, row in df.iterrows():
            nome = str(row[0]) if df.shape[1] > 1 else ""
            num = str(row[1] if df.shape[1] > 1 else row[0])
            if pd.notna(num): self.add_row(nome, num)

    # === FORMATAÇÃO ===
    def ins_nome(self):
        self.txt_msg.textCursor().insertText("{nome}")
    def ins_nome_anex(self):
        r = self.tabela_anexos.currentRow()
        if r >= 0:
            item = self.tabela_anexos.item(r, 2) or QTableWidgetItem("")
            item.setText(item.text() + "{nome}")
            self.tabela_anexos.setItem(r, 2, item)

    def negrito(self):
        c = self.txt_msg.textCursor()
        if c.hasSelection():
            c.insertText(f"*{c.selectedText()}*")
        else:
            self.txt_msg.insertPlainText("*texto*")

    def italico(self):
        c = self.txt_msg.textCursor()
        if c.hasSelection():
            c.insertText(f"_{c.selectedText()}_")
        else:
            self.txt_msg.insertPlainText("_texto_")

    # === ANEXOS ===
    def anexar(self):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "Todos (*.*)")
        if path:
            r = self.tabela_anexos.rowCount()
            self.tabela_anexos.insertRow(r)
            self.tabela_anexos.setItem(r, 0, QTableWidgetItem(path))
            self.tabela_anexos.setItem(r, 1, QTableWidgetItem("Arquivo"))
            self.tabela_anexos.setItem(r, 2, QTableWidgetItem(""))
            self.anexos.append(path)

    def limpar_anexos(self):
        self.tabela_anexos.setRowCount(0)
        self.anexos.clear()

    # === CONTROLE ===
    def config(self):
        d = ConfigDialog()
        if d.exec(): self.delay = d.get_delay()

    def iniciar_driver(self):
        if hasattr(self, 'driver') and self.driver:
            return True
        path = './chromedriver.exe'
        if not os.path.exists(path):
            QMessageBox.critical(self, "Erro", "chromedriver.exe não encontrado!")
            return False
        service = Service(path)
        opts = Options()
        opts.add_argument('--disable-gcm')
        opts.add_argument('--log-level=3')
        opts.add_argument('--no-sandbox')
        try:
            self.driver = webdriver.Chrome(service=service, options=opts)
            self.driver.get("https://web.whatsapp.com")
            print("[INFO] Escaneie o QR Code...")
            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "side")))
            print("[OK] WhatsApp conectado!")
            return True
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao iniciar: {e}")
            return False

    def enviar(self):
        if self.tabela.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Adicione contatos")
            return
        if not self.txt_msg.toPlainText() and not self.anexos:
            QMessageBox.warning(self, "Aviso", "Digite mensagem ou anexe")
            return

        if not self.iniciar_driver():
            return

        self.btn_enviar.setEnabled(False)
        self.btn_pausar.setEnabled(True)
        self.btn_parar.setEnabled(True)
        self.progress_bar.setValue(0)

        with self.lock:
            self.thread = SenderThread(
                self.driver,
                self.txt_msg.toPlainText(),
                self.anexos,
                self.tabela,
                self.tabela_anexos,
                self.delay
            )
            self.thread.update_status.connect(lambda r, s: self.tabela.setItem(r, 2, QTableWidgetItem(s)))
            self.thread.add_log.connect(lambda n, s: self.add_log(n, s))
            self.thread.finished.connect(self.finalizado)
            self.thread.update_pause.connect(lambda p: self.btn_pausar.setText("Retomar" if p else "Pausar"))
            self.thread.progress.connect(self.atualizar_progresso)
            self.thread.start()

    def pausar(self):
        with self.lock:
            if self.thread: self.thread.pausar()

    def parar(self):
        with self.lock:
            if self.thread: self.thread.parar()
            if hasattr(self, 'driver'):
                try: self.driver.quit()
                except: pass
                finally: delattr(self, 'driver')
        self.finalizado()

    def finalizado(self):
        self.btn_enviar.setEnabled(True)
        self.btn_pausar.setEnabled(False)
        self.btn_parar.setEnabled(False)
        self.btn_pausar.setText("Pausar")
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("Concluído!")

    def add_log(self, num, status):
        r = self.tabela_log.rowCount()
        self.tabela_log.insertRow(r)
        self.tabela_log.setItem(r, 0, QTableWidgetItem(num))
        self.tabela_log.setItem(r, 1, QTableWidgetItem(status))

    def limpar_logs(self):
        self.tabela_log.setRowCount(0)

    def closeEvent(self, e):
        if QMessageBox.question(self, "Sair", "Fechar o app?") == QMessageBox.Yes:
            self.parar()
            e.accept()
        else:
            e.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())