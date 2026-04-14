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
import pyperclip
import pyautogui

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
    finished = Signal(int, int, int)
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
        self._parar = False          
        self.lock = threading.Lock()

    def _browser_vivo(self):
        try:
            _ = self.driver.current_url
            return True
        except:
            return False

    def run(self):
        total = self.tabela.rowCount()
        ok = 0
        falhou = 0
        numero_invalido = 0

        for row in range(total):
            with self.lock:
                if self._parar: break

            # Verifica se o browser ainda está aberto
            if not self._browser_vivo():
                print("[ERRO] Browser fechado — encerrando envio")
                break

            while True:
                with self.lock:
                    if not self.pausado or self._parar: break
                time.sleep(0.1)

            numero = self.tabela.item(row, 1).text()
            nome = self.tabela.item(row, 0).text() if self.tabela.item(row, 0) else ""
            status = ""

            msg_p = self.msg.replace("{nome}", nome) if self.msg else None

            if self.anexos:
                for i, arq in enumerate(self.anexos):
                    with self.lock:
                        if self._parar: break

                    if not self._browser_vivo():
                        print("[ERRO] Browser fechado — encerrando envio")
                        self._parar = True
                        break

                    while True:
                        with self.lock:
                            if not self.pausado or self._parar: break
                        time.sleep(0.1)

                    legenda = ""
                    item = self.tabela_anexos.item(i, 2)
                    if item:
                        legenda = item.text().replace("{nome}", nome)

                    msg_envio = msg_p if i == 0 else None
                    s = self.enviar(numero, mensagem=msg_envio, arquivo=arq, legenda=legenda)
                    status += (", " if status else "") + s
                    time.sleep(self.delay)

            else:
                if msg_p:
                    if not self._browser_vivo():
                        print("[ERRO] Browser fechado — encerrando envio")
                        break
                    s = self.enviar(numero, mensagem=msg_p)
                    status += s
                    time.sleep(self.delay)

            # ── Contabiliza resultado ──
            if "OK" in status:
                ok += 1
            elif "invalido" in status.lower() or "invalid" in status.lower():
                numero_invalido += 1
            else:
                falhou += 1

            with self.lock:
                if self._parar:
                    self.update_status.emit(row, status.strip())
                    self.add_log.emit(numero, status.strip())
                    self.progress.emit(row + 1, total)
                    break

            self.update_status.emit(row, status.strip())
            self.add_log.emit(numero, status.strip())
            self.progress.emit(row + 1, total)

        self.finished.emit(ok, falhou, numero_invalido)

    def enviar(self, numero, mensagem=None, arquivo=None, legenda=None):
        try:
            numero = numero.replace(" ", "").replace("-", "").replace("+", "")
            wait = WebDriverWait(self.driver, 15)

            # ── FLUXO 1: só mensagem de texto ──
            if mensagem and not arquivo:
                self.driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)
                self._digitar_e_enviar(mensagem)
                print(f"[OK] Texto enviado: {numero}")
                return "Texto OK"

            # ── FLUXO 2: só arquivo (sem mensagem, com ou sem legenda) ──
            if arquivo and not mensagem:
                self.driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)
                self._enviar_arquivo(numero, arquivo, legenda, recarregar=False)
                return "Arquivo OK"

            # ── FLUXO 3: mensagem + arquivo + legenda ──
            if mensagem and arquivo:
                self.driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)
                self._digitar_e_enviar(mensagem)
                print(f"[OK] Texto enviado: {numero}")
                time.sleep(1.5)
                print(f"[DEBUG] Iniciando envio do arquivo: {arquivo}")
                self._enviar_arquivo(numero, arquivo, legenda, recarregar=False)
                print(f"[DEBUG] _enviar_arquivo concluído")
                return "Texto + Arquivo OK"

        except Exception as e:
            msg_erro = str(e).lower()
            if "invalid" in msg_erro or "phone" in msg_erro or "404" in msg_erro:
                print(f"[WARN] Número inválido: {numero}")
                return "Numero invalido"
            print(f"[ERRO] {numero}: {e}")
            return "Erro"
        
        except Exception as e:
            erro = str(e).lower()
            if any(x in erro for x in ["invalid session", "no such window", 
                                        "disconnected", "chrome not reachable",
                                        "session deleted"]):
                print(f"[ERRO] Browser foi fechado pelo usuário")
                self._parar = True  # para o envio automaticamente
                return "Erro"
            print(f"[ERRO] {numero}: {e}")
            return "Erro"

    def _digitar_e_enviar(self, mensagem):
        """
        Cola a mensagem no campo do WhatsApp via ClipboardEvent
        e clica no botão de enviar.
        """


        wait = WebDriverWait(self.driver, 15)

        # Aguarda o campo de texto estar presente
        campo = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'p._aupe.copyable-text')
        ))

        # Foca e clica no campo
        self.driver.execute_script("arguments[0].focus(); arguments[0].click();", campo)
        time.sleep(0.3)

        # Cola via ClipboardEvent — replica comportamento humano
        self.driver.execute_script("""
            var campo = arguments[0];
            var texto = arguments[1];
            var dataTransfer = new DataTransfer();
            dataTransfer.setData('text/plain', texto);
            var evento = new ClipboardEvent('paste', {
                clipboardData: dataTransfer,
                bubbles: true,
                cancelable: true
            });
            campo.dispatchEvent(evento);
        """, campo, mensagem)
        time.sleep(0.3)

        # Clica no botão de enviar
        btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[@data-icon='wds-ic-send-filled']")
        ))
        btn.click()

    def _enviar_arquivo(self, numero, arquivo, legenda, recarregar=True):


        wait = WebDriverWait(self.driver, 15)

        # Só recarrega se necessário
        if recarregar:
            self.driver.get(f"https://web.whatsapp.com/send?phone={numero}")
            wait.until(EC.presence_of_element_located((By.ID, "main")))
            time.sleep(3)

        # ── Clica no botão + ──
        btn_plus = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '(//*[@id="main"]//footer//button//span)[1]')
        ))
        btn_plus.click()
        time.sleep(1.2)

        # ── Clica no item correto do menu ──
        ext = os.path.splitext(arquivo)[1].lower()
        is_media = ext in {'.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.webp'}

        xpath_menu = "//span[text()='Fotos e vídeos']" if is_media else "//button[.//span[text()='Documento']]"
        item = wait.until(EC.presence_of_element_located((By.XPATH, xpath_menu)))
        self.driver.execute_script("arguments[0].click();", item)
        print(f"[INFO] Clicou em '{'Fotos e vídeos' if is_media else 'Documento'}'")
        time.sleep(2.0)

        # ── Preenche o explorer nativo ──
        arquivo_abs = os.path.abspath(arquivo)
        pyperclip.copy(arquivo_abs)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.15)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.4)
        pyautogui.press('enter')
        print(f"[INFO] Arquivo colado no explorer: {arquivo_abs}")
        time.sleep(2.5)

        # ── Aguarda o botão de enviar aparecer (preview carregado) ──
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='button' and @aria-label='Enviar']")
        ))

        # ── Legenda via ClipboardEvent ──
        if legenda:
            try:
                campo = wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'p._aupe.copyable-text')
                ))
                self.driver.execute_script(
                    "arguments[0].focus(); arguments[0].click();", campo
                )
                time.sleep(0.3)
                self.driver.execute_script("""
                    var campo = arguments[0];
                    var texto = arguments[1];
                    var dataTransfer = new DataTransfer();
                    dataTransfer.setData('text/plain', texto);
                    var evento = new ClipboardEvent('paste', {
                        clipboardData: dataTransfer,
                        bubbles: true,
                        cancelable: true
                    });
                    campo.dispatchEvent(evento);
                """, campo, legenda)
                time.sleep(0.3)
                print(f"[INFO] Legenda inserida")
            except Exception as e:
                print(f"[WARN] Legenda não inserida: {e}")

        # ── Botão enviar anexo ──
        btn_send = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[@role='button' and @aria-label='Enviar']")
        ))
        btn_send.click()
        print(f"[OK] Arquivo enviado: {numero}")


    # def _get_file_input(self, is_media: bool, arquivo: str):
    #     """
    #     Clica no item correto do menu pelo texto visível,
    #     depois preenche o explorer nativo com pyautogui + pyperclip.
    #     Não usa send_keys em input hidden — o WA abre dialog nativo.
    #     """


    #     label = "Fotos e vídeos" if is_media else "Documento"

    #     # ── Passo 1: clicar no item correto do menu pelo texto ──
    #     try:
    #         itens = self.driver.find_elements(By.XPATH, f'//span[contains(text(), "{label}")]')
    #         item_clicavel = None
    #         for item in itens:
    #             if item.is_displayed():
    #                 item_clicavel = item
    #                 break

    #         if item_clicavel is None:
    #             raise Exception(f"Item '{label}' não visível no menu")

    #         # Sobe ao elemento clicável via JS
    #         pai = self.driver.execute_script("""
    #             var el = arguments[0];
    #             for (var i = 0; i < 6; i++) {
    #                 el = el.parentElement;
    #                 if (!el) break;
    #                 if (el.tagName === 'LI' || el.getAttribute('role') === 'button'
    #                     || el.onclick || el.tagName === 'BUTTON') return el;
    #             }
    #             return arguments[0].parentElement;
    #         """, item_clicavel)

    #         self.driver.execute_script("arguments[0].click();", pai)
    #         print(f"[INFO] Clicou em '{label}' no menu")

    #     except Exception as e:
    #         print(f"[ERRO] Não encontrou '{label}' no menu: {e}")
    #         return False

    #     # ── Passo 2: aguardar o explorer nativo abrir ──
    #     time.sleep(2.0)  # o dialog do Windows leva ~1-2s para aparecer

    #     # ── Passo 3: colar o caminho e confirmar ──
    #     try:
    #         arquivo_abs = os.path.abspath(arquivo)
    #         pyperclip.copy(arquivo_abs)

    #         # Cola o caminho no campo de nome do arquivo do explorer
    #         pyautogui.hotkey('ctrl', 'a')
    #         time.sleep(0.1)
    #         pyautogui.hotkey('ctrl', 'v')
    #         time.sleep(0.5)
    #         pyautogui.press('enter')
    #         time.sleep(2.0)  # aguarda o WA processar o arquivo
    #         print(f"[INFO] Arquivo colado no explorer: {arquivo_abs}")
    #         return True

    #     except Exception as e:
    #         print(f"[ERRO] pyautogui falhou ao preencher explorer: {e}")
    #         return False 

    
    def pausar(self):
        with self.lock:
            self.pausado = not self.pausado
            self.update_pause.emit(self.pausado)

    def parar(self):
        with self.lock:
            self._parar = True       # usa a flag renomeada
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
    
    def on_finished(self, ok, falhou, invalido):
        from PySide6.QtWidgets import QMessageBox
        total = ok + falhou + invalido
        msg = (
            f"Envio finalizado!\n\n"
            f"Total processado:  {total}\n"
            f"✅ Enviados com sucesso:  {ok}\n"
            f"❌ Falharam:  {falhou}\n"
            f"⚠️ Números inválidos:  {invalido}"
        )
        QMessageBox.information(self, "Resumo do Envio", msg)
        # Reabilita botões
        self.btn_enviar.setEnabled(True)
        self.btn_parar.setEnabled(False)
        self.btn_pausar.setEnabled(False)

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
            if not path or not col_num: return
            try:
                df = pd.read_excel(path)
                if col_num not in df.columns: raise ValueError("Coluna não encontrada")
                for _, row in df.iterrows():
                    num = str(row[col_num])
                    nome = str(row[col_n]) if col_n and pd.notna(row[col_n]) else ""
                    if pd.notna(num): self.add_row(nome, num)
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
        # Limpa driver anterior se existir (morto ou vivo)
        if hasattr(self, 'driver'):
            try:
                self.driver.quit()
            except:
                pass
            finally:
                delattr(self, 'driver')

        # === LOCALIZA O CHROMEDRIVER CORRETAMENTE DENTRO DO .EXE ===
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            path = os.path.join(base_path, './chromedriver.exe')
        else:
            path = './chromedriver.exe'

        if not os.path.exists(path):
            QMessageBox.critical(self, "Erro Fatal", f"chromedriver.exe não encontrado!\nEsperado em: {path}")
            return False

        service = Service(path)
        opts = Options()
        opts.add_argument('--disable-gcm')
        opts.add_argument('--log-level=3')
        opts.add_argument('--no-sandbox')
        opts.add_argument('--disable-dev-shm-usage')

        try:
            self.driver = webdriver.Chrome(service=service, options=opts)
            self.driver.get("https://web.whatsapp.com")
            print("[INFO] Escaneie o QR Code...")
            
            # Aguarda o QR ser escaneado sem timeout
            while True:
                try:
                    self.driver.find_element(By.ID, "side")
                    break  # achou o #side — WhatsApp conectado
                except:
                    time.sleep(1)  # ainda não conectou, aguarda 1s e tenta de novo
            
            print("[OK] WhatsApp conectado!")
            return True
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao iniciar Chrome:\n{e}")
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
        self._resetar_ui()  # só reseta a UI, sem resumo

    def finalizado(self, ok, falhou, invalido):
        self._resetar_ui()

        # Só limpa o driver se ele já estiver morto (usuário fechou o Chrome)
        if hasattr(self, 'driver'):
            try:
                _ = self.driver.current_url  # testa se ainda está vivo
                # Se chegou aqui, o Chrome está aberto — não faz nada
            except:
                # Chrome foi fechado pelo usuário — limpa o objeto morto
                try:
                    self.driver.quit()
                except:
                    pass
                finally:
                    delattr(self, 'driver')

        if ok + falhou + invalido > 0:
            total = ok + falhou + invalido
            msg = (
                f"Envio finalizado!\n\n"
                f"Total processado:   {total}\n"
                f"✅ Enviados com sucesso:   {ok}\n"
                f"❌ Falharam:   {falhou}\n"
                f"⚠️ Números inválidos:   {invalido}"
            )
            QMessageBox.information(self, "Resumo do Envio", msg)

    def _resetar_ui(self):
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