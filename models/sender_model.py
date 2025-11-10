# models/sender_model.py
from PySide6.QtCore import QThread, Signal
import threading
from uteis.whatsapp_automatico import WhatsAppAutomatico

class SenderThread(QThread):
    update_status = Signal(int, str)
    add_log = Signal(str, str)
    finished = Signal()
    update_pause_button = Signal(bool)

    def __init__(self, driver, mensagem, anexos, tabela_contatos, tabela_anexos, delay):
        super().__init__()
        self.driver = driver
        self.mensagem = mensagem
        self.anexos = anexos
        self.tabela_contatos = tabela_contatos
        self.tabela_anexos = tabela_anexos
        self.delay = delay
        self.pausado = False
        self.parar = False
        self.lock = threading.Lock()
        self.whatsapp = WhatsAppAutomatico()
        self.whatsapp.driver = driver

    def run(self):
        for linha in range(self.tabela_contatos.rowCount()):
            with self.lock:
                if self.parar: break

            while self.pausado and not self.parar:
                threading.Event().wait(0.1)

            numero = self.tabela_contatos.item(linha, 1).text()
            nome = self.tabela_contatos.item(linha, 0).text() if self.tabela_contatos.item(linha, 0) else ""
            status = ""

            if self.mensagem:
                msg = self.mensagem.replace("{nome}", nome)
                status = self.whatsapp.enviar(numero, mensagem=msg, delay=self.delay)

            if self.anexos and not self.parar:
                for i, arquivo in enumerate(self.anexos):
                    legenda_item = self.tabela_anexos.item(i, 2)
                    legenda = legenda_item.text().replace("{nome}", nome) if legenda_item else None
                    anexo_status = self.whatsapp.enviar(numero, arquivo=arquivo, legenda=legenda, delay=self.delay)
                    status = f"{status}, {anexo_status}" if status else anexo_status

            self.update_status.emit(linha, status)
            self.add_log.emit(numero, status)

        self.finished.emit()

    def pausar(self):
        with self.lock:
            self.pausado = not self.pausado
            self.update_pause_button.emit(self.pausado)

    def parar(self):
        with self.lock:
            self.parar = True
            self.pausado = False
            self.update_pause_button.emit(False)