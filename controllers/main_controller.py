# controllers/main_controller.py
import pandas as pd
from PySide6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem
from PySide6.QtCore import Qt, QObject
import webbrowser

from views.ferramentas.config_ferramenta import ConfigFerramenta
from views.ferramentas.add_contato_ferramenta import AddContatoFerramenta
from views.ferramentas.importar_excel_ferramenta import ImportarExcelFerramenta
from models.sender_model import SenderThread
from uteis.whatsapp_automatico import WhatsAppAutomatico


class MainController:
    def __init__(self, view):
        self.view = view
        self.delay = 2
        self.anexos = []
        self.driver = None
        self.thread = None
        self.whatsapp = None  # ← Não inicia aqui
        self._conectar_sinais()

    def _conectar_sinais(self):
        v = self.view

        # === CONTATOS ===
        self._btn("btn_importar_excel").clicked.connect(self.importar_excel)
        self._btn("btn_importar_csv").clicked.connect(self.importar_csv)
        self._btn("btn_adicionar_manual").clicked.connect(self.adicionar_manual)
        self._btn("btn_limpar_contatos").clicked.connect(self.limpar_tabela)
        self._btn("btn_apagar_linha").clicked.connect(self.apagar_linha)

        # === FORMATAÇÃO MENSAGEM ===
        self._btn("btn_inserir_nome_msg").clicked.connect(self.inserir_nome_mensagem)
        self._btn("btn_negrito").clicked.connect(self.aplicar_negrito)
        self._btn("btn_italico").clicked.connect(self.aplicar_italico)
        self._btn("btn_emoji").clicked.connect(lambda: webbrowser.open("https://emojipedia.org/"))

        # === ANEXOS ===
        self._btn("btn_inserir_nome_legenda").clicked.connect(self.inserir_nome_legenda)
        self._btn("btn_anexar").clicked.connect(self.anexar_arquivo)
        self._btn("btn_limpar_anexos").clicked.connect(self.limpar_anexos)

        # === CONFIGURAÇÃO E ENVIO ===
        self._btn("btn_config").clicked.connect(self.abrir_config)
        v.btn_enviar.clicked.connect(self.enviar_mensagens)
        v.btn_pausar.clicked.connect(self.pausar_envio)
        v.btn_parar.clicked.connect(self.parar_envio)

        # === REGISTROS ===
        self._btn("btn_limpar_registros").clicked.connect(self.limpar_registros)

    def _btn(self, name):
        btn = self.view.findChild(QObject, name)
        if not btn:
            print(f"[ERRO] Botão não encontrado: {name}")
        return btn

    # ================== CONTATOS ==================
    def importar_excel(self):
        dialog = ImportarExcelFerramenta()
        if dialog.exec():
            caminho, col_nome, col_num = dialog.get_dados()
            if not caminho or not col_num:
                QMessageBox.warning(self.view, "Aviso", "Preencha os campos obrigatórios")
                return
            try:
                df = pd.read_excel(caminho)
                if col_num not in df.columns:
                    raise ValueError(f"Coluna '{col_num}' não encontrada")
                for _, r in df.iterrows():
                    num = str(r[col_num])
                    nome = str(r[col_nome]) if col_nome and col_nome in df.columns and pd.notna(r[col_nome]) else ""
                    if pd.notna(num):
                        self._adicionar_linha(nome, num)
            except Exception as e:
                QMessageBox.critical(self.view, "Erro", str(e))

    def importar_csv(self):
        caminho, _ = QFileDialog.getOpenFileName(self.view, "", "", "CSV/TXT (*.csv *.txt)")
        if caminho:
            sep = ";" if caminho.endswith(".txt") else ","
            df = pd.read_csv(caminho, sep=sep, header=None)
            for _, r in df.iterrows():
                nome = str(r[0]) if df.shape[1] > 1 else ""
                num = str(r[1] if df.shape[1] > 1 else r[0])
                if pd.notna(num):
                    self._adicionar_linha(nome, num)

    def adicionar_manual(self):
        dialog = AddContatoFerramenta()
        if dialog.exec():
            nome, num = dialog.get_contato()
            if num.startswith("+55"):
                self._adicionar_linha(nome, num)
            else:
                QMessageBox.warning(self.view, "Aviso", "Número deve começar com +55")

    def limpar_tabela(self):
        self.view.tabela_contatos.setRowCount(0)

    def apagar_linha(self):
        row = self.view.tabela_contatos.currentRow()
        if row >= 0:
            self.view.tabela_contatos.removeRow(row)

    def _adicionar_linha(self, nome, numero, status="Aguardando"):
        row = self.view.tabela_contatos.rowCount()
        self.view.tabela_contatos.insertRow(row)
        self.view.tabela_contatos.setItem(row, 0, QTableWidgetItem(nome))
        self.view.tabela_contatos.setItem(row, 1, QTableWidgetItem(numero))
        self.view.tabela_contatos.setItem(row, 2, QTableWidgetItem(status))

    # ================== FORMATAÇÃO ==================
    def inserir_nome_mensagem(self):
        cursor = self.view.txt_mensagem.textCursor()
        cursor.insertText("{nome}")

    def inserir_nome_legenda(self):
        row = self.view.tabela_anexos.currentRow()
        if row >= 0:
            item = self.view.tabela_anexos.item(row, 2) or QTableWidgetItem("")
            item.setText(item.text() + "{nome}")
            self.view.tabela_anexos.setItem(row, 2, item)

    def aplicar_negrito(self):
        txt = self.view.txt_mensagem
        if txt.textCursor().hasSelection():
            s = txt.textCursor().selectedText()
            txt.textCursor().insertText(f"*{s}*")
        else:
            txt.insertPlainText("*texto*")

    def aplicar_italico(self):
        txt = self.view.txt_mensagem
        if txt.textCursor().hasSelection():
            s = txt.textCursor().selectedText()
            txt.textCursor().insertText(f"_{s}_")
        else:
            txt.insertPlainText("_texto_")

    # ================== ANEXOS ==================
    def anexar_arquivo(self):
        caminho, _ = QFileDialog.getOpenFileName(self.view, "", "", "Todos os arquivos (*.*)")
        if caminho:
            row = self.view.tabela_anexos.rowCount()
            self.view.tabela_anexos.insertRow(row)
            self.view.tabela_anexos.setItem(row, 0, QTableWidgetItem(caminho))
            self.view.tabela_anexos.setItem(row, 1, QTableWidgetItem("Arquivo"))
            self.view.tabela_anexos.setItem(row, 2, QTableWidgetItem(""))
            self.anexos.append(caminho)

    def limpar_anexos(self):
        self.view.tabela_anexos.setRowCount(0)
        self.anexos.clear()

    # ================== CONFIGURAÇÃO ==================
    def abrir_config(self):
        dialog = ConfigFerramenta()
        if dialog.exec():
            self.delay = dialog.get_delay()

    # ================== ENVIO DE MENSAGENS ==================
def enviar_mensagens(self):
    # === VALIDAÇÕES ===
    if self.view.tabela_contatos.rowCount() == 0:
        QMessageBox.warning(self.view, "Aviso", "Adicione contatos")
        return
    if not self.view.txt_mensagem.toPlainText() and not self.anexos:
        QMessageBox.warning(self.view, "Aviso", "Digite uma mensagem ou anexe algo")
        return

    # === INICIAR WHATSAPP (FORÇADO AQUI) ===
    if not hasattr(self, 'whatsapp') or not self.whatsapp:
        self.whatsapp = WhatsAppAutomatico()

    if not self.whatsapp.driver:
        print("[INFO] Iniciando WhatsApp...")
        if not self.whatsapp.iniciar():
            QMessageBox.critical(self.view, "Erro", "Falha ao conectar ao WhatsApp")
            return
        self.driver = self.whatsapp.driver
        print("[SUCESSO] WhatsApp iniciado com sucesso!")
    else:
        self.driver = self.whatsapp.driver

    # === DESABILITAR BOTÕES ===
    self.view.btn_enviar.setEnabled(False)
    self.view.btn_pausar.setEnabled(True)
    self.view.btn_parar.setEnabled(True)

    # === INICIAR THREAD ===
    self.thread = SenderThread(
        self.driver,
        self.view.txt_mensagem.toPlainText(),
        self.anexos,
        self.view.tabela_contatos,
        self.view.tabela_anexos,
        self.delay
    )
    self.thread.update_status.connect(
        lambda r, s: self.view.tabela_contatos.setItem(r, 2, QTableWidgetItem(s))
    )
    self.thread.add_log.connect(self._adicionar_log)
    self.thread.finished.connect(self._finalizado)
    self.thread.update_pause_button.connect(
        lambda p: self.view.btn_pausar.setText("Retomar" if p else "Pausar")
    )
    self.thread.start()

    def pausar_envio(self):
        if self.thread and self.thread.isRunning():
            self.thread.pausar()

    def parar_envio(self):
        if self.driver:
            self.whatsapp.fechar()
            self.driver = None
        if self.thread:
            self.thread.parar()
            self.thread.wait()
        self._resetar_botoes()

    def _finalizado(self):
        self._resetar_botoes()

    def _resetar_botoes(self):
        self.view.btn_enviar.setEnabled(True)
        self.view.btn_pausar.setEnabled(False)
        self.view.btn_parar.setEnabled(False)
        self.view.btn_pausar.setText("Pausar")
        self.thread = None

    def _adicionar_log(self, num, status):
        row = self.view.tabela_registros.rowCount()
        self.view.tabela_registros.insertRow(row)
        self.view.tabela_registros.setItem(row, 0, QTableWidgetItem(num))
        self.view.tabela_registros.setItem(row, 1, QTableWidgetItem(status))

    def limpar_registros(self):
        self.view.tabela_registros.setRowCount(0)