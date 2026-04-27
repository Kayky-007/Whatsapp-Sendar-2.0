# app.py
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox,
    QTextEdit, QLineEdit, QDialog, QFormLayout, QLabel, QProgressBar, QDialogButtonBox, QSpinBox, QRadioButton, QButtonGroup,
    QFrame, QDateTimeEdit, QTimeEdit, QDialogButtonBox
)
from PySide6.QtWidgets import QHeaderView
from PySide6.QtCore import QThread, QTimer, QDateTime, QTime, Signal, Qt, QObject
from PySide6.QtGui import QIcon
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
import sys
from datetime import datetime
import random
import webbrowser
import threading
import pyperclip
import win32gui
import win32con
import pyautogui

# ============================= ESTILOS =============================
DARK_STYLESHEET = """
/* ── Base ── */
QMainWindow, QDialog, QWidget {
    background-color: #111b21;
    color: #e9edef;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}

/* ── Botões padrão ── */
QPushButton {
    background-color: #202c33;
    color: #e9edef;
    border: 1px solid #2a3942;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #2a3942;
    border-color: #00a884;
    color: #ffffff;
}
QPushButton:pressed {
    background-color: #182229;
    border-color: #00a884;
}
QPushButton:disabled {
    background-color: #182229;
    color: #3b4a54;
    border-color: #1f2c34;
}

/* ── Botão ENVIAR ── */
QPushButton[text="ENVIAR"] {
    background-color: #00a884;
    border: none;
    color: #ffffff;
    font-size: 13px;
    letter-spacing: 0.5px;
}
QPushButton[text="ENVIAR"]:hover {
    background-color: #02b894;
}

/* ── Tabelas ── */
/* ── Tabelas ── */
QTableWidget {
    background-color: #111b21;
    gridline-color: #1f2c34;
    color: #e9edef;
    border: 1px solid #1f2c34;
    border-radius: 4px;
    selection-background-color: #2a3942;
    selection-color: #ffffff;
    alternate-background-color: #182229;
    font-size: 13px;
}
QTableWidget::item {
    padding: 4px 8px;
    border-bottom: 1px solid #1f2c34;
    min-height: 28px;
}
QTableWidget::item:selected {
    background-color: #2a3942;
    color: #ffffff;
}
QHeaderView::section {
    background-color: #182229;
    color: #8696a0;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #2a3942;
    font-weight: 600;
    font-size: 11px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}
QHeaderView::section:first {
    border-top-left-radius: 4px;
}
QHeaderView::section:last {
    border-top-right-radius: 4px;
}

/* ── Text/Line Edit ── */
QTextEdit, QLineEdit {
    background-color: #182229;
    color: #e9edef;
    border: 1px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #00a884;
}
QTextEdit:focus, QLineEdit:focus {
    border-color: #00a884;
    background-color: #1f2c34;
}

/* ── Labels ── */
QLabel {
    color: #8696a0;
    font-size: 12px;
}

/* ── Progress Bar ── */
QProgressBar {
    border: none;
    border-radius: 8px;
    text-align: center;
    background: #182229;
    color: #8696a0;
    font-size: 11px;
    height: 18px;
}
QProgressBar::chunk {
    background-color: #00a884;
    border-radius: 8px;
}

/* ── ScrollBar ── */
QScrollBar:vertical {
    background: #111b21;
    width: 6px;
    border-radius: 3px;
}
QScrollBar::handle:vertical {
    background: #2a3942;
    border-radius: 3px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #00a884; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: #111b21;
    height: 6px;
    border-radius: 3px;
}
QScrollBar::handle:horizontal {
    background: #2a3942;
    border-radius: 3px;
}
QScrollBar::handle:horizontal:hover { background: #00a884; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

/* ── SpinBox / ComboBox / DateTimeEdit ── */
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #182229;
    color: #e9edef;
    border: 1px solid #2a3942;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus {
    border-color: #00a884;
}
QSpinBox::up-button, QSpinBox::down-button {
    background: #202c33;
    border: none;
    width: 16px;
}
QComboBox::drop-down {
    background: #202c33;
    border: none;
}
QComboBox QAbstractItemView {
    background: #182229;
    color: #e9edef;
    selection-background-color: #2a3942;
    border: 1px solid #2a3942;
}

/* ── RadioButton / CheckBox ── */
QRadioButton, QCheckBox {
    color: #8696a0;
    spacing: 8px;
}
QRadioButton:hover, QCheckBox:hover { color: #e9edef; }
QRadioButton::indicator, QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border-radius: 8px;
    border: 2px solid #2a3942;
    background: #182229;
}
QRadioButton::indicator:checked {
    background: #00a884;
    border-color: #02b894;
}
QCheckBox::indicator:checked {
    background: #00a884;
    border-color: #02b894;
    border-radius: 4px;
}

/* ── Dialog ── */
QDialog {
    background-color: #111b21;
    border: 1px solid #2a3942;
    border-radius: 12px;
}

/* ── Tooltip ── */
QToolTip {
    background-color: #182229;
    color: #e9edef;
    border: 1px solid #2a3942;
    border-radius: 6px;
    padding: 6px;
    font-size: 12px;
}

/* ── GroupBox ── */
QGroupBox {
    border: 1px solid #2a3942;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 8px;
    color: #8696a0;
    font-weight: 700;
    font-size: 11px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    color: #8696a0;
}

/* ── DialogButtonBox ── */
QDialogButtonBox QPushButton {
    min-width: 80px;
    padding: 7px 16px;
}
"""

LIGHT_STYLESHEET = """
/* ── Base ── */
QMainWindow, QDialog, QWidget {
    background-color: #e8eaed;
    color: #1f2937;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}

/* ── Botões padrão ── */
QPushButton {
    background-color: #d1d5db;
    color: #1f2937;
    border: 1px solid #b8bec7;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #c4c9d4;
    border-color: #00a884;
    color: #00a884;
}
QPushButton:pressed {
    background-color: #b8bec7;
    border-color: #00a884;
}
QPushButton:disabled {
    background-color: #dde1e7;
    color: #9ca3af;
    border-color: #d1d5db;
}

/* ── Botão ENVIAR ── */
QPushButton[text="ENVIAR"] {
    background-color: #00a884;
    border: none;
    color: #ffffff;
    font-size: 13px;
    letter-spacing: 0.5px;
}
QPushButton[text="ENVIAR"]:hover {
    background-color: #02b894;
}

/* ── Tabelas ── */
QTableWidget {
    background-color: #dde1e7;
    gridline-color: #c4c9d4;
    color: #1f2937;
    border: 1px solid #c4c9d4;
    border-radius: 4px;
    selection-background-color: #b8d8d0;
    selection-color: #1f2937;
    alternate-background-color: #e4e7ec;
    font-size: 13px;
}
QTableWidget::item {
    padding: 4px 8px;
    border-bottom: 1px solid #c4c9d4;
    min-height: 28px;
}
QTableWidget::item:selected {
    background-color: #b8d8d0;
    color: #1f2937;
}
QHeaderView::section {
    background-color: #d1d5db;
    color: #6b7280;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #b8bec7;
    font-weight: 700;
    font-size: 11px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}
QHeaderView::section:first { border-top-left-radius: 4px; }
QHeaderView::section:last { border-top-right-radius: 4px; }

/* ── Text/Line Edit ── */
QTextEdit, QLineEdit {
    background-color: #dde1e7;
    color: #1f2937;
    border: 1px solid #b8bec7;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #00a884;
}
QTextEdit:focus, QLineEdit:focus {
    border-color: #00a884;
    background-color: #d4d8df;
}

/* ── Labels ── */
QLabel {
    color: #4b5563;
    font-size: 12px;
}

/* ── Progress Bar ── */
QProgressBar {
    border: none;
    border-radius: 8px;
    text-align: center;
    background: #c4c9d4;
    color: #4b5563;
    font-size: 11px;
    height: 18px;
}
QProgressBar::chunk {
    background-color: #00a884;
    border-radius: 8px;
}

/* ── ScrollBar ── */
QScrollBar:vertical {
    background: #e8eaed;
    width: 6px;
    border-radius: 3px;
}
QScrollBar::handle:vertical {
    background: #b8bec7;
    border-radius: 3px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #00a884; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: #e8eaed;
    height: 6px;
    border-radius: 3px;
}
QScrollBar::handle:horizontal {
    background: #b8bec7;
    border-radius: 3px;
}
QScrollBar::handle:horizontal:hover { background: #00a884; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

/* ── SpinBox / ComboBox / DateTimeEdit ── */
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #dde1e7;
    color: #1f2937;
    border: 1px solid #b8bec7;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #00a884; }
QSpinBox::up-button, QSpinBox::down-button { background: #c4c9d4; border: none; width: 16px; }
QComboBox::drop-down { background: #c4c9d4; border: none; }
QComboBox QAbstractItemView {
    background: #dde1e7;
    color: #1f2937;
    selection-background-color: #b8d8d0;
    border: 1px solid #b8bec7;
    border-radius: 6px;
}

/* ── RadioButton / CheckBox ── */
QRadioButton, QCheckBox { color: #4b5563; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #1f2937; }
QRadioButton::indicator, QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border-radius: 8px;
    border: 2px solid #b8bec7;
    background: #dde1e7;
}
QRadioButton::indicator:checked { background: #00a884; border-color: #02b894; }
QCheckBox::indicator:checked { background: #00a884; border-color: #02b894; border-radius: 4px; }

/* ── Dialog ── */
QDialog {
    background-color: #e8eaed;
    border: 1px solid #b8bec7;
    border-radius: 12px;
}

/* ── Tooltip ── */
QToolTip {
    background-color: #dde1e7;
    color: #1f2937;
    border: 1px solid #b8bec7;
    border-radius: 6px;
    padding: 6px;
    font-size: 12px;
}

/* ── GroupBox ── */
QGroupBox {
    border: 1px solid #c4c9d4;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 8px;
    color: #6b7280;
    font-weight: 700;
    font-size: 11px;
    background-color: #dde1e7;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; color: #6b7280; }

/* ── DialogButtonBox ── */
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
"""

TEMAS = {
    "Dark Green": DARK_STYLESHEET,
    "Light": LIGHT_STYLESHEET,

"Oceano": """
QMainWindow, QDialog, QWidget {
    background-color: #0a1628;
    color: #b8d4f0;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}
QPushButton {
    background-color: #0e2040;
    color: #b8d4f0;
    border: 1px solid #1a3a6e;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #1a3a6e;
    border-color: #4d9de0;
    color: #e0f0ff;
}
QPushButton:pressed { background-color: #0a1628; }
QPushButton:disabled { background-color: #0a1628; color: #2a4060; border-color: #0e2040; }
QTableWidget {
    background-color: #0a1628;
    gridline-color: #0e2040;
    color: #b8d4f0;
    border: 1px solid #0e2040;
    border-radius: 4px;
    selection-background-color: #1a3a6e;
    alternate-background-color: #0c1e36;
    font-size: 13px;
}
QTableWidget::item { padding: 4px 8px; border-bottom: 1px solid #0e2040; }
QTableWidget::item:selected { background-color: #1a3a6e; color: #e0f0ff; }
QHeaderView::section {
    background-color: #0e2040;
    color: #4d7fa8;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #1a3a6e;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
QTextEdit, QLineEdit {
    background-color: #0e2040;
    color: #b8d4f0;
    border: 1px solid #1a3a6e;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #4d9de0;
}
QTextEdit:focus, QLineEdit:focus { border-color: #4d9de0; background-color: #112448; }
QLabel { color: #4d7fa8; font-size: 12px; }
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #0e2040;
    color: #4d7fa8;
    font-size: 11px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #1a6eb8, stop:0.5 #4d9de0, stop:1 #00c8e0);
    border-radius: 8px;
}
QScrollBar:vertical { background: #0a1628; width: 6px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #1a3a6e; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #4d9de0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #0a1628; height: 6px; border-radius: 3px; }
QScrollBar::handle:horizontal { background: #1a3a6e; border-radius: 3px; }
QScrollBar::handle:horizontal:hover { background: #4d9de0; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #0e2040;
    color: #b8d4f0;
    border: 1px solid #1a3a6e;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #4d9de0; }
QSpinBox::up-button, QSpinBox::down-button { background: #0a1628; border: none; width: 16px; }
QComboBox::drop-down { background: #0a1628; border: none; }
QComboBox QAbstractItemView { background: #0e2040; color: #b8d4f0; selection-background-color: #1a3a6e; border: 1px solid #1a3a6e; }
QRadioButton, QCheckBox { color: #4d7fa8; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #b8d4f0; }
QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #1a3a6e; background: #0e2040; }
QRadioButton::indicator:checked { background: #4d9de0; border-color: #6ab8f0; }
QCheckBox::indicator:checked { background: #4d9de0; border-color: #6ab8f0; border-radius: 4px; }
QDialog { background-color: #0a1628; border: 1px solid #1a3a6e; border-radius: 12px; }
QToolTip { background-color: #0e2040; color: #b8d4f0; border: 1px solid #1a3a6e; border-radius: 6px; padding: 6px; }
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
""",

"Ametista": """
QMainWindow, QDialog, QWidget {
    background-color: #13101a;
    color: #e0d8f0;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}
QPushButton {
    background-color: #1e1828;
    color: #e0d8f0;
    border: 1px solid #2e2440;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #2e2440;
    border-color: #9b6dff;
    color: #c8b0ff;
}
QPushButton:pressed { background-color: #13101a; }
QPushButton:disabled { background-color: #13101a; color: #3a2e50; border-color: #1e1828; }
QTableWidget {
    background-color: #13101a;
    gridline-color: #1e1828;
    color: #e0d8f0;
    border: 1px solid #1e1828;
    border-radius: 4px;
    selection-background-color: #2e2440;
    alternate-background-color: #171220;
    font-size: 13px;
}
QTableWidget::item { padding: 4px 8px; border-bottom: 1px solid #1e1828; }
QTableWidget::item:selected { background-color: #2e2440; color: #ffffff; }
QHeaderView::section {
    background-color: #1e1828;
    color: #6a5490;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #2e2440;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
QTextEdit, QLineEdit {
    background-color: #1e1828;
    color: #e0d8f0;
    border: 1px solid #2e2440;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #9b6dff;
}
QTextEdit:focus, QLineEdit:focus { border-color: #9b6dff; background-color: #221c30; }
QLabel { color: #6a5490; font-size: 12px; }
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #1e1828;
    color: #6a5490;
    font-size: 11px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #6b2fa0, stop:0.5 #9b6dff, stop:1 #c44dff);
    border-radius: 8px;
}
QScrollBar:vertical { background: #13101a; width: 6px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #2e2440; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #9b6dff; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #13101a; height: 6px; border-radius: 3px; }
QScrollBar::handle:horizontal { background: #2e2440; border-radius: 3px; }
QScrollBar::handle:horizontal:hover { background: #9b6dff; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #1e1828;
    color: #e0d8f0;
    border: 1px solid #2e2440;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #9b6dff; }
QSpinBox::up-button, QSpinBox::down-button { background: #13101a; border: none; width: 16px; }
QComboBox::drop-down { background: #13101a; border: none; }
QComboBox QAbstractItemView { background: #1e1828; color: #e0d8f0; selection-background-color: #2e2440; border: 1px solid #2e2440; }
QRadioButton, QCheckBox { color: #6a5490; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #e0d8f0; }
QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #2e2440; background: #1e1828; }
QRadioButton::indicator:checked { background: #9b6dff; border-color: #b890ff; }
QCheckBox::indicator:checked { background: #9b6dff; border-color: #b890ff; border-radius: 4px; }
QDialog { background-color: #13101a; border: 1px solid #2e2440; border-radius: 12px; }
QToolTip { background-color: #1e1828; color: #e0d8f0; border: 1px solid #2e2440; border-radius: 6px; padding: 6px; }
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
""",
    
    "Midnight Blue": """
QMainWindow, QDialog, QWidget {
    background-color: #0d1117;
    color: #cdd9e5;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}
QPushButton {
    background-color: #161b22;
    color: #cdd9e5;
    border: 1px solid #30363d;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #1f2937;
    border-color: #58a6ff;
    color: #58a6ff;
}
QPushButton:pressed { background-color: #0d1117; }
QPushButton:disabled { background-color: #0d1117; color: #484f58; border-color: #21262d; }
QTableWidget {
    background-color: #0d1117;
    gridline-color: #21262d;
    color: #cdd9e5;
    border: 1px solid #21262d;
    border-radius: 4px;
    selection-background-color: #1f2937;
    alternate-background-color: #161b22;
    font-size: 13px;
}
QTableWidget::item { padding: 4px 8px; border-bottom: 1px solid #21262d; }
QTableWidget::item:selected { background-color: #1f2937; color: #ffffff; }
QHeaderView::section {
    background-color: #161b22;
    color: #8b949e;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #30363d;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
}
QTextEdit, QLineEdit {
    background-color: #161b22;
    color: #cdd9e5;
    border: 1px solid #30363d;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #58a6ff;
}
QTextEdit:focus, QLineEdit:focus { border-color: #58a6ff; }
QLabel { color: #8b949e; font-size: 12px; }
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #161b22;
    color: #8b949e;
    font-size: 11px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk { background-color: #58a6ff; border-radius: 8px; }
QScrollBar:vertical { background: #0d1117; width: 6px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #30363d; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #58a6ff; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #0d1117; height: 6px; border-radius: 3px; }
QScrollBar::handle:horizontal { background: #30363d; border-radius: 3px; }
QScrollBar::handle:horizontal:hover { background: #58a6ff; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #161b22;
    color: #cdd9e5;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #58a6ff; }
QSpinBox::up-button, QSpinBox::down-button { background: #0d1117; border: none; width: 16px; }
QComboBox::drop-down { background: #0d1117; border: none; }
QComboBox QAbstractItemView { background: #161b22; color: #cdd9e5; selection-background-color: #1f2937; border: 1px solid #30363d; }
QRadioButton, QCheckBox { color: #8b949e; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #cdd9e5; }
QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #30363d; background: #161b22; }
QRadioButton::indicator:checked { background: #58a6ff; border-color: #79baff; }
QCheckBox::indicator:checked { background: #58a6ff; border-color: #79baff; border-radius: 4px; }
QDialog { background-color: #0d1117; border: 1px solid #30363d; border-radius: 12px; }
QToolTip { background-color: #161b22; color: #cdd9e5; border: 1px solid #30363d; border-radius: 6px; padding: 6px; }
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
""",

    "Rose": """
QMainWindow, QDialog, QWidget {
    background-color: #1c1418;
    color: #f0e6ea;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}
QPushButton {
    background-color: #2a1e24;
    color: #f0e6ea;
    border: 1px solid #3d2832;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #3d2832;
    border-color: #c9748a;
    color: #f0c0cc;
}
QPushButton:pressed { background-color: #1c1418; }
QPushButton:disabled { background-color: #1c1418; color: #5a4048; border-color: #2a1e24; }
QTableWidget {
    background-color: #1c1418;
    gridline-color: #2a1e24;
    color: #f0e6ea;
    border: 1px solid #2a1e24;
    border-radius: 4px;
    selection-background-color: #3d2832;
    alternate-background-color: #221a1e;
    font-size: 13px;
}
QTableWidget::item { padding: 4px 8px; border-bottom: 1px solid #2a1e24; }
QTableWidget::item:selected { background-color: #3d2832; color: #ffffff; }
QHeaderView::section {
    background-color: #221a1e;
    color: #9e7080;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #3d2832;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
}
QTextEdit, QLineEdit {
    background-color: #221a1e;
    color: #f0e6ea;
    border: 1px solid #3d2832;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #c9748a;
}
QTextEdit:focus, QLineEdit:focus { border-color: #c9748a; }
QLabel { color: #9e7080; font-size: 12px; }
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #221a1e;
    color: #9e7080;
    font-size: 11px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk { background-color: #c9748a; border-radius: 8px; }
QScrollBar:vertical { background: #1c1418; width: 6px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #3d2832; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #c9748a; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #1c1418; height: 6px; border-radius: 3px; }
QScrollBar::handle:horizontal { background: #3d2832; border-radius: 3px; }
QScrollBar::handle:horizontal:hover { background: #c9748a; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #221a1e;
    color: #f0e6ea;
    border: 1px solid #3d2832;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #c9748a; }
QSpinBox::up-button, QSpinBox::down-button { background: #1c1418; border: none; width: 16px; }
QComboBox::drop-down { background: #1c1418; border: none; }
QComboBox QAbstractItemView { background: #221a1e; color: #f0e6ea; selection-background-color: #3d2832; border: 1px solid #3d2832; }
QRadioButton, QCheckBox { color: #9e7080; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #f0e6ea; }
QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #3d2832; background: #221a1e; }
QRadioButton::indicator:checked { background: #c9748a; border-color: #e090a0; }
QCheckBox::indicator:checked { background: #c9748a; border-color: #e090a0; border-radius: 4px; }
QDialog { background-color: #1c1418; border: 1px solid #3d2832; border-radius: 12px; }
QToolTip { background-color: #221a1e; color: #f0e6ea; border: 1px solid #3d2832; border-radius: 6px; padding: 6px; }
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
""",

    "Forest": """
QMainWindow, QDialog, QWidget {
    background-color: #141a14;
    color: #d4e6d4;
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}
QPushButton {
    background-color: #1c261c;
    color: #d4e6d4;
    border: 1px solid #2a3a2a;
    padding: 7px 14px;
    font-weight: 600;
    border-radius: 8px;
}
QPushButton:hover {
    background-color: #2a3a2a;
    border-color: #5a9e5a;
    color: #90d090;
}
QPushButton:pressed { background-color: #141a14; }
QPushButton:disabled { background-color: #141a14; color: #3a4e3a; border-color: #1c261c; }
QTableWidget {
    background-color: #141a14;
    gridline-color: #1c261c;
    color: #d4e6d4;
    border: 1px solid #1c261c;
    border-radius: 4px;
    selection-background-color: #2a3a2a;
    alternate-background-color: #181e18;
    font-size: 13px;
}
QTableWidget::item { padding: 4px 8px; border-bottom: 1px solid #1c261c; }
QTableWidget::item:selected { background-color: #2a3a2a; color: #ffffff; }
QHeaderView::section {
    background-color: #1c261c;
    color: #6a8e6a;
    padding: 5px 8px;
    border: none;
    border-bottom: 1px solid #2a3a2a;
    font-weight: 700;
    font-size: 11px;
    text-transform: uppercase;
}
QTextEdit, QLineEdit {
    background-color: #1c261c;
    color: #d4e6d4;
    border: 1px solid #2a3a2a;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #5a9e5a;
}
QTextEdit:focus, QLineEdit:focus { border-color: #5a9e5a; }
QLabel { color: #6a8e6a; font-size: 12px; }
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #1c261c;
    color: #6a8e6a;
    font-size: 11px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk { background-color: #5a9e5a; border-radius: 8px; }
QScrollBar:vertical { background: #141a14; width: 6px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #2a3a2a; border-radius: 3px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #5a9e5a; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #141a14; height: 6px; border-radius: 3px; }
QScrollBar::handle:horizontal { background: #2a3a2a; border-radius: 3px; }
QScrollBar::handle:horizontal:hover { background: #5a9e5a; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QSpinBox, QComboBox, QDateTimeEdit, QTimeEdit {
    background-color: #1c261c;
    color: #d4e6d4;
    border: 1px solid #2a3a2a;
    border-radius: 6px;
    padding: 5px 8px;
}
QSpinBox:focus, QComboBox:focus, QDateTimeEdit:focus { border-color: #5a9e5a; }
QSpinBox::up-button, QSpinBox::down-button { background: #141a14; border: none; width: 16px; }
QComboBox::drop-down { background: #141a14; border: none; }
QComboBox QAbstractItemView { background: #1c261c; color: #d4e6d4; selection-background-color: #2a3a2a; border: 1px solid #2a3a2a; }
QRadioButton, QCheckBox { color: #6a8e6a; spacing: 8px; }
QRadioButton:hover, QCheckBox:hover { color: #d4e6d4; }
QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #2a3a2a; background: #1c261c; }
QRadioButton::indicator:checked { background: #5a9e5a; border-color: #7abe7a; }
QCheckBox::indicator:checked { background: #5a9e5a; border-color: #7abe7a; border-radius: 4px; }
QDialog { background-color: #141a14; border: 1px solid #2a3a2a; border-radius: 12px; }
QToolTip { background-color: #1c261c; color: #d4e6d4; border: 1px solid #2a3a2a; border-radius: 6px; padding: 6px; }
QDialogButtonBox QPushButton { min-width: 80px; padding: 7px 16px; }
"""
}

NOMES_TEMAS = list(TEMAS.keys())


class AgendarDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Agendar Envio")
        self.setMinimumWidth(420)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        # ── Título ──
        titulo = QLabel("🕐 Configurar Agendamento")
        titulo.setStyleSheet("font-size: 15px; font-weight: bold;")
        layout.addWidget(titulo)

        # ── Opções de agendamento ──
        self.grupo = QButtonGroup(self)

        self.rb_horario = QRadioButton("Horário específico")
        self.rb_intervalo = QRadioButton("Repetir a cada X horas")
        self.rb_diario = QRadioButton("Todo dia no mesmo horário")

        self.grupo.addButton(self.rb_horario, 1)
        self.grupo.addButton(self.rb_intervalo, 2)
        self.grupo.addButton(self.rb_diario, 3)
        self.rb_horario.setChecked(True)

        layout.addWidget(self.rb_horario)
        layout.addWidget(self.rb_intervalo)
        layout.addWidget(self.rb_diario)

        # ── Frame horário específico ──
        self.frame_horario = QWidget()
        h_layout = QHBoxLayout(self.frame_horario)
        h_layout.setContentsMargins(20, 0, 0, 0)
        h_layout.addWidget(QLabel("Data e hora:"))
        self.dt_picker = QDateTimeEdit()
        self.dt_picker.setDisplayFormat("dd/MM/yyyy HH:mm")
        self.dt_picker.setDateTime(QDateTime.currentDateTime().addSecs(180))  # começa +3min
        self.dt_picker.setCalendarPopup(True)
        # SEM setMinimumDateTime — a validação fica só no _validar
        h_layout.addWidget(self.dt_picker)
        layout.addWidget(self.frame_horario)

        # ── Frame intervalo ──
        self.frame_intervalo = QWidget()
        i_layout = QVBoxLayout(self.frame_intervalo)
        i_layout.setContentsMargins(20, 4, 0, 4)
        i_layout.setSpacing(8)

        # Linha 1: repetir a cada X horas
        linha1 = QHBoxLayout()
        linha1.addWidget(QLabel("Repetir a cada:"))
        self.spin_horas = QSpinBox()
        self.spin_horas.setMinimum(1)
        self.spin_horas.setMaximum(24)
        self.spin_horas.setValue(2)
        self.spin_horas.setFixedWidth(70)
        linha1.addWidget(self.spin_horas)
        linha1.addWidget(QLabel("horas"))
        linha1.addStretch()
        i_layout.addLayout(linha1)

        # Linha 2: começando às
        linha2 = QHBoxLayout()
        linha2.addWidget(QLabel("Começando às:"))
        self.dt_intervalo_inicio = QDateTimeEdit()
        self.dt_intervalo_inicio.setDisplayFormat("dd/MM/yyyy HH:mm")
        self.dt_intervalo_inicio.setDateTime(QDateTime.currentDateTime().addSecs(300))
        self.dt_intervalo_inicio.setCalendarPopup(True)
        self.dt_intervalo_inicio.setFixedWidth(160)
        linha2.addWidget(self.dt_intervalo_inicio)
        linha2.addStretch()
        i_layout.addLayout(linha2)

        self.frame_intervalo.setVisible(False)
        layout.addWidget(self.frame_intervalo)

        # ── Frame diário ──
        self.frame_diario = QWidget()
        d_layout = QHBoxLayout(self.frame_diario)
        d_layout.setContentsMargins(20, 0, 0, 0)
        d_layout.addWidget(QLabel("Todo dia às:"))
        self.time_diario = QTimeEdit()
        self.time_diario.setDisplayFormat("HH:mm")
        self.time_diario.setTime(QTime(9, 0))
        d_layout.addWidget(self.time_diario)
        self.frame_diario.setVisible(False)
        layout.addWidget(self.frame_diario)

        # ── Mostra frame correto ──
        self.rb_horario.toggled.connect(lambda v: self.frame_horario.setVisible(v))
        self.rb_intervalo.toggled.connect(lambda v: self.frame_intervalo.setVisible(v))
        self.rb_diario.toggled.connect(lambda v: self.frame_diario.setVisible(v))

        # ── Preview do agendamento ──
        self.lbl_preview = QLabel("")
        self.lbl_preview.setStyleSheet("color: #4fc3f7; font-size: 12px;")
        self.lbl_preview.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_preview)

        # Atualiza preview ao mudar valores
        self.dt_picker.dateTimeChanged.connect(self._atualizar_preview)
        self.dt_intervalo_inicio.dateTimeChanged.connect(self._atualizar_preview)
        self.spin_horas.valueChanged.connect(self._atualizar_preview)
        self.time_diario.timeChanged.connect(self._atualizar_preview)
        self.grupo.buttonToggled.connect(self._atualizar_preview)
        self._atualizar_preview()

        # ── Botões ──
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Agendar")
        btns.accepted.connect(self._validar)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _atualizar_preview(self, *args):
        opcao = self.grupo.checkedId()
        if opcao == 1:
            dt = self.dt_picker.dateTime().toPython()
            self.lbl_preview.setText(f"Envio único em: {dt.strftime('%d/%m/%Y às %H:%M')}")
        elif opcao == 2:
            dt = self.dt_intervalo_inicio.dateTime().toPython()
            h = self.spin_horas.value()
            self.lbl_preview.setText(f"Primeiro disparo: {dt.strftime('%H:%M')} — depois a cada {h}h")
        elif opcao == 3:
            t = self.time_diario.time().toPython()
            self.lbl_preview.setText(f"Todo dia às {t.strftime('%H:%M')}")

    def _validar(self):
        from datetime import datetime, timedelta
        agora = datetime.now()
        minimo = agora + timedelta(minutes=3)
        opcao = self.grupo.checkedId()

        if opcao == 1:
            alvo = self.dt_picker.dateTime().toPython()
            if alvo <= agora:
                QMessageBox.warning(self, "Aviso", "O horário deve ser no futuro.")
                return
            if alvo < minimo:
                QMessageBox.warning(self, "Aviso",
                    f"Agende com pelo menos 3 minutos de antecedência.\n"
                    f"Horário mínimo permitido: {minimo.strftime('%H:%M')}")
                return

        elif opcao == 2:
            inicio = self.dt_intervalo_inicio.dateTime().toPython()
            if inicio < minimo:
                QMessageBox.warning(self, "Aviso",
                    f"O primeiro disparo deve ser em pelo menos 3 minutos.\n"
                    f"Horário mínimo permitido: {minimo.strftime('%H:%M')}")
                return

        elif opcao == 3:
            hora_alvo = self.time_diario.time().toPython()
            agora_dt = datetime.now()
            alvo = agora_dt.replace(
                hour=hora_alvo.hour,
                minute=hora_alvo.minute,
                second=0, microsecond=0
            )
            if alvo <= agora_dt:
                alvo_amanha = alvo + timedelta(days=1)
                resposta = QMessageBox.question(self, "Horário já passou",
                    f"O horário {hora_alvo.strftime('%H:%M')} já passou hoje.\n"
                    f"Deseja agendar para amanhã às {hora_alvo.strftime('%H:%M')}?",
                    QMessageBox.Yes | QMessageBox.No)
                if resposta != QMessageBox.Yes:
                    return

        self.accept()

    def get_config(self):
        opcao = self.grupo.checkedId()
        if opcao == 1:
            return {
                "tipo": "unico",
                "datetime": self.dt_picker.dateTime().toPython()
            }
        elif opcao == 2:
            return {
                "tipo": "intervalo",
                "inicio": self.dt_intervalo_inicio.dateTime().toPython(),
                "horas": self.spin_horas.value()
            }
        elif opcao == 3:
            return {
                "tipo": "diario",
                "hora": self.time_diario.time().toPython()
            }



class BarraAgendamento(QWidget):
    cancelar_signal = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVisible(False)
        self.setStyleSheet("""
            QWidget {
                background-color: #1a3a5c;
                border: 1px solid #4fc3f7;
                border-radius: 6px;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)

        # Ícone pulsante
        self.lbl_icone = QLabel("🕐")
        self.lbl_icone.setStyleSheet("font-size: 18px; border: none;")
        layout.addWidget(self.lbl_icone)

        # Texto do agendamento
        self.lbl_texto = QLabel("")
        self.lbl_texto.setStyleSheet("color: #4fc3f7; font-size: 12px; font-weight: bold; border: none;")
        layout.addWidget(self.lbl_texto)

        layout.addStretch()

        # Countdown até o próximo disparo
        self.lbl_countdown = QLabel("")
        self.lbl_countdown.setStyleSheet("color: #ffffff; font-size: 12px; border: none;")
        layout.addWidget(self.lbl_countdown)

        # Botão cancelar
        self.btn_cancelar = QPushButton("✕ Cancelar agendamento")
        self.btn_cancelar.setStyleSheet("""
            QPushButton {
                background-color: #c62828;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 4px 10px;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #e53935; }
        """)
        self.btn_cancelar.clicked.connect(self.cancelar_signal)
        layout.addWidget(self.btn_cancelar)

        # Timer para piscar o ícone e atualizar countdown
        self._piscar = False
        self._proximo_disparo = None
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(1000)

    def ativar(self, texto: str, proximo_disparo=None):
        self.lbl_texto.setText(texto)
        self._proximo_disparo = proximo_disparo
        self.setVisible(True)

    def desativar(self):
        self.setVisible(False)
        self._proximo_disparo = None

    def atualizar_proximo(self, proximo_disparo):
        self._proximo_disparo = proximo_disparo

    def _tick(self):
        # Pisca o ícone
        self._piscar = not self._piscar
        self.lbl_icone.setText("🕐" if self._piscar else "⏰")

        # Countdown até próximo disparo
        if self._proximo_disparo:
            from datetime import datetime
            agora = datetime.now()
            diff = (self._proximo_disparo - agora).total_seconds()
            if diff > 0:
                h = int(diff // 3600)
                m = int((diff % 3600) // 60)
                s = int(diff % 60)
                if h > 0:
                    self.lbl_countdown.setText(f"Próximo disparo em: {h:02d}:{m:02d}:{s:02d}")
                else:
                    self.lbl_countdown.setText(f"Próximo disparo em: {m:02d}:{s:02d}")


class AgendadorThread(QThread):
    disparar = Signal()  # emite quando chega a hora de enviar

    def __init__(self, config: dict):
        super().__init__()
        self.config = config
        self._parar = False

    def parar(self):
        self._parar = True
        self.requestInterruption()  # acorda a thread imediatamente

    def run(self):
        tipo = self.config["tipo"]
        ultimo_disparo = None

        while not self._parar and not self.isInterruptionRequested():
            agora = datetime.now()

            if tipo == "unico":
                alvo = self.config["datetime"]
                if agora >= alvo:
                    print(f"[AGENDADOR] Disparo único às {agora.strftime('%H:%M:%S')}")
                    self.disparar.emit()
                    break

            elif tipo == "intervalo":
                inicio = self.config["inicio"]
                intervalo_seg = self.config["horas"] * 3600
                if agora >= inicio:
                    if ultimo_disparo is None:
                        self.disparar.emit()
                        ultimo_disparo = agora
                    elif (agora - ultimo_disparo).total_seconds() >= intervalo_seg:
                        self.disparar.emit()
                        ultimo_disparo = agora

            elif tipo == "diario":
                hora_alvo = self.config["hora"]
                alvo = agora.replace(hour=hora_alvo.hour, minute=hora_alvo.minute, second=0, microsecond=0)
                diff_seg = (agora - alvo).total_seconds()
                if 0 <= diff_seg <= 30:
                    if ultimo_disparo is None or (agora - ultimo_disparo).total_seconds() > 3600:
                        self.disparar.emit()
                        ultimo_disparo = agora

            # Sleep em fatias de 0.5s — permite interrupção rápida
            for _ in range(20):  # 20 x 0.5s = 10s total
                if self._parar or self.isInterruptionRequested():
                    break
                time.sleep(0.5)

        print("[AGENDADOR] Thread encerrada.")

class ResumoDialog(QDialog):
    def __init__(self, parent=None, ok=0, falhou=0, invalido=0, 
                 titulo="Resumo do Envio", mensagem=None, segundos=20):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(320)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Monta mensagem automaticamente se não for passada
        if mensagem is None:
            total = ok + falhou + invalido
            mensagem = (
                f"Envio finalizado!\n\n"
                f"Total processado:   {total}\n"
                f"✅ Enviados com sucesso:   {ok}\n"
                f"❌ Falharam:   {falhou}\n"
                f"⚠️ Números inválidos:   {invalido}"
            )

        lbl = QLabel(mensagem)
        lbl.setStyleSheet("font-size: 13px;")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        self.btn_ok = QPushButton(f"OK ({segundos}s)")
        self.btn_ok.setStyleSheet("padding: 8px; font-size: 13px;")
        self.btn_ok.clicked.connect(self.accept)
        layout.addWidget(self.btn_ok)

        self._contador = segundos
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(1000)

    def _tick(self):
        self._contador -= 1
        self.btn_ok.setText(f"OK ({self._contador}s)")
        if self._contador <= 0:
            self._timer.stop()
            self.accept()


class ConexaoThread(QThread):
    conectado = Signal()
    falhou = Signal()

    def __init__(self, main_window):
        super().__init__()
        self.mw = main_window

    def run(self):
        try:

            # Limpa driver morto se existir
            if hasattr(self.mw, 'driver') and self.mw.driver:
                try:
                    _ = self.mw.driver.window_handles
                    self.conectado.emit()
                    return
                except:
                    try:
                        self.mw.driver.quit()
                    except:
                        pass
                    self.mw.driver = None

            # Localiza chromedriver
            if getattr(sys, 'frozen', False):
                path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
            else:
                path = './chromedriver.exe'

            if not os.path.exists(path):
                self.falhou.emit()
                return

            service = Service(path)
            opts = Options()
            opts.add_argument('--disable-gcm')
            opts.add_argument('--log-level=3')
            opts.add_argument('--no-sandbox')
            opts.add_argument('--disable-dev-shm-usage')

            self.mw.driver = webdriver.Chrome(service=service, options=opts)
            self.mw.driver.get("https://web.whatsapp.com")
            print("[INFO] Escaneie o QR Code...")

            # Loop aguardando QR — detecta fechamento do Chrome
            while True:
                try:
                    _ = self.mw.driver.window_handles
                except:
                    print("[INFO] Chrome fechado durante QR")
                    self.mw.driver = None
                    self.falhou.emit()
                    return

                try:
                    self.mw.driver.find_element(By.ID, "side")
                    break
                except:
                    time.sleep(1)

            print("[OK] WhatsApp conectado!")
            self.conectado.emit()

        except Exception as e:
            print(f"[ERRO] ConexaoThread: {e}")
            self.falhou.emit()

# ============================ DIALOG DE PRÉ-ENVIO =============================
class PreEnvioDialog(QDialog):
    def __init__(self, parent=None, total=0, com_nome=0, delay_config=None):
        super().__init__(parent)
        self.setWindowTitle("Pronto para enviar")
        self.setMinimumWidth(400)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setSpacing(16)

        # ── Resumo ──
        layout.addWidget(QLabel("📋 Resumo do envio:"))

        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        frame_layout = QVBoxLayout(frame)

        d = delay_config or {"min": 10, "max": 15}
        tempo_min = total * d["min"]
        tempo_max = total * d["max"]

        frame_layout.addWidget(QLabel(f"👥  Total de contatos:     {total}"))
        frame_layout.addWidget(QLabel(f"✅  Com nome:               {com_nome}"))
        frame_layout.addWidget(QLabel(f"⚠️  Sem nome:               {total - com_nome}"))
        frame_layout.addWidget(QLabel(f"⏱️  Tempo estimado:        {tempo_min//60}m{tempo_min%60}s – {tempo_max//60}m{tempo_max%60}s"))
        layout.addWidget(frame)

        # ── Botões ──
        btns = QHBoxLayout()
        self.btn_agora = QPushButton("▶ Enviar Agora")
        self.btn_agora.setStyleSheet("background-color: #1a7a3c; color: white; padding: 8px; font-size: 14px; border-radius: 6px;")
        self.btn_agendar = QPushButton("🕐 Agendar")
        self.btn_agendar.setStyleSheet("background-color: #1a4a7a; color: white; padding: 8px; font-size: 14px; border-radius: 6px;")
        self.btn_agendar.setEnabled(False)  # futuro

        btns.addWidget(self.btn_agora)
        btns.addWidget(self.btn_agendar)
        layout.addLayout(btns)

        self.btn_agora.clicked.connect(self.accept)
        self.btn_agendar.clicked.connect(self.reject)



# ============================= DIALOG DE ACOMPANHAMENTO =============================

class AcompanhamentoDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Acompanhamento do Envio")
        self.setMinimumWidth(380)
        self.setMinimumHeight(260)
        self.setModal(False)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # ── Contato atual ──
        self.lbl_status = QLabel("Iniciando envio...")
        self.lbl_status.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.lbl_status)

        self.lbl_contato = QLabel("")
        self.lbl_contato.setStyleSheet("font-size: 12px; color: gray;")
        layout.addWidget(self.lbl_contato)

        self.lbl_prog = QLabel("0 de 0 enviados")
        self.lbl_prog.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.lbl_prog)

        # ── Countdown grande ──
        self.lbl_countdown_label = QLabel("")
        self.lbl_countdown_label.setAlignment(Qt.AlignCenter)
        self.lbl_countdown_label.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.lbl_countdown_label)

        self.lbl_countdown = QLabel("")
        self.lbl_countdown.setStyleSheet("font-size: 64px; font-weight: bold; color: #4fc3f7;")
        self.lbl_countdown.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_countdown)

        # ── Previsão de término (abaixo do countdown) ──
        self.lbl_previsao = QLabel("")
        self.lbl_previsao.setAlignment(Qt.AlignCenter)
        self.lbl_previsao.setStyleSheet("color: #aaaaaa; font-size: 12px;")
        layout.addWidget(self.lbl_previsao)

        # ── Botões ──
        btns = QHBoxLayout()
        self.btn_pausar = QPushButton("⏸ Pausar")
        self.btn_pausar.setStyleSheet("background-color: #1565c0; color: white; padding: 8px; border-radius: 6px; font-size: 13px;")
        self.btn_parar = QPushButton("⏹ Parar")
        self.btn_parar.setStyleSheet("background-color: #c62828; color: white; padding: 8px; border-radius: 6px; font-size: 13px;")
        btns.addWidget(self.btn_pausar)
        btns.addWidget(self.btn_parar)
        layout.addLayout(btns)

        # ── Timer ──
        self._countdown_valor = 0
        self._pausado = False
        self._delay_medio = 0
        self._total = 0  
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def iniciar_countdown(self, segundos: int):
        self._pausado = False
        self._countdown_valor = int(segundos)
        self.lbl_countdown_label.setText("próximo envio em:")
        self.lbl_countdown.setText(f"{self._countdown_valor}s")
        self._timer.start(1000)

    def congelar_countdown(self):
        """Pausa o timer mas mantém o número na tela."""
        self._pausado = True
        self._timer.stop()
        self.lbl_countdown_label.setText("⏸ pausado")

    def retomar_countdown(self):
        """Retoma o timer de onde parou."""
        if self._countdown_valor > 0:
            self._pausado = False
            self.lbl_countdown_label.setText("próximo envio em:")
            self._timer.start(1000)

    def parar_countdown(self):
        self._timer.stop()
        self.lbl_countdown.setText("")
        self.lbl_countdown_label.setText("")

    def _tick(self):
        if self._countdown_valor <= 0:
            self._timer.stop()
            self.lbl_countdown.setText("")
            self.lbl_countdown_label.setText("")
            return
        self.lbl_countdown.setText(f"{self._countdown_valor}s")
        self._countdown_valor -= 1

    # Atualiza o atualizar() para calcular a previsão:
    def atualizar(self, atual: int, total: int, numero: str, nome: str, status: str = ""):
        self.lbl_prog.setText(f"{atual} de {total} enviados")
        nome_exib = nome if nome else numero
        self.lbl_contato.setText(f"Contato: {nome_exib}  ({numero})")
        if status:
            self.lbl_status.setText(f"Status: {status}")
        self._atualizar_previsao(atual, total)

    def _atualizar_previsao(self, atual: int, total: int):
        restantes = total - atual
        if restantes > 0 and self._delay_medio > 0:
            from datetime import datetime, timedelta
            segundos_restantes = restantes * self._delay_medio
            previsao = datetime.now() + timedelta(seconds=segundos_restantes)
            self.lbl_previsao.setText(f"Término estimado: {previsao.strftime('%d/%m/%Y às %H:%M')}")
        elif restantes == 0:
            self.lbl_previsao.setText("Finalizando...")

    def closeEvent(self, event):
        event.ignore()  # impede fechar pelo X


    # Atualiza o iniciar_countdown() para guardar o delay médio:
    # Atualiza o iniciar_countdown para também atualizar a previsão:
    def iniciar_countdown(self, segundos: int):
        self._pausado = False
        self._countdown_valor = int(segundos)
        self._delay_medio = segundos
        self.lbl_countdown_label.setText("próximo envio em:")
        self.lbl_countdown.setText(f"{self._countdown_valor}s")
        self._timer.start(1000)
        # Atualiza previsão sempre que chega novo delay
        if self._total > 0:
            atual = int(self.lbl_prog.text().split(" de ")[0])
            self._atualizar_previsao(atual, self._total)

    def parar_countdown(self):
        self._timer.stop()
        self.lbl_countdown.setText("")
        self.lbl_countdown_label.setText("")

    def _tick(self):
        if self._countdown_valor <= 0:
            self._timer.stop()
            self.lbl_countdown.setText("")
            self.lbl_countdown_label.setText("")
            return
        self.lbl_countdown.setText(f"{self._countdown_valor}s")
        self._countdown_valor -= 1

    def closeEvent(self, event):
        event.ignore()  # impede fechar clicando no X



# ============================= DIALOGS =============================
class ConfigDialog(QDialog):
    def __init__(self, parent=None, delay_config=None):
        super().__init__(parent)
        self.setWindowTitle("Configurações")
        self.setMinimumWidth(320)
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Intervalo entre envios:"))

        self.opcoes = QButtonGroup(self)

        self.rb1 = QRadioButton("10 a 15 segundos")
        self.rb2 = QRadioButton("15 a 25 segundos")
        self.rb3 = QRadioButton("25 a 35 segundos")
        self.rb4 = QRadioButton("Personalizado")

        self.opcoes.addButton(self.rb1, 1)
        self.opcoes.addButton(self.rb2, 2)
        self.opcoes.addButton(self.rb3, 3)
        self.opcoes.addButton(self.rb4, 4)

        layout.addWidget(self.rb1)
        layout.addWidget(self.rb2)
        layout.addWidget(self.rb3)
        layout.addWidget(self.rb4)

        # ── Campo personalizado ──
        self.frame_custom = QWidget()
        custom_layout = QHBoxLayout(self.frame_custom)
        custom_layout.setContentsMargins(20, 0, 0, 0)
        custom_layout.addWidget(QLabel("Mínimo (s):"))
        self.spin_min = QSpinBox()
        self.spin_min.setMinimum(1)
        self.spin_min.setMaximum(999)
        self.spin_min.setValue(16)
        custom_layout.addWidget(self.spin_min)
        custom_layout.addWidget(QLabel("Máximo (s):"))
        self.spin_max = QSpinBox()
        self.spin_max.setMinimum(16)
        self.spin_max.setMaximum(999)
        self.spin_max.setValue(30)
        custom_layout.addWidget(self.spin_max)
        self.frame_custom.setEnabled(False)
        layout.addWidget(self.frame_custom)

        # ── Habilita/desabilita o campo personalizado ──
        self.rb4.toggled.connect(self.frame_custom.setEnabled)

        # ── Restaura configuração anterior ──
        if delay_config:
            id_opcao = delay_config.get("opcao", 1)
            self.opcoes.button(id_opcao).setChecked(True)
            self.spin_min.setValue(delay_config.get("min", 16))
            self.spin_max.setValue(delay_config.get("max", 30))
        else:
            self.rb1.setChecked(True)

        # ── Botões OK/Cancelar ──
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._validar)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _validar(self):
        if self.rb4.isChecked():
            if self.spin_min.value() < 15:
                QMessageBox.warning(self, "Aviso", "O mínimo personalizado não pode ser menor que 15 segundos.")
                return
            if self.spin_max.value() <= self.spin_min.value():
                QMessageBox.warning(self, "Aviso", "O máximo deve ser maior que o mínimo.")
                return
        self.accept()

    def get_config(self):
        """Retorna o dict com a configuração escolhida."""
        opcao = self.opcoes.checkedId()
        ranges = {
            1: (10, 15),
            2: (15, 25),
            3: (25, 35),
            4: (self.spin_min.value(), self.spin_max.value())
        }
        minimo, maximo = ranges[opcao]
        return {"opcao": opcao, "min": minimo, "max": maximo}

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


class EsperaQRThread(QThread):
    conectado = Signal()
    falhou = Signal()

    def __init__(self, driver):
        super().__init__()
        self.driver = driver

    def run(self):
        from selenium.webdriver.common.by import By
        while True:
            try:
                _ = self.driver.window_handles
            except:
                print("[INFO] Chrome fechado durante QR")
                self.falhou.emit()
                return

            try:
                self.driver.find_element(By.ID, "side")
                print("[OK] WhatsApp conectado!")
                self.conectado.emit()
                return
            except:
                time.sleep(1)


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
    countdown = Signal(int)
    contato_atual = Signal(int, int, str, str)  # atual, total, numero, nome

    def __init__(self, driver, msg, anexos, tabela, tabela_anexos, delay_config):
        super().__init__()
        pyautogui.FAILSAFE = False  # ADD
        self.driver = driver
        self.msg = msg
        self.anexos = anexos
        self.tabela = tabela
        self.tabela_anexos = tabela_anexos
        self.delay_config = delay_config
        self.pausado = False
        self._parar = False          
        self.lock = threading.Lock()

    def _browser_vivo(self):
        try:
            _ = self.driver.current_url
            return True
        except:
            return False
        
    # Método para sortear o delay
    def _sortear_delay(self):
        return random.uniform(self.delay_config["min"], self.delay_config["max"])

    def run(self):
        total = self.tabela.rowCount()
        ok = 0
        falhou = 0
        numero_invalido = 0

        for row in range(total):
            with self.lock:
                if self._parar: break

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
                    btn_leg = self.tabela_anexos.cellWidget(i, 2)
                    if btn_leg:
                        legenda = btn_leg.property("legenda") or ""
                        legenda = legenda.replace("{nome}", nome)

                    msg_envio = msg_p if i == 0 else None
                    s = self.enviar(numero, mensagem=msg_envio, arquivo=arq, legenda=legenda)
                    status += (", " if status else "") + s

                    # Delay com countdown
                    delay_sorteado = self._sortear_delay()
                    self.countdown.emit(int(delay_sorteado))
                    self.contato_atual.emit(row + 1, total, numero, nome)
                    time.sleep(delay_sorteado)

            else:
                if msg_p:
                    if not self._browser_vivo():
                        print("[ERRO] Browser fechado — encerrando envio")
                        break
                    s = self.enviar(numero, mensagem=msg_p)
                    status += s

                    # Delay com countdown
                    delay_sorteado = self._sortear_delay()
                    self.countdown.emit(int(delay_sorteado))
                    self.contato_atual.emit(row + 1, total, numero, nome)
                    time.sleep(delay_sorteado)

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

    def _clicar_elemento(self, elemento):
        """Clica via JS — funciona mesmo com janela minimizada."""
        try:
            self.driver.execute_script("arguments[0].click();", elemento)
        except:
            elemento.click()  # fallback para click normal


    def _digitar_e_enviar(self, mensagem):
        wait = WebDriverWait(self.driver, 15)
        self._focar_chrome()

        campo = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'p._aupe.copyable-text')
        ))
        self.driver.execute_script("arguments[0].focus(); arguments[0].click();", campo)
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
        """, campo, mensagem)
        time.sleep(0.3)

        btn = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//span[@data-icon='wds-ic-send-filled']")
        ))
        self._clicar_elemento(btn)  # JS click — funciona minimizado


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

        if recarregar:
            self.driver.get(f"https://web.whatsapp.com/send?phone={numero}")
            wait.until(EC.presence_of_element_located((By.ID, "main")))
            time.sleep(3)

        # ── Restaura Chrome antes de qualquer interação ──
        self._focar_chrome()

        # ── Clica no botão + via JS ──
        btn_plus = wait.until(EC.presence_of_element_located(
            (By.XPATH, '(//*[@id="main"]//footer//button//span)[1]')
        ))
        self._clicar_elemento(btn_plus)
        time.sleep(1.2)
        # ── Clica no item do menu via JS ──
        ext = os.path.splitext(arquivo)[1].lower()
        is_media = ext in {'.jpg', '.jpeg', '.png', '.gif', '.mp4', '.mov', '.avi', '.webp'}
        xpath_menu = "//span[text()='Fotos e vídeos']" if is_media else "//button[.//span[text()='Documento']]"
        item = wait.until(EC.presence_of_element_located((By.XPATH, xpath_menu)))
        self._clicar_elemento(item)
        print(f"[INFO] Clicou em '{'Fotos e vídeos' if is_media else 'Documento'}'")
        time.sleep(2.0)

        # ── Restaura Chrome para o pyautogui funcionar ──
        self._focar_chrome()
        time.sleep(1.0)

        arquivo_abs = os.path.abspath(arquivo)
        pyperclip.copy(arquivo_abs)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')
        print(f"[INFO] Arquivo colado no explorer: {arquivo_abs}")
        time.sleep(2.5)

        # ── Aguarda botão enviar ──
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='button' and @aria-label='Enviar']")
        ))

        # ── Legenda via JS ──
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
            except Exception as e:
                print(f"[WARN] Legenda não inserida: {e}")

        # ── Botão enviar via JS ──
        btn_send = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='button' and @aria-label='Enviar']")
        ))
        self._clicar_elemento(btn_send)
        print(f"[OK] Arquivo enviado: {numero}")


    def pausar(self):
        with self.lock:
            self.pausado = not self.pausado
            self.update_pause.emit(self.pausado)

    def parar(self):
        with self.lock:
            self._parar = True
            self.pausado = False
            self.update_pause.emit(False)

    def _focar_chrome(self):
        """Restaura o Chrome sem minimizar — mantém visível atrás."""
        try:
            import win32gui
            import win32con

            def callback(hwnd, hwnds):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if 'chrome' in title.lower() or 'whatsapp' in title.lower():
                        hwnds.append(hwnd)
                return True

            hwnds = []
            win32gui.EnumWindows(callback, hwnds)

            if hwnds:
                hwnd = hwnds[0]
                # Restaura se minimizado — NUNCA deixa minimizado
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1.0)
                return True
        except Exception as e:
            print(f"[WARN] Não foi possível focar o Chrome: {e}")
        return False

# ============================= MAIN WINDOW =============================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhatsApp Sender - VDemo | KAYKY")
        self.setGeometry(100, 100, 1400, 700)
        self.setWindowIcon(QIcon("icone.ico"))
        self.tema_index = 0  # começa no Dark
        self.setStyleSheet(DARK_STYLESHEET)
        self.tema_escuro = True

        self.led_status = QPushButton("🔴 WhatsApp desconectado")
        self.led_status.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #888888;
                font-size: 11px;
                text-align: right;
            }
            QPushButton:hover { color: #0041ce; }
        """)
        self.led_status.setCursor(Qt.PointingHandCursor)
        self.led_status.clicked.connect(self._clicar_led)

        # === BARRA DE AGENDAMENTO ===
        self.barra_agendamento = BarraAgendamento()
        self.barra_agendamento.cancelar_signal.connect(self._cancelar_agendamento)

        # === PROGRESSO ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        # === TABELA CONTATOS ===
        self.tabela = QTableWidget(0, 3)
        self.tabela.setHorizontalHeaderLabels(["Nome", "Número", "Status"])
        self.tabela.verticalHeader().setDefaultSectionSize(32)  # altura das linhas
        self.tabela.verticalHeader().setVisible(False)  # esconde o número da linha (1, 2, 3...)
        self.tabela.setMinimumWidth(350)
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

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
        self.tabela_anexos.setHorizontalHeaderLabels(["Arquivo", "Preview", "Legenda"])
        self.tabela_anexos.setColumnWidth(1, 80)
        self.tabela_anexos.verticalHeader().setDefaultSectionSize(60)
        self.tabela_anexos.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Mas a coluna de preview precisa de tamanho fixo:
        self.tabela_anexos.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabela_anexos.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.tabela_anexos.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tabela_anexos.setColumnWidth(1, 80)

        
        btns_anex = QHBoxLayout()
        btns_anex.addWidget(self.btn("Anexar", self.anexar))
        btns_anex.addWidget(self.btn("Limpar", self.limpar_anexos))

        btns_ctrl = QHBoxLayout()
        btns_ctrl.addWidget(self.btn("Timer", self.config))
        self.btn_enviar = self.btn("ENVIAR", self.enviar, "green")
        btns_ctrl.addWidget(self.btn_enviar)
        self.btn_tema = self.btn("Tema: Dark Green", self.trocar_tema, "#333")
        btns_ctrl.addWidget(self.btn_tema)

        center = QVBoxLayout()
        center.addWidget(self.barra_agendamento)
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
        self.tabela_log.verticalHeader().setDefaultSectionSize(32)
        self.tabela_log.verticalHeader().setVisible(False)
        self.tabela_log.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        right = QVBoxLayout()
        right.addWidget(QLabel("Registros"))
        right.addWidget(self.tabela_log)
        right.addWidget(self.btn("Limpar Logs", self.limpar_logs))

        # === MAIN LAYOUT ===
        main = QHBoxLayout()
        main.addLayout(left, 1)
        main.addLayout(center, 2)
        main.addLayout(right, 1)


        # No topo do main layout, antes do container principal
        topo = QHBoxLayout()
        topo.addStretch()
        topo.addWidget(self.led_status)

        main_vertical = QVBoxLayout()
        main_vertical.setContentsMargins(0, 4, 8, 0)
        main_vertical.addLayout(topo)
        main_vertical.addLayout(main)  # seu layout principal existente

        container = QWidget()
        container.setLayout(main_vertical)
        self.setCentralWidget(container)


        # === VARIAVEIS ===
        self.delay_config = {"opcao": 1, "min": 1, "max": 5}  # padrão 1 a 5s
        self.anexos = []
        self.driver = None
        self.thread = None
        self.agendador = None
        self._iniciando_driver = False  # ADD
        self.lock = threading.Lock()

        # No __init__, junto com as outras variáveis:
        self._timer_led = QTimer(self)
        self._timer_led.timeout.connect(self._atualizar_led)
        self._timer_led.start(5000)  # verifica a cada 5s

    def _cancelar_agendamento(self):
        resposta = QMessageBox.question(self, "Cancelar Agendamento",
            "Tem certeza que deseja cancelar o agendamento?",
            QMessageBox.Yes | QMessageBox.No)
        if resposta == QMessageBox.Yes:
            if self.agendador:
                self.agendador.parar()
                self.agendador.quit()  # pede para a thread parar
                self.agendador.wait(3000)  # aguarda até 3s
                self.agendador = None
            self.barra_agendamento.desativar()
            print("[INFO] Agendamento cancelado pelo usuário")




    def _resumo_agendamento(self, config):
        tipo = config["tipo"]
        if tipo == "unico":
            return f"Disparo único em: {config['datetime'].strftime('%d/%m/%Y às %H:%M')}"
        elif tipo == "intervalo":
            return f"A cada {config['horas']}h — início: {config['inicio'].strftime('%H:%M')}"
        elif tipo == "diario":
            return f"Todo dia às {config['hora'].strftime('%H:%M')}"


    def _clicar_led(self):
        if not self._whatsapp_conectado():
            self.led_status.setText("⏳ Conectando...")
            self.led_status.setEnabled(False)
            self.iniciar_driver()
            self.led_status.setEnabled(True)
            self._atualizar_led()

    def _on_conectado(self):
        self._thread_conexao.quit()
        self._thread_conexao.wait()
        self.led_status.setEnabled(True)
        self._atualizar_led()
        print("[OK] LED atualizado para verde")

    def _on_conexao_falhou(self):
        self._thread_conexao.quit()
        self._thread_conexao.wait()
        self.led_status.setEnabled(True)
        self._atualizar_led()
        print("[INFO] Conexão cancelada ou falhou")


    def _whatsapp_conectado(self):
        try:
            self.driver.find_element(By.ID, "side")
            return True
        except:
            return False

    def _atualizar_led(self):
        if getattr(self, '_iniciando_driver', False):
            return
        
        # Não interfere durante envio ativo
        if hasattr(self, 'thread') and self.thread and self.thread.isRunning():
            self.led_status.setText("🟢 WhatsApp conectado")
            return

        conectado = False
        if hasattr(self, 'driver') and self.driver:
            try:
                _ = self.driver.window_handles
                self.driver.find_element(By.ID, "side")
                conectado = True
            except:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

        if conectado:
            self.led_status.setText("🟢 WhatsApp conectado")
            self.led_status.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    color: #4caf50;
                    font-size: 11px;
                }
                QPushButton:hover { color: #81c784; }
            """)
            self.led_status.setCursor(Qt.ArrowCursor)
        else:
            self.led_status.setText("🔴 WhatsApp desconectado")
            self.led_status.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    color: #888888;
                    font-size: 11px;
                }
                QPushButton:hover { color: #ffffff; }
            """)
            self.led_status.setCursor(Qt.PointingHandCursor)

        # No PreEnvioDialog, o btn_agendar agora chama:
    def _abrir_agendar(self, pre_dialog):
        pre_dialog.hide()
        dlg = AgendarDialog(self)
        if dlg.exec():
            config = dlg.get_config()
            self._iniciar_agendador(config)
            pre_dialog.done(2)
        else:
            pre_dialog.show()

    def _iniciar_agendador(self, config: dict):
        if not self.iniciar_driver():
            QMessageBox.warning(self, "Aviso", "Falha ao conectar o WhatsApp.")
            return

        # Se já tem agendador rodando, cancela antes
        if self.agendador:
            self.agendador.parar()
            self.agendador.deleteLater()
            self.agendador = None

        self.agendador = AgendadorThread(config)
        self.agendador.disparar.connect(self._disparar_agendado)
        self.agendador.start()

        # Calcula próximo disparo para o countdown
        from datetime import datetime, timedelta
        agora = datetime.now()
        if config["tipo"] == "unico":
            proximo = config["datetime"]
        elif config["tipo"] == "intervalo":
            proximo = config["inicio"]
        elif config["tipo"] == "diario":
            hora = config["hora"]
            proximo = agora.replace(hour=hora.hour, minute=hora.minute, second=0)
            if proximo <= agora:
                proximo += timedelta(days=1)

        self.barra_agendamento.ativar(
            f"⚡ Agendamento ativo: {self._resumo_agendamento(config)}",
            proximo_disparo=proximo
        )

        dlg = ResumoDialog(self, 
            titulo="✅ Envio Agendado!",
            mensagem=f"{self._resumo_agendamento(config)}\n\nO WhatsApp já está conectado e aguardando.",
            segundos=20
        )
        dlg.exec()


    def _disparar_agendado(self):
        from datetime import datetime, timedelta
        tipo = self.agendador.config["tipo"] if self.agendador else "unico"

        if tipo == "unico":
            self.barra_agendamento.desativar()
            if self.agendador:
                self.agendador.parar()
                self.agendador.deleteLater()
                self.agendador = None
        elif tipo == "intervalo" and self.agendador:
            # Atualiza countdown para próximo intervalo
            proximo = datetime.now() + timedelta(hours=self.agendador.config["horas"])
            self.barra_agendamento.atualizar_proximo(proximo)
        elif tipo == "diario" and self.agendador:
            # Atualiza countdown para amanhã
            hora = self.agendador.config["hora"]
            proximo = datetime.now().replace(hour=hora.hour, minute=hora.minute, second=0)
            proximo += timedelta(days=1)
            self.barra_agendamento.atualizar_proximo(proximo)

        if not self.iniciar_driver():
            return
        self._iniciar_thread_envio()


    def _iniciar_thread_envio(self):
        delay_config = getattr(self, 'delay_config', {"opcao": 1, "min": 10, "max": 15})

        self.acomp = AcompanhamentoDialog(self)
        self.acomp.btn_pausar.clicked.connect(self.pausar)
        self.acomp.btn_parar.clicked.connect(self._parar_e_fechar_acomp)
        self.acomp.show()

        self.btn_enviar.setEnabled(False)
        self.progress_bar.setValue(0)

        with self.lock:
            self.thread = SenderThread(
                self.driver,
                self.txt_msg.toPlainText(),
                self.anexos,
                self.tabela,
                self.tabela_anexos,
                delay_config,
            )
            self.thread.update_status.connect(lambda r, s: self.tabela.setItem(r, 2, QTableWidgetItem(s)))
            self.thread.add_log.connect(lambda n, s: self.add_log(n, s))
            self.thread.finished.connect(self.finalizado)
            self.thread.update_pause.connect(self._on_pause)
            self.thread.progress.connect(self.atualizar_progresso)
            self.thread.countdown.connect(self.acomp.iniciar_countdown)
            self.thread.contato_atual.connect(self._atualizar_acomp)
            self.thread.start()


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
        self.tema_index = (self.tema_index + 1) % len(NOMES_TEMAS)
        nome = NOMES_TEMAS[self.tema_index]
        self.setStyleSheet(TEMAS[nome])
        self.btn_tema.setText(f"Tema: {nome}")

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

            # ── Coluna 0: nome do arquivo ──
            self.tabela_anexos.setItem(r, 0, QTableWidgetItem(path))

            # ── Coluna 1: preview ──
            self._inserir_preview(r, path)

            # ── Coluna 2: legenda ──
            self._inserir_btn_legenda(r, "")
            self.anexos.append(path)

    def _inserir_preview(self, row, path):
        from PySide6.QtGui import QPixmap, QIcon
        from PySide6.QtCore import QSize
        import os

        ext = os.path.splitext(path)[1].lower()
        label = QLabel()
        label.setAlignment(Qt.AlignCenter)
        label.setFixedSize(76, 56)

        if ext in {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp'}:
            # ── Imagem: mostra thumbnail real ──
            pixmap = QPixmap(path)
            if not pixmap.isNull():
                pixmap = pixmap.scaled(72, 52, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                label.setPixmap(pixmap)
            else:
                label.setText("🖼️")
                label.setStyleSheet("font-size: 24px;")

        elif ext in {'.mp4', '.mov', '.avi', '.mkv', '.webm'}:
            # ── Vídeo: tenta extrair primeiro frame ──
            try:
                import cv2
                cap = cv2.VideoCapture(path)
                ret, frame = cap.read()
                cap.release()
                if ret:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    h, w, ch = frame.shape
                    from PySide6.QtGui import QImage
                    img = QImage(frame.data, w, h, ch * w, QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(img).scaled(72, 52, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    # Adiciona ícone de play por cima
                    overlay = QLabel(label)
                    overlay.setText("▶")
                    overlay.setStyleSheet("color: white; font-size: 18px; background: rgba(0,0,0,120); border-radius: 4px; padding: 2px 6px;")
                    overlay.move(26, 18)
                    label.setPixmap(pixmap)
                else:
                    raise Exception("Frame vazio")
            except:
                label.setText("🎬")
                label.setStyleSheet("font-size: 24px;")

        elif ext == '.pdf':
            label.setText("📄")
            label.setStyleSheet("font-size: 24px;")

        elif ext in {'.doc', '.docx'}:
            label.setText("📝")
            label.setStyleSheet("font-size: 24px;")

        elif ext in {'.xls', '.xlsx'}:
            label.setText("📊")
            label.setStyleSheet("font-size: 24px;")

        elif ext in {'.zip', '.rar', '.7z'}:
            label.setText("🗜️")
            label.setStyleSheet("font-size: 24px;")

        else:
            label.setText("📎")
            label.setStyleSheet("font-size: 24px;")

        self.tabela_anexos.setCellWidget(row, 1, label)
        self.tabela_anexos.setRowHeight(row, 60)

    def _inserir_btn_legenda(self, row, texto=""):
        btn = QPushButton("✏️ Adicionar legenda" if not texto else f"✏️ {texto[:25]}...")
        btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: 1px dashed #3e3e6e;
                color: #8888cc;
                border-radius: 6px;
                padding: 4px 8px;
                text-align: left;
                font-size: 11px;
            }
            QPushButton:hover {
                border-color: #7777cc;
                color: #ffffff;
                background: #1e1e3a;
            }
        """)
        btn.setProperty("legenda", texto)
        btn.clicked.connect(lambda: self._editar_legenda(row, btn))
        self.tabela_anexos.setCellWidget(row, 2, btn)

    def _editar_legenda(self, row, btn):
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QDialogButtonBox, QHBoxLayout

        dlg = QDialog(self)
        dlg.setWindowTitle("Legenda do anexo")
        dlg.setMinimumWidth(400)
        layout = QVBoxLayout(dlg)

        lbl = QLabel("Digite a legenda (use {nome} para personalizar):")
        lbl.setStyleSheet("color: #8888cc; font-size: 12px;")
        layout.addWidget(lbl)

        txt = QTextEdit()
        txt.setPlaceholderText("Ex: Olá {nome}, segue o arquivo...")
        txt.setMinimumHeight(100)
        txt.setText(btn.property("legenda"))
        layout.addWidget(txt)

        # ── Botão Nome ──
        btn_nome = QPushButton("Nome")
        btn_nome.setStyleSheet("""
            QPushButton {
                background: #1e1e3a;
                color: #8888cc;
                border: 1px solid #3e3e6e;
                border-radius: 6px;
                padding: 5px 12px;
                font-size: 12px;
            }
            QPushButton:hover { background: #2a2a5a; color: #ffffff; }
        """)
        btn_nome.clicked.connect(lambda: txt.insertPlainText("{nome}"))

        # ── Botões OK/Cancelar ──
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)

        # ── Layout inferior com Nome + OK/Cancelar ──
        bottom = QHBoxLayout()
        bottom.addWidget(btn_nome)
        bottom.addStretch()
        bottom.addWidget(btns)
        layout.addLayout(bottom)

        if dlg.exec():
            legenda = txt.toPlainText().strip()
            btn.setProperty("legenda", legenda)
            if legenda:
                preview = legenda[:25] + "..." if len(legenda) > 25 else legenda
                btn.setText(f"✏️ {preview}")
            else:
                btn.setText("✏️ Adicionar legenda")

    def limpar_anexos(self):
        self.tabela_anexos.setRowCount(0)
        self.anexos.clear()

    # === CONTROLE ===
    def config(self):
        dlg = ConfigDialog(self, delay_config=getattr(self, 'delay_config', None))
        if dlg.exec():
            self.delay_config = dlg.get_config()
            print(f"[INFO] Delay configurado: {self.delay_config['min']}s a {self.delay_config['max']}s")

    def iniciar_driver(self):
        from PySide6.QtWidgets import QApplication

        self._iniciando_driver = True  # bloqueia _atualizar_led AQUI no início

        if hasattr(self, 'driver') and self.driver:
            try:
                _ = self.driver.window_handles
                self._iniciando_driver = False
                return True
            except:
                print("[INFO] Chrome fechado detectado — reiniciando...")
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

        if getattr(sys, 'frozen', False):
            path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
        else:
            path = './chromedriver.exe'

        if not os.path.exists(path):
            QMessageBox.critical(self, "Erro Fatal", f"chromedriver.exe não encontrado!\nEsperado em: {path}")
            self._iniciando_driver = False
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

            while True:
                QApplication.processEvents()  # _atualizar_led vai retornar imediatamente por causa da flag

                if self.driver is None:
                    print("[DEBUG] driver zerado externamente — abortando")
                    self._iniciando_driver = False
                    return False

                try:
                    _ = self.driver.window_handles
                except:
                    print("[INFO] Chrome fechado durante QR — abortando")
                    self.driver = None
                    self._iniciando_driver = False
                    return False

                try:
                    self.driver.find_element(By.ID, "side")
                    break
                except:
                    time.sleep(0.5)

            print("[OK] WhatsApp conectado!")
            self._iniciando_driver = False
            self._atualizar_led()
            return True

        except Exception as e:
            print(f"[ERRO] Falha ao iniciar Chrome: {e}")
            self.driver = None
            self._iniciando_driver = False
            QMessageBox.critical(self, "Erro", f"Falha ao iniciar Chrome:\n{e}")
            return False    

    # Método enviar() da MainWindow — substitui pelo novo fluxo:
    def enviar(self):
        if self.tabela.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Adicione contatos")
            return
        if not self.txt_msg.toPlainText() and not self.anexos:
            QMessageBox.warning(self, "Aviso", "Digite mensagem ou anexe")
            return

        total = self.tabela.rowCount()
        com_nome = sum(
            1 for r in range(total)
            if self.tabela.item(r, 0) and self.tabela.item(r, 0).text().strip()
        )
        delay_config = getattr(self, 'delay_config', {"opcao": 1, "min": 10, "max": 15})

        # ── Abre janela de pré-envio ──
        pre = PreEnvioDialog(self, total=total, com_nome=com_nome, delay_config=delay_config)
        pre.btn_agendar.setEnabled(True)
        pre.btn_agendar.clicked.connect(lambda: self._abrir_agendar(pre))

        resultado = pre.exec()

        if resultado == QDialog.Accepted:
            # Enviar agora
            if not self.iniciar_driver():
                return
            self._iniciar_thread_envio()

        # resultado == 2 significa agendado — não faz nada aqui, o agendador cuida


    def _atualizar_acomp(self, atual, total, num, nome):
        if hasattr(self, 'acomp') and self.acomp:
             self.acomp.atualizar(int(atual), int(total), num, nome)

    def _on_pause(self, pausado):
        if hasattr(self, 'acomp') and self.acomp:
            self.acomp.btn_pausar.setText("▶ Retomar" if pausado else "⏸ Pausar")
            if pausado:
                self.acomp.congelar_countdown()
            else:
                self.acomp.retomar_countdown()

    def _parar_e_fechar_acomp(self):
        # Para o envio
        if hasattr(self, 'thread') and self.thread:
            self.thread._parar = True

        # Fecha só a janela de acompanhamento
        if hasattr(self, 'acomp') and self.acomp:
            self.acomp.parar_countdown()
            self.acomp.closeEvent = lambda e: e.accept()
            self.acomp.close()

        self._resetar_ui()

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
        print(f"[DEBUG] finalizado chamado: ok={ok} falhou={falhou} invalido={invalido}")
        
        if hasattr(self, 'acomp'):
            self.acomp.parar_countdown()
            self.acomp.closeEvent = lambda e: e.accept()
            self.acomp.close()

        self._resetar_ui()

        if hasattr(self, 'driver'):
            try:
                _ = self.driver.current_url
            except:
                try:
                    self.driver.quit()
                except:
                    pass
                finally:
                    delattr(self, 'driver')

        if ok + falhou + invalido > 0:
            dlg = ResumoDialog(self, ok=ok, falhou=falhou, invalido=invalido)
            dlg.exec()


    def _resetar_ui(self):
        self.btn_enviar.setEnabled(True)
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