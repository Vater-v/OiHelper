import sys
import win32gui
import win32api
import win32con
import win32process
import requests
import threading
import os
import subprocess
import zipfile
import time
import logging
import ctypes
from ctypes import wintypes, windll
import queue
import random
import math
from typing import Optional, List, Tuple

# Пытаемся импортировать библиотеки для визуального контроля и устанавливаем флаги
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    

# ===================================================================
# -2. ИМПОРТЫ PyQt6
# ===================================================================
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel,
    QGridLayout, QHBoxLayout, QFrame, QSizePolicy, QGraphicsDropShadowEffect,
    QComboBox
)
from PyQt6.QtCore import QObject, QTimer, pyqtSignal, Qt, QPropertyAnimation, QRect, QEasingCurve, pyqtProperty
from PyQt6.QtGui import QIcon, QColor, QFont, QPainter, QBrush, QPen
from PyQt6.QtSvg import QSvgRenderer


# ===================================================================
# -1. ПРОВЕРКА ЗАПУСКА ОДНОЙ КОПИИ (MUTEX)
# ===================================================================
class SingleInstance:
    """Класс для проверки и удержания мьютекса, чтобы предотвратить запуск второй копии приложения."""
    def __init__(self, name):
        self.mutex = windll.kernel32.CreateMutexW(None, False, name)
        self.last_error = windll.kernel32.GetLastError()

    def is_already_running(self):
        return self.last_error == 183 # ERROR_ALREADY_EXISTS

# ===================================================================
# 0. НАСТРОЙКА ЛОГИРОВАНИЯ
# ===================================================================
try:
    log_dir = os.path.join(os.getenv('APPDATA'), 'OiHelper')
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, 'app.log')
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', force=True)

    file_handler = logging.FileHandler(log_file_path, encoding='utf-8', mode='a')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    root_logger = logging.getLogger()
    root_logger.addHandler(file_handler)
except Exception as e:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.error(f"Не удалось настроить логирование в файл: {e}")

# ===================================================================
# 1. КОНФИГУРАЦИЯ И СТИЛИ
# ===================================================================

class AppConfig:
    """Централизованная конфигурация приложения."""
    DEBUG_MODE = False
    CURRENT_VERSION = "6.05"
    GITHUB_REPO = "Vater-v/OiHelper"
    ASSET_NAME = "OiHelper.zip"
    ICON_PATH = 'icon.ico'

    TELEGRAM_BOT_TOKEN = os.environ.get("OIHELPER_TG_TOKEN", '')
    TELEGRAM_CHAT_ID = os.environ.get("OIHELPER_TG_CHAT_ID", '')

    PLAYER_LAUNCHER_TITLE = "Holdem"
    PLAYER_GAME_LAUNCHER_TITLE = "Game"
    CAMTASIA_PROCESS_NAME = "recorder"
    INJECTOR_WINDOW_TITLE = "injector"

    TELEGRAM_REPORT_LEVEL = 'all'
    LOG_SENDER_KEYWORDS = ["endsess", "logbot"]
    LOG_SENDER_TIMEOUT_S = 300
    PLAYER_RELAUNCH_DELAY_S = 15
    RECORD_RESTART_COOLDOWN_S = 10

    DEFAULT_WIDTH = 350
    DEFAULT_HEIGHT = 250 # Increased for dropdown
    GG_UI_WIDTH = 800
    GG_UI_HEIGHT = 100
    WINDOW_MARGIN = 1

    APP_TITLE_TEMPLATE = "OiHelper v{version}"
    LOBBY_MSG_UPDATE_CHECK = "Проверка обновлений..."
    LOBBY_MSG_SEARCHING = "Поиск проекта..."
    LOBBY_MSG_PLAYER_NOT_FOUND = "Плеер не найден. Запустите его."

    MSG_ARRANGE_TABLES = "Расставить столы"
    MSG_ARRANGE_SYSTEM = "Системные окна"
    MSG_CLICK_COMMAND = "Закрыть все столы"
    MSG_PROGRESS_LABEL = "Перезапись через: {}"
    MSG_LIMIT_REACHED = "<b>Лимит!</b>"
    MSG_UPDATE_FAIL = "Ошибка обновления. Работа в оффлайн-режиме."
    MSG_UPTIME_WARNING = "Компьютер не перезагружался более 5 дней."
    MSG_ADMIN_WARNING = "Нет прав администратора. Функции могут быть ограничены."
    MSG_ARRANGE_TABLES_NOT_FOUND = "Столы для расстановки не найдены."
    MSG_PROJECT_UNDEFINED = "Проект не определен. Расстановка невозможна."

    PLAYER_CHECK_INTERVAL = 3000
    PLAYER_AUTOSTART_INTERVAL = 5000
    AUTO_RECORD_INTERVAL = 3000
    AUTO_ARRANGE_INTERVAL = 2000
    RECORDER_CHECK_INTERVAL = 10000
    POPUP_CHECK_INTERVAL_FAST = 750
    POPUP_CHECK_INTERVAL_SLOW = 10000
    POPUP_FAST_SCAN_DURATION_S = 120
    NOTIFICATION_DURATION = 4000
    STATUS_MESSAGE_DURATION = 3000

logging.getLogger().setLevel(logging.DEBUG if AppConfig.DEBUG_MODE else logging.INFO)

class ColorPalette:
    BACKGROUND = "#F9FAFB"
    SURFACE = "#FFFFFF"
    PRIMARY = "#3B82F6"
    PRIMARY_HOVER = "#60A5FA"
    PRIMARY_PRESSED = "#2563EB"
    SECONDARY = "#E5E7EB"
    SECONDARY_HOVER = "#D1D5DB"
    SECONDARY_PRESSED = "#9CA3AF"
    GREEN = "#10B981"
    RED = "#EF4444"
    AMBER = "#F59E0B"
    TEXT_PRIMARY = "#1F2937"
    TEXT_SECONDARY = "#6B7280"
    BORDER = "#D1D5DB"

class StyleSheet:
    MAIN_WINDOW = f"background-color: {ColorPalette.BACKGROUND};"
    SPLASH_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 14px; color: {ColorPalette.TEXT_SECONDARY};"
    LOBBY_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 14px; color: {ColorPalette.TEXT_SECONDARY};"
    STATUS_LABEL = f"font-family: 'Segoe UI Semibold', 'Roboto'; font-size: 13px; color: {ColorPalette.TEXT_PRIMARY};"
    PROGRESS_BAR_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 11px; color: {ColorPalette.TEXT_SECONDARY};"
    COMBO_BOX = f"""
        QComboBox {{
            font-family: 'Segoe UI', 'Roboto';
            font-size: 12px;
            color: {ColorPalette.TEXT_PRIMARY};
            background-color: {ColorPalette.SURFACE};
            border: 1px solid {ColorPalette.BORDER};
            border-radius: 6px;
            padding: 5px 10px;
        }}
        QComboBox:hover {{
            border-color: {ColorPalette.PRIMARY_HOVER};
        }}
        QComboBox::drop-down {{
            border: none;
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
        }}
        QComboBox::down-arrow {{
            image: url(down_arrow.png); 
        }}
        QComboBox QAbstractItemView {{
            background-color: {ColorPalette.SURFACE};
            border: 1px solid {ColorPalette.BORDER};
            selection-background-color: {ColorPalette.PRIMARY};
            selection-color: #FFFFFF;
        }}
    """

    @staticmethod
    def get_button_style(primary=True):
        base_style = "QPushButton { font-family: 'Segoe UI Semibold', 'Roboto'; font-size: 12px; border-radius: 6px; padding: 8px 12px; }"
        if primary:
            return base_style + f"QPushButton {{ background-color: {ColorPalette.PRIMARY}; color: #FFFFFF; border: none; }} QPushButton:hover {{ background-color: {ColorPalette.PRIMARY_HOVER}; }} QPushButton:pressed {{ background-color: {ColorPalette.PRIMARY_PRESSED}; }} QPushButton:disabled {{ background-color: {ColorPalette.SECONDARY}; color: #9CA3AF; }}"
        else:
            return base_style + f"QPushButton {{ background-color: {ColorPalette.SURFACE}; color: {ColorPalette.TEXT_PRIMARY}; border: 1px solid {ColorPalette.BORDER}; }} QPushButton:hover {{ background-color: {ColorPalette.BACKGROUND}; }} QPushButton:pressed {{ background-color: {ColorPalette.SECONDARY}; }} QPushButton:disabled {{ background-color: {ColorPalette.BACKGROUND}; color: #9CA3AF; border-color: {ColorPalette.SECONDARY}; }}"

PROJECT_CONFIGS = {
    "GG": {
        "TABLE": {
            "FIND_METHOD": "RATIO",
            "W": 557,      # ширина окна ClubGG
            "H": 424,      # высота окна ClubGG
            "TOLERANCE": 0.035
        },
        "LOBBY": {
            "FIND_METHOD": "RATIO",
            "W": 333,
            "H": 623,
            "TOLERANCE": 0.07,
            "X": 1580,     # ПРАВО от столов, скорректируй под свою рабочую область!
            "Y": 140
        },
        "PLAYER": {
            "W": 700,
            "H": 365,
            "X": 1385,
            "Y": 0
        },
        "TABLE_SLOTS": [
            (-5, 0),   # 1 стол: левый верх
            (271, 423),   # 2 стол: левый низ
            (816, 0),   # 3 стол: правый верх
            (1086, 423),   # 4 стол: правый низ
        ],
        "EXCLUDED_TITLES": [
            "OiHelper", "NekoRay", "NekoBox", "Chrome",
            "Sandbo", "Notepad", "Explorer"
        ],
        "EXCLUDED_PROCESSES": [
            "explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe",
            "Taskmgr.exe", "firefox.exe", "msedge.exe",
            "RuntimeBroker.exe", "ApplicationFrameHost.exe",
            "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe"
        ],
        "SESSION_MAX_DURATION_S": 4 * 3600,
        "SESSION_WARN_TIME_S": 3.5 * 3600,
        "ARRANGE_MINIMIZED_TABLES": False,
        "POPUPS": {
            "CONFIDENCE": 0.85,
            "BUY_IN": {
                "trigger": "buyin_window.png",
                "max_button": "max_button.png",
                "confirm_button": "buyin_button.png"
            }
        }
    },
    "QQ": { 
        "TABLE": { "FIND_METHOD": "TITLE_AND_SIZE", "TITLE": "QQPK", "W": 400, "H": 700, "TOLERANCE": 2 }, 
        "LOBBY": { "FIND_METHOD": "RATIO", "W": 400, "H": 700, "TOLERANCE": 0.07, "X": 1418, "Y": 0 }, 
        "CV_SERVER": { "FIND_METHOD": "TITLE", "TITLE": "OpenCv", "X": 1789, "Y": 367, "W": 993, "H": 605 }, 
        "PLAYER": { "X": 1418, "Y": 942, "W": 724, "H": 370 }, 
        "TABLE_SLOTS": [(0, 0), (401, 0), (802, 0), (1203, 0)], 
        "TABLE_SLOTS_5": [(0, 0), (346, 0), (692, 0), (1038, 0), (1384, 0)], 
        "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], 
        "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "OpenCvServer.exe"], 
        "SESSION_MAX_DURATION_S": 3 * 3600, 
        "SESSION_WARN_TIME_S": -1, 
        "ARRANGE_MINIMIZED_TABLES": True,
        "POPUPS": {
            "CONFIDENCE": 0.83,
            "SPAM": [
                {"poster": "popup_poster_1.png", "close_button": "popup_close_btn_1.png"},
                {"poster": "popup_poster_2.png", "close_button": "popup_close_btn_2.png"},
                {"poster": "popup_poster_3.png", "close_button": "popup_close_button.png"},
                {"poster": "popup_poster_4.png", "close_button": "popup_close_button.png"}
            ],
            "BONUS": {
                "wheel": "bonus_wheel.png",
                "spin_button": "bonus_spin_button.png",
                "close_button": "bonus_close_button.png"
            }
        }
    }
}

# ===================================================================
# 2. СИСТЕМА ИНТЕРФЕЙСА (SPLASH, УВЕДОМЛЕНИЯ, ИНДИКАТОРЫ)
# ===================================================================

class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(300, 100)
        container = QFrame(self)
        container.setStyleSheet(f"background-color: {ColorPalette.SURFACE}; border-radius: 8px;")
        # ИСПРАВЛЕНО: Эффект создается без родителя, чтобы избежать двойного владения.
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 4)
        container.setGraphicsEffect(shadow)
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(container)
        layout = QVBoxLayout(container)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label = QLabel("Загрузка OiHelper...")
        self.status_label.setStyleSheet(StyleSheet.SPLASH_LABEL)
        layout.addWidget(self.status_label)
    def update_status(self, message: str): self.status_label.setText(message)

class ToggleSwitch(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFixedSize(40, 22)
        self.clicked.connect(self.update_state)
        self._circle_pos = 3
        self.animation = QPropertyAnimation(self, b"circle_pos", self)
        self.animation.setDuration(200)
        self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)
    def get_circle_pos(self): return self._circle_pos
    def set_circle_pos(self, pos): self._circle_pos = pos; self.update()
    circle_pos = pyqtProperty(int, fget=get_circle_pos, fset=set_circle_pos)
    def update_state(self): self.animation.setStartValue(self._circle_pos); self.animation.setEndValue(20 if self.isChecked() else 3); self.animation.start()
    def setChecked(self, checked): super().setChecked(checked); self.update_state()
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        bg_color = QColor(ColorPalette.PRIMARY) if self.isChecked() else QColor(ColorPalette.SECONDARY)
        painter.setBrush(QBrush(bg_color))
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 11, 11)
        painter.setBrush(QBrush(QColor(ColorPalette.SURFACE)))
        painter.drawEllipse(self._circle_pos, 3, 16, 16)

class Notification(QWidget):
    closed = pyqtSignal(QWidget)
    COLORS = {"info": ColorPalette.PRIMARY, "warning": ColorPalette.AMBER, "error": ColorPalette.RED}
    def __init__(self, message, message_type):
        super().__init__()
        self.message = message
        self.message_type = message_type
        self.is_closing = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        container = QFrame(self)
        color = self.COLORS.get(message_type, self.COLORS["info"])
        container.setStyleSheet(f"background-color: {ColorPalette.SURFACE}; border-radius: 8px; border-left: 6px solid {color};")
        shadow_layout = QVBoxLayout(self)
        shadow_layout.setContentsMargins(8, 8, 8, 8)
        shadow_layout.addWidget(container)
        # ИСПРАВЛЕНО: Эффект создается без родителя, чтобы избежать двойного владения.
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 90))
        shadow.setOffset(0, 4)
        container.setGraphicsEffect(shadow)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(20, 15, 20, 15)
        text_label = QLabel(message)
        text_label.setFont(QFont("Segoe UI", 14))
        text_label.setWordWrap(True)
        text_label.setStyleSheet(f"background: transparent; border: none; color: {ColorPalette.TEXT_PRIMARY};")
        layout.addWidget(text_label, 1)
        self.setFixedWidth(420)
        self.opacity = 0.0
        self.setWindowOpacity(self.opacity)
        self.fade_in_timer = QTimer(self)
        self.fade_in_timer.timeout.connect(self.fade_in)
        self.fade_out_timer = QTimer(self)
        self.fade_out_timer.timeout.connect(self.fade_out)
        QTimer.singleShot(AppConfig.NOTIFICATION_DURATION, self.start_fade_out)
    def fade_in(self):
        self.opacity += 0.1
        self.setWindowOpacity(min(self.opacity, 1.0))
        if self.opacity >= 1.0: self.fade_in_timer.stop()
    def fade_out(self):
        self.opacity -= 0.1
        self.setWindowOpacity(max(self.opacity, 0.0))
        if self.opacity <= 0.0: self.fade_out_timer.stop(); self.close()
    def show_animation(self): self.show(); self.fade_in_timer.start(20)
    def start_fade_out(self):
        if self.is_closing: return
        self.is_closing = True
        self.fade_in_timer.stop()
        self.fade_out_timer.start(20)
    def closeEvent(self, event): self.closed.emit(self); super().closeEvent(event)

class NotificationManager(QObject):
    def __init__(self):
        super().__init__()
        self.notifications = []
    def show(self, message: str, message_type: str):
        logging.info(f"Уведомление [{message_type}]: {message}")
        if any(n.message == message for n in self.notifications): return
        if len(self.notifications) >= 5: self.notifications.pop(0).start_fade_out()
        notification = Notification(message, message_type)
        notification.closed.connect(self.on_notification_closed)
        self.notifications.append(notification)
        self.reposition_all()
        notification.show_animation()
    def on_notification_closed(self, notification: QWidget):
        if notification in self.notifications: self.notifications.remove(notification)
        self.reposition_all()
    def reposition_all(self):
        try: screen_geo = QApplication.primaryScreen().availableGeometry()
        except AttributeError: screen_geo = QApplication.screens()[0].availableGeometry()
        margin = 15
        total_height = 0
        for n in reversed(self.notifications):
            n.adjustSize()
            width, height = n.width(), n.height()
            x = screen_geo.right() - width - margin
            y = screen_geo.bottom() - height - margin - total_height
            n.move(x, y)
            total_height += height + 10

class AnimatedProgressBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._value = 0.0
        self._maximum = 14400.0
        self._animation = QPropertyAnimation(self, b"progress_value", self)
        self._animation.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self._alert_mode = False
        self._alert_visible = True
        self._alert_timer = QTimer(self)
        self._alert_timer.timeout.connect(self._toggle_alert_visibility)
    def _toggle_alert_visibility(self): self._alert_visible = not self._alert_visible; self.update()
    def get_progress_value(self): return self._value
    def set_progress_value(self, value): self._value = value; self.update()
    progress_value = pyqtProperty(float, fget=get_progress_value, fset=set_progress_value)
    def setValue(self, value):
        value = max(0, min(float(value), self._maximum))
        if self._animation.state() == QPropertyAnimation.State.Running: self._animation.stop()
        self._animation.setStartValue(self.progress_value)
        self._animation.setEndValue(value)
        self._animation.setDuration(1000)
        self._animation.start()
        is_limit_reached = value >= self._maximum
        if is_limit_reached and not self._alert_mode: self._alert_mode = True; self._alert_timer.start(400)
        elif not is_limit_reached and self._alert_mode: self._alert_mode = False; self._alert_timer.stop(); self._alert_visible = True; self.update()
    def setMaximum(self, value): self._maximum = float(value); self.update()
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        bg_rect = self.rect()
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(ColorPalette.SECONDARY))
        painter.drawRoundedRect(bg_rect, self.height() // 2, self.height() // 2)
        percent = (self.progress_value / self._maximum) if self._maximum > 0 else 0
        progress_width = percent * bg_rect.width()
        if self._alert_mode: color = QColor(ColorPalette.RED) if self._alert_visible else QColor(ColorPalette.SECONDARY)
        elif percent > 0.85: color = QColor(ColorPalette.RED)
        elif percent > 0.50: color = QColor(ColorPalette.AMBER)
        else: color = QColor(ColorPalette.PRIMARY)
        if progress_width > 0:
            progress_rect = QRect(0, 0, int(progress_width), self.height())
            painter.setBrush(color)
            painter.drawRoundedRect(progress_rect, self.height() // 2, self.height() // 2)

class ClickIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(30, 30)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide)
    def paintEvent(self, event): painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.setBrush(QColor(255, 0, 0, 150)); painter.setPen(Qt.PenStyle.NoPen); painter.drawEllipse(self.rect())
    def show_at(self, x, y): self.move(x - self.width() // 2, y - self.height() // 2); self.show(); self.timer.start(200)

class WindowManager(QObject):
    log_request = pyqtSignal(str, str)
    click_visual_request = pyqtSignal(int, int)

    def __init__(self):
        super().__init__()
        self.INPUT_MOUSE = 0
        self.MOUSEEVENTF_LEFTDOWN = 0x0002
        self.MOUSEEVENTF_LEFTUP = 0x0004
        self.MOUSEEVENTF_MOVE = 0x0001
        self.MOUSEEVENTF_ABSOLUTE = 0x8000
        class MOUSEINPUT(ctypes.Structure): _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))]
        class INPUT(ctypes.Structure): _fields_ = [("type", wintypes.DWORD), ("mi", MOUSEINPUT)]
        self.INPUT_STRUCT = INPUT
        self.screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    def _send_mouse_input(self, flags, x=0, y=0):
        if flags & self.MOUSEEVENTF_ABSOLUTE:
            x = (x * 65535) // self.screen_width
            y = (y * 65535) // self.screen_height
        mouse_input = self.INPUT_STRUCT(type=self.INPUT_MOUSE, mi=self.INPUT_STRUCT._fields_[1][1](dx=x, dy=y, mouseData=0, dwFlags=flags, time=0, dwExtraInfo=None))
        ctypes.windll.user32.SendInput(1, ctypes.byref(mouse_input), ctypes.sizeof(mouse_input))

    def _send_input_click(self, x: int, y: int):
        try:
            self._send_mouse_input(self.MOUSEEVENTF_MOVE | self.MOUSEEVENTF_ABSOLUTE, x, y)
            time.sleep(0.05)
            self._send_mouse_input(self.MOUSEEVENTF_LEFTDOWN | self.MOUSEEVENTF_ABSOLUTE, x, y)
            time.sleep(random.uniform(0.1, 0.2))
            self._send_mouse_input(self.MOUSEEVENTF_LEFTUP | self.MOUSEEVENTF_ABSOLUTE, x, y)
            logging.info(f"Выполнен клик (SendInput) по ({x},{y})")
            return True
        except Exception as e:
            logging.error(f"Ошибка клика (SendInput) по ({x},{y})", exc_info=True)
            self.log_request.emit(f"Клик SendInput не удался: {e}", "error")
            return False




    def find_template(self, template_path: str, confidence=0.8) -> Optional[Tuple[int, int]]:
        """Находит шаблон и возвращает его координаты, но не кликает."""
        if not CV2_AVAILABLE: return None
        try:
            template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
            if template is None:
                if os.path.exists(template_path): logging.error(f"Не удалось загрузить шаблон: {template_path}")
                return None

            screenshot = pyautogui.screenshot()
            screen_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

            res = cv2.matchTemplate(screen_cv, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)

            if max_val >= confidence:
                w, h = template.shape[::-1]
                return (max_loc[0] + w // 2, max_loc[1] + h // 2)
            else:
                logging.debug(f"Шаблон '{os.path.basename(template_path)}' не найден (уверенность: {max_val:.2f})")
                return None
        except Exception as e:
            logging.error(f"Ошибка при поиске шаблона '{template_path}'", exc_info=True)
            self.log_request.emit(f"Ошибка OpenCV: {e}", "error")
            return None

    def humanized_click(self, x: int, y: int, hwnd: int = None, log_prefix: str = ""):
        self.click_visual_request.emit(x, y)
        if self.robust_click(x, y, hwnd=hwnd, log_prefix=log_prefix):
            return
        logging.warning("robust_click не удался. Использую резервный метод (pyautogui).")
        try:
            if PYAUTOGUI_AVAILABLE:
                pyautogui.click(x, y)
                logging.info(f"{log_prefix} Выполнен резервный клик (pyautogui) по ({x},{y})")
            else:
                self.log_request.emit("Резервный метод клика (pyautogui) недоступен.", "error")
        except Exception as e:
            logging.error(f"{log_prefix} Резервный клик (pyautogui) также не удался: {e}", exc_info=True)

    def find_and_click_template(self, template_path: str, confidence=0.8, hwnd: int = None, log_prefix: str = "") -> bool:
        coords = self.find_template(template_path, confidence)
        if coords:
            self.humanized_click(coords[0], coords[1], hwnd=hwnd, log_prefix=log_prefix)
            return True
        return False

    def find_windows_by_config(self, config: dict, config_key: str, main_window_hwnd: int) -> List[int]:
        window_config = config.get(config_key, {})
        if not window_config: return []
        find_method = window_config.get("FIND_METHOD")
        EXCLUDED_TITLES = config.get("EXCLUDED_TITLES", [])
        EXCLUDED_PROCESSES = config.get("EXCLUDED_PROCESSES", [])
        arrange_minimized = config.get("ARRANGE_MINIMIZED_TABLES", False) and config_key == "TABLE"
        found_windows = []
        def enum_windows_callback(hwnd, _):
            if hwnd == main_window_hwnd or not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd) or (not arrange_minimized and win32gui.IsIconic(hwnd)): return True
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                win32api.CloseHandle(h_process)
                if process_name.lower() in EXCLUDED_PROCESSES: return True
                title = win32gui.GetWindowText(hwnd)
                if any(excluded.lower() in title.lower() for excluded in EXCLUDED_TITLES): return True
                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                if w == 0 or h == 0: return True
                match = False
                if find_method == "RATIO":
                    if h != 0 and (window_config["W"] / window_config["H"]) * (1 - window_config["TOLERANCE"]) <= (w / h) <= (window_config["W"] / window_config["H"]) * (1 + window_config["TOLERANCE"]): match = True
                elif find_method == "TITLE_AND_SIZE":
                    if window_config["TITLE"].lower() in title.lower() and abs(w - window_config["W"]) <= window_config["TOLERANCE"] and abs(h - window_config["H"]) <= window_config["TOLERANCE"]: match = True
                elif find_method == "TITLE":
                    if window_config["TITLE"].lower() in title.lower(): match = True
                if match: found_windows.append(hwnd)
            except Exception as e: logging.debug(f"Ошибка при перечислении окна {hwnd}: {e}")
            return True
        try: win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e: logging.error("Критическая ошибка в EnumWindows", exc_info=True)
        try: found_windows.sort(key=lambda hwnd: win32gui.GetWindowRect(hwnd)[0])
        except Exception as e: self.log_request.emit(f"Ошибка сортировки окон: {e}", "warning")
        return found_windows

    def find_first_window_by_title(self, text_in_title: str, exact_match: bool = False) -> Optional[int]:
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                try: title = win32gui.GetWindowText(hwnd)
                except Exception: return
                if (exact_match and text_in_title == title) or (not exact_match and text_in_title.lower() in title.lower()): hwnds.append(hwnd)
        try: win32gui.EnumWindows(callback, None)
        except Exception as e: logging.error("Критическая ошибка в EnumWindows (поиск по заголовку)", exc_info=True)
        return hwnds[0] if hwnds else None

    def is_process_running(self, process_name_to_find: str) -> bool:
        pids = win32process.EnumProcesses()
        for pid in pids:
            if pid == 0: continue
            try:
                h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                if h_process:
                    process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                    win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in process_name.lower(): return True
            except Exception: continue
        return False

    def find_first_window_by_process_name(self, process_name_to_find: str, check_visible: bool = True) -> Optional[int]:
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd):
                if check_visible and (not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd)):
                    return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                    win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in process_name.lower(): hwnds.append(hwnd)
                except Exception: pass
        try: win32gui.EnumWindows(callback, None)
        except Exception as e: logging.error("Критическая ошибка в EnumWindows (поиск по процессу)", exc_info=True)
        return hwnds[0] if hwnds else None

    def press_key(self, key_code: int):
        try: 
            win32api.keybd_event(key_code, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e: 
            logging.error("Ошибка эмуляции нажатия клавиши", exc_info=True)
            self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

    def close_window(self, hwnd: int):
        """Отправляет окну сообщение о закрытии."""
        try:
            if win32gui.IsWindow(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                logging.info(f"Отправлено сообщение WM_CLOSE окну {hwnd}")
                return True
        except Exception as e:
            logging.error(f"Не удалось отправить WM_CLOSE окну {hwnd}", exc_info=True)
        return False
    def robust_click(self, x: int, y: int, hwnd: int = None, delay: float = 0.35, log_prefix: str = "") -> bool:
        """
        Самый надёжный клик по экрану: активирует окно, ждёт, кликает мышью, логирует.
        Работает на всех Windows-программах, которые не защищены от automation на уровне драйверов.
        """
        try:
            if hwnd:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(delay)
                fg = win32gui.GetForegroundWindow()
                if fg != hwnd:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                    fg = win32gui.GetForegroundWindow()
                if fg != hwnd:
                    logging.error(f"{log_prefix} robust_click: Не удалось активировать окно для клика (hwnd={hwnd}, fg={fg})")
                    return False
            old_pos = win32api.GetCursorPos()
            ctypes.windll.user32.SetCursorPos(x, y)
            time.sleep(0.10)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            time.sleep(0.06)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            logging.info(f"{log_prefix} robust_click: Клик выполнен по ({x},{y}) с hwnd={hwnd}")
            ctypes.windll.user32.SetCursorPos(old_pos[0], old_pos[1])
            return True
        except Exception as e:
            logging.error(f"{log_prefix} robust_click: Ошибка при клике — {e}", exc_info=True)
            return False

class TelegramNotifier(QObject):
    def __init__(self, token, chat_id):
        super().__init__()
        self.token = token
        self.chat_id = chat_id
        self.api_url = f"https://api.telegram.org/bot{self.token}/sendMessage"
        self.message_queue = queue.Queue()
        self.worker_thread = threading.Thread(target=self._worker, daemon=True)
        self.worker_thread.start()

    def send_message(self, message: str):
        if not self.token or not self.chat_id: 
            logging.warning("Токен или ID чата Telegram не настроены.")
            return
        if len(message) > 4096:
            message = message[:4090] + "\n[...]"
        self.message_queue.put(message)

    def _worker(self):
        while True:
            message = self.message_queue.get()
            try: 
                requests.post(self.api_url, data={'chat_id': self.chat_id, 'text': message}, timeout=10).raise_for_status()
                logging.info("Сообщение в Telegram успешно отправлено.")
            except requests.RequestException as e: 
                logging.error(f"Не удалось отправить сообщение в Telegram: {e}")
            finally:
                self.message_queue.task_done()

class UpdateManager(QObject):
    log_request = pyqtSignal(str, str)
    check_finished = pyqtSignal()
    status_update = pyqtSignal(str)

    def __init__(self): super().__init__(); self.update_info = {}

    def is_new_version_available(self, current_v_str: str, latest_v_str: str) -> bool:
        try:
            current = [int(p) for p in current_v_str.split('-')[0].lstrip('v').split('.')]
            latest = [int(p) for p in latest_v_str.lstrip('v').split('.')]
            max_len = max(len(current), len(latest))
            current += [0] * (max_len - len(current))
            latest += [0] * (max_len - len(latest))
            return latest > current
        except Exception as e:
            logging.error(f"Ошибка сравнения версий: {e}.", exc_info=True)
            return latest_v_str > current_v_str

    def check_for_updates(self):
        self.status_update.emit(AppConfig.LOBBY_MSG_UPDATE_CHECK)
        try:
            api_url = f"https://api.github.com/repos/{AppConfig.GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            latest_release = response.json()
            if (latest_version := latest_release.get("tag_name")) and self.is_new_version_available(AppConfig.CURRENT_VERSION, latest_version):
                self.status_update.emit(f"Найдена версия {latest_version}...")
                self.log_request.emit(f"Доступна новая версия: {latest_version}. Обновление...", "info")
                self.update_info = latest_release
                threading.Thread(target=self.apply_update, daemon=True).start()
            else: self.log_request.emit("Вы используете последнюю версию.", "info"); self.check_finished.emit()
        except requests.RequestException as e: self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error"); self.check_finished.emit()
        except Exception as e: logging.error("Неожиданная ошибка при проверке обновлений", exc_info=True); self.check_finished.emit()

    def apply_update(self):
        download_url = next((asset["browser_download_url"] for asset in self.update_info.get("assets", []) if asset["name"] == AppConfig.ASSET_NAME), None)
        if not download_url: self.log_request.emit("Не удалось найти ZIP-архив в релизе.", "error"); return
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url: str):
        update_zip_name = "update.zip"
        try:
            self.status_update.emit("Скачивание обновления...")
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192): f.write(chunk)

            self.status_update.emit("Распаковка архива...")
            update_folder = "update_temp"
            if os.path.isdir(update_folder): import shutil; shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref: zip_ref.extractall(update_folder)

            self.status_update.emit("Перезапуск...")
            updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            exe_name = os.path.basename(current_exe_path)
            script_content = f'@echo off\nchcp 65001 > NUL\necho Waiting for OiHelper to close...\ntimeout /t 2 /nobreak > NUL\ntaskkill /pid {os.getpid()} /f > NUL\necho Waiting for process to terminate...\ntimeout /t 3 /nobreak > NUL\necho Moving new files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\nrd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\ndel "{current_dir}\\{update_zip_name}"\necho Starting new version...\nstart "" "{exe_name}"\n(goto) 2>nul & del "%~f0"'
            with open(updater_script_path, "w", encoding="cp866") as f: f.write(script_content)
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()
        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error")
            logging.error("Ошибка при обновлении", exc_info=True)
            if os.path.exists(update_zip_name):
                try: os.remove(update_zip_name)
                except OSError as err: logging.error(f"Не удалось удалить временный файл обновления: {err}")
            self.log_request.emit(AppConfig.MSG_UPDATE_FAIL, "warning")
            self.check_finished.emit()

class MainWindow(QMainWindow):
    def __init__(self, splash):
        super().__init__()
        self.splash = splash
        self.notification_manager = NotificationManager()
        self.window_manager = WindowManager()
        self.update_manager = UpdateManager()
        self.telegram_notifier = TelegramNotifier(AppConfig.TELEGRAM_BOT_TOKEN, AppConfig.TELEGRAM_CHAT_ID)
        self.click_indicator = ClickIndicator()

        self.current_project = None
        self.is_auto_record_enabled = True
        self.is_automation_enabled = True
        self.is_auto_popup_closing_enabled = False # New state for spam toggle
        self.last_table_count = 0
        self.recording_start_time = 0
        self.is_sending_logs = False
        self.player_was_open = False
        self.is_record_stopping = False
        self.is_launching_player = False
        self.is_launching_recorder = False
        self.chrome_was_visible = False 
        self.last_arrangement_time = 0 

        self.shell = win32com.client.Dispatch("WScript.Shell") if WIN32COM_AVAILABLE else None

        self.init_ui()
        self.connect_signals()
        self.init_timers()
        self.init_startup_checks()

    def closeEvent(self, event):
        logging.info("Приложение закрывается, остановка таймеров...")
        for timer in self.timers.values():
            timer.stop()
        super().closeEvent(event)

    def get_current_screen(self):
        try: return self.screen() or QApplication.primaryScreen()
        except Exception: return QApplication.primaryScreen()

    def log(self, message: str, message_type: str):
        self.notification_manager.show(message, message_type)
        if message_type == 'error':
            self.telegram_notifier.send_message(f"OiHelper Критическая ошибка: {message}")

    def _clear_layout(self):
        # ИСПРАВЛЕНО: Это самый безопасный способ полной перезагрузки UI.
        # Мы создаем новый, пустой центральный виджет и заменяем им старый.
        # Система родитель-потомок в Qt автоматически позаботится об удалении
        # старого виджета и всех его дочерних элементов (лэйаутов, кнопок и т.д.).
        old_widget = self.centralWidget()
        if old_widget is not None:
            old_widget.setParent(None)
            old_widget.deleteLater()

        # Устанавливаем новый пустой виджет
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)


    def init_ui(self):
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet(StyleSheet.MAIN_WINDOW)
        if os.path.exists(AppConfig.ICON_PATH): self.setWindowIcon(QIcon(AppConfig.ICON_PATH))
        self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

    def build_lobby_ui(self, message: str = AppConfig.LOBBY_MSG_SEARCHING):
        try:
            self._clear_layout()
            self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
            layout = QVBoxLayout(self.central_widget)
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

            self.lobby_status_label = QLabel(message)
            self.lobby_status_label.setStyleSheet(StyleSheet.LOBBY_LABEL)
            self.lobby_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

            manual_select_label = QLabel("или выберите вручную:")
            manual_select_label.setStyleSheet(StyleSheet.LOBBY_LABEL)
            manual_select_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

            self.project_combo = QComboBox()
            self.project_combo.setStyleSheet(StyleSheet.COMBO_BOX)
            self.project_combo.addItem("Выберите проект...")
            self.project_combo.addItems(PROJECT_CONFIGS.keys())
            self.project_combo.currentTextChanged.connect(self.on_project_selected_manually)

            layout.addWidget(self.lobby_status_label)
            layout.addSpacing(20)
            layout.addWidget(manual_select_label)
            layout.addWidget(self.project_combo)
            layout.addStretch()

            self.update_window_title()
        except Exception as e:
            logging.critical("Критическая ошибка в build_lobby_ui", exc_info=True)


    def on_project_selected_manually(self, project_name: str):
        if project_name in PROJECT_CONFIGS:
            self.log(f"Проект '{project_name}' выбран вручную.", "info")
            # ИСПРАВЛЕНО: Используем QTimer.singleShot для отложенного вызова,
            # чтобы избежать удаления виджета (QComboBox) во время обработки его сигнала.
            QTimer.singleShot(0, lambda: self.on_project_changed(project_name))

    def build_project_ui(self):
        if not self.current_project:
            return
        
        try:
            self._clear_layout()

            if self.current_project == "GG":
                self.setFixedSize(AppConfig.GG_UI_WIDTH, AppConfig.GG_UI_HEIGHT)
                main_layout = QHBoxLayout(self.central_widget)
                main_layout.setContentsMargins(12, 10, 12, 10)
                main_layout.setSpacing(15)

                # --- Toggles ---
                toggles_frame = QFrame()
                toggles_layout = QGridLayout(toggles_frame)
                toggles_layout.setContentsMargins(0,0,0,0)
                self.automation_toggle = ToggleSwitch()
                self.auto_record_toggle = ToggleSwitch()
                self.auto_popup_toggle = ToggleSwitch()
                automation_label = QLabel("Автоматика")
                automation_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                auto_record_label = QLabel("Автозапись")
                auto_record_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                auto_popup_label = QLabel("Автозакрытие спама (beta)")
                auto_popup_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                toggles_layout.addWidget(automation_label, 0, 0)
                toggles_layout.addWidget(self.automation_toggle, 0, 1)
                toggles_layout.addWidget(auto_record_label, 1, 0)
                toggles_layout.addWidget(self.auto_record_toggle, 1, 1)
                toggles_layout.addWidget(auto_popup_label, 2, 0)
                toggles_layout.addWidget(self.auto_popup_toggle, 2, 1)
                main_layout.addWidget(toggles_frame)

                # --- Separator ---
                separator = QFrame()
                separator.setFrameShape(QFrame.Shape.VLine)
                separator.setFrameShadow(QFrame.Shadow.Sunken)
                main_layout.addWidget(separator)

                # --- Buttons ---
                buttons_frame = QFrame()
                buttons_layout = QHBoxLayout(buttons_frame)
                buttons_layout.setContentsMargins(0,0,0,0)
                buttons_layout.setSpacing(8)
                self.arrange_tables_button = QPushButton(AppConfig.MSG_ARRANGE_TABLES)
                self.arrange_system_button = QPushButton(AppConfig.MSG_ARRANGE_SYSTEM)
                self.close_tables_button = QPushButton(AppConfig.MSG_CLICK_COMMAND)
                buttons_layout.addWidget(self.arrange_tables_button)
                buttons_layout.addWidget(self.arrange_system_button)
                buttons_layout.addWidget(self.close_tables_button)
                main_layout.addWidget(buttons_frame, 1) 

                # --- Progress Bar ---
                self.progress_frame = QFrame()
                progress_layout = QVBoxLayout(self.progress_frame)
                progress_layout.setContentsMargins(0,0,0,0)
                self.progress_bar = AnimatedProgressBar()
                self.progress_bar.setFixedHeight(6)
                self.progress_bar_label = QLabel()
                self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)
                progress_layout.addWidget(self.progress_bar)
                progress_layout.addWidget(self.progress_bar_label)
                main_layout.addWidget(self.progress_frame, 1)

            else: # Default for QQ
                self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
                main_layout = QVBoxLayout(self.central_widget)
                main_layout.setContentsMargins(12, 12, 12, 12)
                main_layout.setSpacing(10)

                # --- Toggles ---
                toggles_frame = QFrame()
                toggles_layout = QGridLayout(toggles_frame)
                toggles_layout.setContentsMargins(0,0,0,0)
                self.automation_toggle = ToggleSwitch()
                self.auto_record_toggle = ToggleSwitch()
                self.auto_popup_toggle = ToggleSwitch()
                automation_label = QLabel("Автоматика")
                automation_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                auto_record_label = QLabel("Автозапись")
                auto_record_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                auto_popup_label = QLabel("Автозакрытие спама (beta)")
                auto_popup_label.setStyleSheet(StyleSheet.STATUS_LABEL)
                toggles_layout.addWidget(automation_label, 0, 0)
                toggles_layout.addWidget(self.automation_toggle, 0, 1)
                toggles_layout.addWidget(auto_record_label, 1, 0)
                toggles_layout.addWidget(self.auto_record_toggle, 1, 1)
                toggles_layout.addWidget(auto_popup_label, 2, 0)
                toggles_layout.addWidget(self.auto_popup_toggle, 2, 1)
                toggles_layout.setColumnStretch(0, 1)
                main_layout.addWidget(toggles_frame)

                # --- Buttons ---
                buttons_frame = QFrame()
                buttons_layout = QHBoxLayout(buttons_frame)
                buttons_layout.setContentsMargins(0,0,0,0)
                buttons_layout.setSpacing(8)
                self.arrange_tables_button = QPushButton(AppConfig.MSG_ARRANGE_TABLES)
                self.arrange_system_button = QPushButton(AppConfig.MSG_ARRANGE_SYSTEM)
                buttons_layout.addWidget(self.arrange_tables_button)
                buttons_layout.addWidget(self.arrange_system_button)
                main_layout.addWidget(buttons_frame)
                main_layout.addStretch(1)

                # --- Progress Bar ---
                self.progress_frame = QFrame()
                progress_layout = QHBoxLayout(self.progress_frame)
                progress_layout.setContentsMargins(0,0,0,0)
                self.progress_bar = AnimatedProgressBar()
                self.progress_bar.setFixedHeight(6)
                self.progress_bar_label = QLabel()
                self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)
                progress_layout.addWidget(self.progress_bar, 1)
                progress_layout.addWidget(self.progress_bar_label)
                main_layout.addWidget(self.progress_frame)

            self.apply_styles_and_effects()
            self.update_project_ui_state()
            self.connect_project_signals()
        except Exception as e:
            logging.critical("Критическая ошибка в build_project_ui", exc_info=True)


    def apply_styles_and_effects(self):
        try:
            self.arrange_tables_button.setStyleSheet(StyleSheet.get_button_style(primary=True))
            self.arrange_system_button.setStyleSheet(StyleSheet.get_button_style(False))
            if hasattr(self, 'close_tables_button'): self.close_tables_button.setStyleSheet(StyleSheet.get_button_style(False))

            buttons_to_style = [self.arrange_tables_button, self.arrange_system_button]
            if hasattr(self, 'close_tables_button'):
                buttons_to_style.append(self.close_tables_button)

            for btn in buttons_to_style:
                if btn:
                    # ИСПРАВЛЕНО: Эффект создается без родителя, чтобы избежать двойного владения.
                    shadow = QGraphicsDropShadowEffect()
                    shadow.setBlurRadius(15)
                    shadow.setColor(QColor(0, 0, 0, 40))
                    shadow.setOffset(0, 2)
                    btn.setGraphicsEffect(shadow)
        except Exception as e:
            logging.critical("Критическая ошибка в apply_styles_and_effects", exc_info=True)


    def update_window_title(self):
        title = AppConfig.APP_TITLE_TEMPLATE.format(version=AppConfig.CURRENT_VERSION)
        if self.current_project: title += f" ({self.current_project})"
        self.setWindowTitle(title)

    def update_project_ui_state(self):
        if not self.current_project: return
        self.automation_toggle.setChecked(self.is_automation_enabled)
        self.auto_record_toggle.setChecked(self.is_auto_record_enabled)
        self.auto_popup_toggle.setChecked(self.is_auto_popup_closing_enabled)
        self.progress_frame.setVisible(self.recording_start_time > 0)
        is_project_active = self.current_project is not None
        self.arrange_tables_button.setEnabled(is_project_active)
        self.arrange_system_button.setEnabled(is_project_active)
        if hasattr(self, 'close_tables_button'): self.close_tables_button.setEnabled(is_project_active)

    def connect_signals(self):
        self.window_manager.log_request.connect(self.log)
        self.window_manager.click_visual_request.connect(self.click_indicator.show_at)
        self.update_manager.log_request.connect(self.log)
        self.update_manager.status_update.connect(self.splash.update_status)
        self.update_manager.check_finished.connect(self.show_main_window_and_start_logic)

    def connect_project_signals(self):
        if not self.current_project: return
        try:
            self.automation_toggle.clicked.disconnect()
            self.auto_record_toggle.clicked.disconnect()
            self.auto_popup_toggle.clicked.disconnect()
            self.arrange_tables_button.clicked.disconnect()
            self.arrange_system_button.clicked.disconnect()
            if hasattr(self, 'close_tables_button'):
                self.close_tables_button.clicked.disconnect()
        except TypeError:
            pass

        self.automation_toggle.clicked.connect(self.toggle_automation)
        self.auto_record_toggle.clicked.connect(self.toggle_auto_record)
        self.auto_popup_toggle.clicked.connect(self.toggle_auto_popup_closing)
        self.arrange_tables_button.clicked.connect(self.arrange_tables)
        self.arrange_system_button.clicked.connect(self.arrange_other_windows)
        if hasattr(self, 'close_tables_button'): self.close_tables_button.clicked.connect(self.close_all_tables)

    def init_timers(self):
        self.timers = { 
            "player_check": QTimer(self), "recorder_check": QTimer(self), 
            "auto_record": QTimer(self), "auto_arrange": QTimer(self), 
            "session": QTimer(self), "player_start": QTimer(self), 
            "record_cooldown": QTimer(self), "popup_check": QTimer(self)
        }
        self.timers["player_check"].timeout.connect(self.check_for_player)
        self.timers["recorder_check"].timeout.connect(self.check_for_recorder)
        self.timers["auto_record"].timeout.connect(self.check_auto_record_logic)
        self.timers["auto_arrange"].timeout.connect(self.check_for_new_tables)
        self.timers["session"].timeout.connect(self.update_session_progress)
        self.timers["player_start"].timeout.connect(self.attempt_player_start_click)
        self.timers["record_cooldown"].setSingleShot(True)
        self.timers["record_cooldown"].timeout.connect(lambda: setattr(self, 'is_record_stopping', False))
        self.timers["popup_check"].timeout.connect(self.check_for_popups)

    def init_startup_checks(self):
        self.update_window_title()
        threading.Thread(target=self.update_manager.check_for_updates, daemon=True).start()

    def show_main_window_and_start_logic(self):
        self.splash.close()
        self.build_lobby_ui()
        self.show()
        self.position_window_default()
        self.start_main_logic()

    def start_main_logic(self):
        if hasattr(self, 'lobby_status_label'): self.lobby_status_label.setText(AppConfig.LOBBY_MSG_SEARCHING)
        self.check_for_player()
        self.check_for_recorder()
        self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        QTimer.singleShot(1000, self.initial_recorder_sync_check)
        if AppConfig.TELEGRAM_REPORT_LEVEL == 'all': self.telegram_notifier.send_message(f"OiHelper {AppConfig.CURRENT_VERSION} запущен.")
        self.check_system_uptime()
        self.check_admin_rights()

    def on_project_changed(self, new_project_name: Optional[str]):
        if self.current_project == new_project_name:
            return

        # ИСПРАВЛЕНО: Очищаем ссылки на виджеты, специфичные для старого проекта,
        # ПЕРЕД тем, как Qt удалит сами объекты. Это предотвращает "висячие" указатели.
        if self.current_project == "GG" and hasattr(self, 'close_tables_button'):
            del self.close_tables_button

        self.current_project = new_project_name
        self.last_table_count = 0
        self.timers["auto_arrange"].stop()
        self.timers["popup_check"].stop()
        self.update_window_title()

        if new_project_name:
            self.build_project_ui() 
            self.timers["auto_arrange"].start(AppConfig.AUTO_ARRANGE_INTERVAL)

            if new_project_name == "QQ":
                self.position_window_top_right()
                self.timers["popup_check"].start(AppConfig.POPUP_CHECK_INTERVAL_FAST)
            elif new_project_name == "GG":
                self.position_gg_panel()
                self.minimize_injector_window()
                self.timers["popup_check"].start(AppConfig.POPUP_CHECK_INTERVAL_FAST) 

            if self.is_automation_enabled:
                if new_project_name == "QQ": self.check_and_launch_opencv_server()
                self.arrange_other_windows()
        else:
            self.build_lobby_ui(AppConfig.LOBBY_MSG_SEARCHING)
            self.position_window_default()

    def check_for_player(self):
        if self.is_sending_logs: return
        player_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if not player_hwnd:
            if self.current_project is not None: self.on_project_changed(None)
            if hasattr(self, 'lobby_status_label'): self.lobby_status_label.setText(AppConfig.LOBBY_MSG_PLAYER_NOT_FOUND)
            if self.player_was_open: self.player_was_open = False; self.handle_player_close()
            else:
                if self.is_automation_enabled: self.check_and_launch_player()
        else:
            self.player_was_open = True
            project_name = None
            try:
                title = win32gui.GetWindowText(player_hwnd)
                project_name = next((short for full, short in {"QQPoker": "QQ", "ClubGG": "GG"}.items() if f"[{full}]" in title), None)
            except Exception: project_name = None
            if project_name:
                self.timers["player_start"].stop()
                if self.timers["player_start"].isActive(): self.log("Авто-старт плеера успешно завершен.", "info")
                self.on_project_changed(project_name)
            else:
                if self.current_project is not None: self.on_project_changed(None)
                if hasattr(self, 'lobby_status_label'): self.lobby_status_label.setText("Ожидание старта в лаунчере...")
                if self.is_automation_enabled and not self.timers["player_start"].isActive():
                    self.log("Лаунчер найден. Включаю попытки авто-старта...", "info")
                    self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
                    self.attempt_player_start_click()
        if not self.timers["player_check"].isActive(): self.timers["player_check"].start(AppConfig.PLAYER_CHECK_INTERVAL)

    def check_admin_rights(self):
        try:
            if not ctypes.windll.shell32.IsUserAnAdmin(): self.log(AppConfig.MSG_ADMIN_WARNING, "warning")
        except Exception as e: logging.error("Не удалось проверить права администратора", exc_info=True)

    def check_system_uptime(self):
        try:
            if (ctypes.windll.kernel32.GetTickCount64() / (1000 * 60 * 60 * 24)) > 5: self.log(AppConfig.MSG_UPTIME_WARNING, "warning")
        except Exception as e: logging.error("Не удалось проверить время работы системы", exc_info=True)

    def handle_player_close(self):
        if not self.is_automation_enabled: return
        self.is_sending_logs = True
        self.log("Плеер закрыт. Запускаю отправку логов...", "info")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        launched_count = 0
        try: shortcuts_to_launch = [(k, os.path.join(desktop_path, f)) for f in os.listdir(desktop_path) if f.lower().endswith('.lnk') for k in AppConfig.LOG_SENDER_KEYWORDS if k in f.lower()]
        except Exception as e: self.log(f"Ошибка при поиске скриптов логов: {e}", "error"); shortcuts_to_launch = []
        if not shortcuts_to_launch: self.log("Скрипты для отправки логов не найдены.", "warning")
        else:
            for keyword, path in shortcuts_to_launch:
                if self.window_manager.find_first_window_by_title(keyword): self.log(f"Скрипт '{keyword}' уже запущен.", "info")
                else:
                    try: os.startfile(path); launched_count += 1; self.log(f"Запущен скрипт отправки логов '{keyword}'.", "info")
                    except Exception as e: self.log(f"Не удалось запустить ярлык для '{keyword}': {e}", "error")
        if launched_count > 0: self.log(f"Всего запущено скриптов: {launched_count}", "info")
        self.log(f"Перезапуск плеера через {AppConfig.PLAYER_RELAUNCH_DELAY_S} секунд...", "info")
        QTimer.singleShot(AppConfig.PLAYER_RELAUNCH_DELAY_S * 1000, lambda: self.wait_for_logs_to_finish(time.monotonic()))

    def wait_for_logs_to_finish(self, start_time: float):
        if (time.monotonic() - start_time) > AppConfig.LOG_SENDER_TIMEOUT_S: self.log("Тайм-аут ожидания отправки логов. Возобновление работы.", "error"); self.is_sending_logs = False; self.check_for_player(); return
        if any(self.window_manager.find_first_window_by_title(k) for k in AppConfig.LOG_SENDER_KEYWORDS): self.log("Ожидание завершения отправки логов...", "info"); QTimer.singleShot(3000, lambda: self.wait_for_logs_to_finish(start_time))
        else: self.log("Отправка логов завершена.", "info"); self.is_sending_logs = False; self.check_for_player()

    def check_and_launch_player(self):
        if self.is_launching_player or self.is_sending_logs: return
        if self.window_manager.find_first_window_by_title(AppConfig.PLAYER_GAME_LAUNCHER_TITLE) or self.window_manager.find_first_window_by_title("launch"): self.log("Обнаружен процесс запуска/обновления плеера. Ожидание...", "info"); return
        self.log("Плеер не найден, ищу ярлык 'launch'...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'launch' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: 
                self.log("Найден ярлык плеера. Запускаю...", "info")
                os.startfile(shortcut_path)
                self.is_launching_player = True
                QTimer.singleShot(10000, lambda: setattr(self, 'is_launching_player', False))
                self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
            else: self.log("Ярлык плеера 'launch' на рабочем столе не найден.", "error")
        except Exception as e: self.log(f"Ошибка при поиске/запуске ярлыка: {e}", "error")

    def attempt_player_start_click(self):
        if not self.is_automation_enabled: self.timers["player_start"].stop(); return
        launcher_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if not launcher_hwnd: self.log("Не удалось найти окно лаунчера для авто-старта.", "warning"); self.timers["player_start"].stop(); return
        try: 
            self.focus_window(launcher_hwnd)
            rect = win32gui.GetWindowRect(launcher_hwnd)
            x, y = rect[0] + 50, rect[1] + 50
            self.log("Попытка авто-старта плеера...", "info")
            self.window_manager.robust_click(x, y, hwnd=launcher_hwnd, log_prefix="PlayerStart")
        except Exception as e: self.log(f"Не удалось активировать окно плеера: {e}", "error")

    def check_and_launch_opencv_server(self):
        if not self.is_automation_enabled: return
        config = PROJECT_CONFIGS.get("QQ")
        if not config or self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId()): return
        self.log("Сервер OpenCV не найден, ищу ярлык...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'opencv' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: self.log("Найден ярлык OpenCV. Запускаю...", "info"); os.startfile(shortcut_path); QTimer.singleShot(4000, self.arrange_other_windows)
            else: self.log("Ярлык для OpenCV на рабочем столе не найден.", "error")
        except Exception as e: self.log(f"Ошибка при поиске/запуске ярлыка OpenCV: {e}", "error")

    def check_for_recorder(self):
        if self.is_launching_recorder or self.window_manager.is_process_running(AppConfig.CAMTASIA_PROCESS_NAME):
            if self.timers["recorder_check"].isActive(): self.timers["recorder_check"].stop()
            return
        if self.is_automation_enabled:
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            try:
                shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.CAMTASIA_PROCESS_NAME in f.lower() and f.lower().endswith('.lnk')), None)
                if shortcut_path: 
                    self.log("Camtasia не найдена. Запускаю...", "warning")
                    os.startfile(shortcut_path)
                    self.is_launching_recorder = True
                    QTimer.singleShot(10000, lambda: setattr(self, 'is_launching_recorder', False))
                    self.timers["recorder_check"].stop()
                    QTimer.singleShot(3000, self.check_for_recorder)
                else: self.log("Ярлык для Camtasia на рабочем столе не найден.", "error")
            except Exception as e: self.log(f"Ошибка при поиске/запуске Camtasia: {e}", "error")
        if not self.timers["recorder_check"].isActive(): self.timers["recorder_check"].start(AppConfig.RECORDER_CHECK_INTERVAL)

    def initial_recorder_sync_check(self):
        if self.window_manager.find_first_window_by_title("Recording..."): self.log("Обнаружена активная запись. Перезапускаю для синхронизации...", "warning"); self.stop_recording_session(); QTimer.singleShot(2000, self.check_auto_record_logic)

    def toggle_auto_record(self):
        self.is_auto_record_enabled = not self.is_auto_record_enabled
        self.log(f"Автозапись {'включена' if self.is_auto_record_enabled else 'выключена'}.", "info")
        if self.is_auto_record_enabled: self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        else: self.timers["auto_record"].stop()
        self.auto_record_toggle.setChecked(self.is_auto_record_enabled)

    def toggle_automation(self):
        self.is_automation_enabled = not self.is_automation_enabled
        self.log(f"Автоматика {'включена' if self.is_automation_enabled else 'выключена'}.", "info")
        if self.is_automation_enabled: self.check_for_player(); self.check_for_recorder()
        self.automation_toggle.setChecked(self.is_automation_enabled)

    def toggle_auto_popup_closing(self):
        self.is_auto_popup_closing_enabled = not self.is_auto_popup_closing_enabled
        self.log(f"Автозакрытие спама (beta) {'включено' if self.is_auto_popup_closing_enabled else 'выключено'}.", "info")
        self.auto_popup_toggle.setChecked(self.is_auto_popup_closing_enabled)

    def check_auto_record_logic(self):
        if not self.is_auto_record_enabled or not self.current_project or self.is_record_stopping:
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            return

        try:
            should_be_recording = False
            trigger_reason = ""

            if self.current_project == "GG":
                # Проверяем clubgg.exe только среди видимых и не свёрнутых окон!
                is_clubgg_visible = False
                all_hwnds = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
                for hwnd in all_hwnds:
                    try:
                        if win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd):
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            h_process = win32api.OpenProcess(
                                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                            win32api.CloseHandle(h_process)
                            if process_name.lower() == "clubgg.exe":
                                is_clubgg_visible = True
                                break
                    except Exception:
                        continue

                is_chrome_visible = self.window_manager.find_first_window_by_process_name("chrome.exe", check_visible=True) is not None

                if is_clubgg_visible:
                    should_be_recording = True
                    trigger_reason = "ClubGG (видимое окно)"
                elif is_chrome_visible:
                    should_be_recording = True
                    trigger_reason = "активное окно Chrome"
            else:
                # Для других проектов (QQ)
                if self.window_manager.find_windows_by_config(config, "LOBBY", self.winId()) or self.window_manager.find_windows_by_config(config, "TABLE", self.winId()):
                    should_be_recording = True
                    trigger_reason = f"активность {self.current_project}"

            is_recording = self.window_manager.find_first_window_by_title("Recording...") is not None
            is_paused = self.window_manager.find_first_window_by_title("Paused...") is not None

            # --- Дебаунс остановки записи ---
            if not hasattr(self, "last_should_record_false_time"):
                self.last_should_record_false_time = None

            if should_be_recording:
                self.last_should_record_false_time = None
                if not (is_recording or is_paused):
                    self.log(f"Начинаю автозапись (обнаружен {trigger_reason})...", "info")
                    self.start_recording_session()
                elif is_paused:
                    self.log("Запись на паузе. Возобновляю...", "warning")
                    self.perform_camtasia_action("resume")
            else:
                now = time.monotonic()
                if self.last_should_record_false_time is None:
                    self.last_should_record_false_time = now
                if (now - self.last_should_record_false_time) < 300:
                    # Ждём 5 минут без нужных окон
                    pass
                else:
                    if is_recording or is_paused:
                        self.log("Нет активности >5 минут. Останавливаю запись...", "info")
                        self.stop_recording_session()

            if self.current_project == "GG":
                # Просто фиксация хрома для логики, не влияет на остановку!
                self.chrome_was_visible = is_chrome_visible

        except Exception as e:
            self.log(f"Ошибка в логике автозаписи: {e}", "error")
            logging.error("Ошибка в check_auto_record_logic", exc_info=True)


    def start_recording_session(self):
        self.perform_camtasia_action("start")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        self.recording_start_time = time.monotonic()
        self.timers["session"].start(1000)
        self.progress_bar.setMaximum(config["SESSION_MAX_DURATION_S"])
        self.progress_bar.setValue(0)
        self.update_project_ui_state()

    def stop_recording_session(self):
        self.perform_camtasia_action("stop")
        self.timers["session"].stop()
        self.recording_start_time = 0
        self.is_record_stopping = True
        self.timers["record_cooldown"].start(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000)
        self.update_project_ui_state()

    def perform_camtasia_action(self, action: str):
        self.log(f"Выполняю команду '{action}' для Camtasia...", "info")
        self.position_recorder_window()  # Обязательно выводим окно на передний план!

        opencv_attempted = False
        success = False

        if CV2_AVAILABLE:
            for attempt in range(3):  # 3 попытки поиска + клика
                self.focus_camtasia_window()
                time.sleep(0.5)  # Дать окну стабильно сфокусироваться

                if action == "start":
                    # Попробуй fullscreen (иногда скрывает панель)
                    if self.window_manager.find_and_click_template('templates/camtasia_fullscreen.png'):
                        time.sleep(0.3)
                    # Новый этап: ищем координаты Rec и делаем robust_click
                    coords = self.window_manager.find_template('templates/camtasia_rec.png')
                    if coords:
                        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME, check_visible=False)
                        success = self.window_manager.robust_click(coords[0], coords[1], hwnd=recorder_hwnd, log_prefix="Camtasia-Rec")
                        if success:
                            self.log("Запуск записи через robust_click по Rec (OpenCV) выполнен.", "info")
                            break
                    # Старый способ (оставь на всякий случай)
                    if self.window_manager.find_and_click_template('templates/camtasia_rec.png'):
                        self.log("Запуск записи через кнопку Rec (OpenCV) выполнен.", "info")
                        success = True
                        break

                    if action == "stop":
                        coords = self.window_manager.find_template('templates/camtasia_stop.png')
                        if coords:
                            recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME, check_visible=False)
                            success = self.window_manager.robust_click(coords[0], coords[1], hwnd=recorder_hwnd, log_prefix="Camtasia-Stop")
                            if success:
                                self.log("Остановка записи через robust_click по Stop (OpenCV) выполнена.", "info")
                                break
                        if self.window_manager.find_and_click_template('templates/camtasia_stop.png'):
                            self.log("Остановка записи через кнопку Stop (OpenCV) выполнена.", "info")
                            success = True
                            break

                    elif action == "resume":
                        coords = self.window_manager.find_template('templates/camtasia_resume.png')
                        if coords:
                            recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME, check_visible=False)
                            success = self.window_manager.robust_click(coords[0], coords[1], hwnd=recorder_hwnd, log_prefix="Camtasia-Resume")
                            if success:
                                self.log("Возобновление записи через robust_click по Resume (OpenCV) выполнено.", "info")
                                break
                        if self.window_manager.find_and_click_template('templates/camtasia_resume.png'):
                            self.log("Возобновление записи через кнопку Resume (OpenCV) выполнено.", "info")
                            success = True
                            break



                time.sleep(0.7)  # Пауза между попытками
            opencv_attempted = True

        # Если OpenCV не справился — fallback на клавиатуру
        if not success:
            logging.warning("OpenCV не смог найти кнопку, fallback на hotkey.")
            key_map = {"start": win32con.VK_F9, "stop": win32con.VK_F10, "resume": win32con.VK_F9}
            key_code = key_map.get(action)
            if key_code:
                self.focus_camtasia_window()
                time.sleep(0.3)
                self.window_manager.press_key(key_code)
                QTimer.singleShot(600, self.position_recorder_window)  # Слегка задержать, Camtasia иногда тупит
                self.log(f"Горячая клавиша для {action} Camtasia отправлена.", "info")
            else:
                self.log(f"Действие '{action}' не поддерживается!", "error")

    def focus_camtasia_window(self):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME, check_visible=False)
        if recorder_hwnd:
            if win32gui.IsIconic(recorder_hwnd):
                win32gui.ShowWindow(recorder_hwnd, win32con.SW_RESTORE)
            self.focus_window(recorder_hwnd)
        else:
            self.log("Окно Camtasia не найдено для фокуса.", "warning")


    def update_session_progress(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "SESSION_MAX_DURATION_S" not in config: return
        elapsed = time.monotonic() - self.recording_start_time
        self.progress_bar.setValue(elapsed)
        if elapsed >= config["SESSION_MAX_DURATION_S"]:
            self.progress_bar_label.setText(AppConfig.MSG_LIMIT_REACHED)
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL + f" color: {ColorPalette.RED}; font-weight: bold;")
            QTimer.singleShot(100, self.handle_session_limit_reached)
        else:
            remaining_s = max(0, config["SESSION_MAX_DURATION_S"] - elapsed)
            self.progress_bar_label.setText(AppConfig.MSG_PROGRESS_LABEL.format(time.strftime('%H:%M:%S', time.gmtime(remaining_s))))
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)

    def handle_session_limit_reached(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        self.log(f"{config['SESSION_MAX_DURATION_S']/3600:.0f} часа записи истекли. Перезапуск...", "info")
        self.stop_recording_session()
        QTimer.singleShot(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000 + 1000, self.check_auto_record_logic)

    def check_for_new_tables(self):
        if not self.is_automation_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        if self.current_project == "GG" and not self.window_manager.is_process_running("ClubGG.exe"): self.last_table_count = 0; return
        current_tables = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        current_count = len(current_tables)
        if current_count != self.last_table_count: self.log(f"Изменилось количество столов: {self.last_table_count} -> {current_count}. Перерасстановка...", "info"); QTimer.singleShot(500, self.arrange_tables)
        self.last_table_count = current_count

    def arrange_tables(self):
        if not self.current_project: self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error"); return

        self.last_arrangement_time = time.monotonic()
        if self.current_project in ["QQ", "GG"] and self.timers["popup_check"].interval() == AppConfig.POPUP_CHECK_INTERVAL_SLOW:
            self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_FAST)
            self.log("Ускорен поиск попапов после расстановки.", "info")

        if self.current_project == "GG":
            self.arrange_gg_tables()
            return

        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "TABLE" not in config: self.log(f"Нет конфига столов для {self.current_project}.", "warning"); return
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        if found_windows:
            titles = [win32gui.GetWindowText(hwnd) for hwnd in found_windows if win32gui.IsWindow(hwnd)]
            logging.info(f"Найдены окна для расстановки ({len(titles)}): {titles}")
        else: self.log(AppConfig.MSG_ARRANGE_TABLES_NOT_FOUND, "warning"); return
        if self.current_project == "QQ" and len(found_windows) > 4: self.arrange_dynamic_qq_tables(found_windows, config); return
        slots_key = "TABLE_SLOTS_5" if self.current_project == "QQ" and len(found_windows) >= 5 else "TABLE_SLOTS"
        SLOTS = config[slots_key]
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS) or not win32gui.IsWindow(hwnd): continue
            x, y = SLOTS[i]
            _ = win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) if win32gui.IsIconic(hwnd) else None
            try:
                if self.current_project == "GG": win32gui.MoveWindow(hwnd, x, y, config["TABLE"]["W"], config["TABLE"]["H"], True)
                else: rect = win32gui.GetWindowRect(hwnd); win32gui.MoveWindow(hwnd, x, y, rect[2] - rect[0], rect[3] - rect[1], True)
                arranged_count += 1
            except Exception as e: self.log(f"Не удалось разместить стол {i+1}: {e}", "error")
        if arranged_count > 0: self.log(f"Расставлено столов: {arranged_count}", "info")

    def arrange_gg_tables(self):
        """
        Классическая расстановка до 4-х окон clubgg.exe по фиксированным слотам TABLE_SLOTS (без уменьшения, как было раньше).
        """
        config = PROJECT_CONFIGS.get("GG")
        if not config:
            self.log("GG конфиг не найден!", "error")
            return

        # Находим clubgg.exe окна (без учета регистра)
        all_hwnds = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        gg_tables = []
        for hwnd in all_hwnds:
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                h_process = win32api.OpenProcess(
                    win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                win32api.CloseHandle(h_process)
                if process_name.lower() == "clubgg.exe":
                    gg_tables.append(hwnd)
            except Exception as e:
                logging.debug(f"Ошибка фильтрации clubgg.exe hwnd={hwnd}: {e}")
        found_windows = gg_tables[:4]  # максимум 4 окна

        slots = config.get("TABLE_SLOTS", [])
        base_w, base_h = config["TABLE"]["W"], config["TABLE"]["H"]
        arranged_count = 0

        for i, hwnd in enumerate(found_windows):
            if i >= len(slots):
                break
            x, y = slots[i]
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.MoveWindow(hwnd, x, y, base_w, base_h, True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось разместить GG стол {i+1}: {e}", "error")

        if arranged_count > 0:
            self.log(f"GG столы расставлены по фиксированным слотам ({arranged_count}).", "info")

        # Лобби (если нужно — по config["LOBBY"])
        lobbies = self.window_manager.find_windows_by_config(config, "LOBBY", self.winId())
        cfg = config.get("LOBBY", {})
        if lobbies and cfg:
            try:
                win32gui.MoveWindow(lobbies[0], cfg["X"], cfg["Y"], cfg["W"], cfg["H"], True)
                self.log("Лобби GG размещено по координатам из конфига.", "info")
            except Exception as e:
                self.log(f"Не удалось разместить лобби GG: {e}", "error")




    def arrange_dynamic_qq_tables(self, found_windows, config):
        max_tables = len(found_windows)
        screen_geo = self.get_current_screen().availableGeometry()
        base_width, base_height = config["TABLE"]["W"], config["TABLE"]["H"]
        tables_per_row = min(max_tables, 5)
        rows = (max_tables + tables_per_row - 1) // tables_per_row
        new_width = max(int(base_width * 0.6), screen_geo.width() // tables_per_row)
        new_height = max(int(base_height * 0.6), screen_geo.height() // rows)
        overlap_x = 0.93 if tables_per_row > 1 else 1.0
        overlap_y = 0.93 if rows > 1 else 1.0
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            row = i // tables_per_row
            col = i % tables_per_row
            x = screen_geo.left() + int(col * new_width * overlap_x)
            y = screen_geo.top() + int(row * new_height * overlap_y)
            if not win32gui.IsWindow(hwnd): continue
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try: win32gui.MoveWindow(hwnd, x, y, new_width, new_height, True); arranged_count += 1
            except Exception as e: self.log(f"Не удалось динамически разместить стол {i+1}: {e}", "error")
        if arranged_count > 0: self.log(f"Динамически расставлено столов: {arranged_count}", "info")

    def arrange_other_windows(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            self.log(f"Нет конфига для {self.current_project}.", "warning")
            return

        # Позиционируем плеер
        self.position_player_window(config)

        # --- ДОБАВЬ ВОТ ЭТУ СТРОЧКУ ---
        if self.current_project == "GG":
            self.minimize_injector_window()

        self.position_recorder_window()
        if self.current_project == "GG":
            self.position_lobby_window(config)
        elif self.current_project == "QQ":
            self.position_cv_server_window(config)
        self.log("Системные окна расставлены.", "info")




    def position_window(self, hwnd: Optional[int], x: int, y: int, w: int, h: int, log_fail: str):
        if hwnd and win32gui.IsWindow(hwnd):
            try: win32gui.MoveWindow(hwnd, x, y, w, h, True)
            except Exception as e: logging.error(f"Ошибка позиционирования окна: {log_fail}", exc_info=True)
        else: self.log(log_fail, "warning")

    def position_player_window(self, config): player_config = config.get("PLAYER", {});_ = self.position_window(self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE), player_config["X"], player_config["Y"], player_config["W"], player_config["H"], "Плеер не найден.") if player_config else None
    def position_lobby_window(self, config): lobbies = self.window_manager.find_windows_by_config(config, "LOBBY", self.winId()); cfg = config.get("LOBBY", {});_ = self.position_window(lobbies[0] if lobbies else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "Лобби не найдено.") if cfg else None
    def position_cv_server_window(self, config): cv_windows = self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId()); cfg = config.get("CV_SERVER", {});_ = self.position_window(cv_windows[0] if cv_windows else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "CV Сервер не найден.") if cfg else None

    def position_recorder_window(self):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME)
        if recorder_hwnd and win32gui.IsWindow(recorder_hwnd):
            try:
                if win32gui.IsIconic(recorder_hwnd):
                    win32gui.ShowWindow(recorder_hwnd, win32con.SW_RESTORE)
                    time.sleep(0.2)
                screen_rect = self.get_current_screen().availableGeometry()
                rect = win32gui.GetWindowRect(recorder_hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]

                # Центр по горизонтали, по нижнему краю
                x = screen_rect.left() + (screen_rect.width() - w) // 2
                y = screen_rect.bottom() - h

                win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
            except Exception as e:
                self.log(f"Ошибка позиционирования Camtasia: {e}", "error")


    def minimize_injector_window(self):
        injector_hwnd = self.window_manager.find_first_window_by_title(AppConfig.INJECTOR_WINDOW_TITLE, exact_match=True)
        if injector_hwnd and win32gui.IsWindow(injector_hwnd):
            try: 
                win32gui.ShowWindow(injector_hwnd, win32con.SW_MINIMIZE)
                self.log("Окно 'injector' свернуто.", "info")
            except Exception as e: 
                self.log(f"Не удалось свернуть окно 'injector': {e}", "error")
        else:
            logging.warning("Окно 'injector' для сворачивания не найдено.")

    def focus_window(self, hwnd: int):
        try:
            if self.shell: self.shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception as e: logging.error(f"Не удалось сфокусировать окно {hwnd}", exc_info=True)

    def close_all_tables(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return

        self.log(f"Закрываю все столы для проекта {self.current_project}...", "info")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            return

        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        if not found_windows:
            self.log("Столы для закрытия не найдены.", "warning")
            return

        closed_count = 0
        for hwnd in found_windows:
            if self.window_manager.close_window(hwnd):
                closed_count += 1

        if closed_count > 0:
            self.log(f"Закрыто столов: {closed_count}", "info")


    def check_for_popups(self):
        if not self.is_automation_enabled or not self.current_project or not self.is_auto_popup_closing_enabled:
            return

        if self.last_arrangement_time > 0:
            elapsed = time.monotonic() - self.last_arrangement_time
            current_interval = self.timers["popup_check"].interval()
            if elapsed > AppConfig.POPUP_FAST_SCAN_DURATION_S:
                if current_interval != AppConfig.POPUP_CHECK_INTERVAL_SLOW:
                    self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_SLOW)
                    self.log(f"Замедлен поиск попапов для {self.current_project}.", "info")
            elif current_interval != AppConfig.POPUP_CHECK_INTERVAL_FAST:
                self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_FAST)

        if self.current_project == "QQ":
            self._handle_qq_popups()
        elif self.current_project == "GG":
            self._handle_gg_popups()

    def _handle_qq_popups(self):
        config = PROJECT_CONFIGS.get("QQ", {})
        popup_config = config.get("POPUPS")
        if not popup_config: return

        templates_dir = self._get_templates_dir("qq")
        if not templates_dir: return

        wm = self.window_manager
        confidence = popup_config.get("CONFIDENCE", 0.83)

        for rule in popup_config.get("SPAM", []):
            poster_path = os.path.join(templates_dir, rule["poster"])
            if os.path.exists(poster_path) and wm.find_template(poster_path, confidence=confidence):
                self.log(f"Обнаружен спам '{rule['poster']}'. Ищу кнопку '{rule['close_button']}'...", "info")
                time.sleep(0.5)
                close_btn_path = os.path.join(templates_dir, rule['close_button'])
                if wm.find_and_click_template(close_btn_path, confidence=confidence):
                    self.log(f"Успешно закрыт спам с помощью '{rule['close_button']}'.", "info")
                else:
                    self.log(f"Не удалось найти кнопку закрытия '{rule['close_button']}'.", "warning")
                return 

        bonus_rule = popup_config.get("BONUS")
        if not bonus_rule: return

        wheel_path = os.path.join(templates_dir, bonus_rule['wheel'])
        if os.path.exists(wheel_path) and wm.find_template(wheel_path, confidence=confidence):
            self.log("Обнаружено бонусное колесо. Ищу кнопку вращения...", "info")
            time.sleep(0.5)
            spin_btn_path = os.path.join(templates_dir, bonus_rule['spin_button'])
            if wm.find_and_click_template(spin_btn_path, confidence=confidence):
                self.log("Бонус: нажата кнопка вращения. Ожидание...", "info")
                time.sleep(4)
                close_bonus_btn_path = os.path.join(templates_dir, bonus_rule['close_button'])
                if wm.find_and_click_template(close_bonus_btn_path, confidence=confidence):
                    self.log("Бонусное окно успешно закрыто.", "info")
                else:
                    self.log("Не удалось найти кнопку закрытия бонусного окна.", "warning")
            else:
                self.log("Не удалось найти кнопку вращения бонуса.", "warning")
            return

    def _handle_gg_popups(self):
        config = PROJECT_CONFIGS.get("GG", {})
        popup_config = config.get("POPUPS")
        if not popup_config: return

        templates_dir = self._get_templates_dir("gg")
        if not templates_dir: return

        wm = self.window_manager
        confidence = popup_config.get("CONFIDENCE", 0.85)

        buyin_rule = popup_config.get("BUY_IN")
        if not buyin_rule: return

        trigger_path = os.path.join(templates_dir, buyin_rule['trigger'])
        if os.path.exists(trigger_path) and wm.find_template(trigger_path, confidence=confidence):
            self.log("Обнаружено окно Buy-in. Выполняю авто-нажатия...", "info")
            time.sleep(0.5)

            max_btn_path = os.path.join(templates_dir, buyin_rule['max_button'])
            if wm.find_and_click_template(max_btn_path, confidence=confidence):
                self.log("Нажата кнопка 'Max'.", "info")
                time.sleep(0.5)

                confirm_btn_path = os.path.join(templates_dir, buyin_rule['confirm_button'])
                if wm.find_and_click_template(confirm_btn_path, confidence=confidence):
                    self.log("Нажата кнопка подтверждения Buy-in.", "info")
                else:
                    self.log("Не удалось найти кнопку подтверждения Buy-in.", "warning")
            else:
                self.log("Не удалось найти кнопку 'Max'.", "warning")
            return

    def _get_templates_dir(self, project_name: str) -> Optional[str]:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        templates_dir = os.path.join(base_path, "templates", project_name)
        if not os.path.isdir(templates_dir):
            if self.timers["popup_check"].isActive():
                self.log(f"Папка шаблонов {templates_dir} не найдена. Отключаю проверку.", "warning")
                self.timers["popup_check"].stop()
            return None
        return templates_dir

    def position_window_top_right(self):
        try: screen = self.get_current_screen(); geo = screen.availableGeometry(); margin = AppConfig.WINDOW_MARGIN; self.move(geo.right() - self.frameGeometry().width() - margin, geo.top() + margin)
        except Exception as e: logging.error("Could not position window top-right", exc_info=True)

    def position_window_default(self):
        try: screen = self.get_current_screen(); geo = screen.availableGeometry(); margin = AppConfig.WINDOW_MARGIN; self.move(geo.left() + margin, geo.bottom() - self.frameGeometry().height() - margin)
        except Exception as e: logging.error("Could not position window default", exc_info=True)

    def position_gg_panel(self):
        try:
            screen = self.get_current_screen()
            geo = screen.availableGeometry()
            margin = AppConfig.WINDOW_MARGIN
            panel_width = self.frameGeometry().width()
            panel_height = self.frameGeometry().height()
            x = geo.left() + margin
            y = geo.bottom() - panel_height - margin
            self.move(x, y)
        except Exception as e:
            logging.error("Could not position GG panel", exc_info=True)

# ===================================================================
# 6. ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ===================================================================
if __name__ == '__main__':
    instance = SingleInstance("OiHelperMutex")
    if instance.is_already_running():
        logging.warning("Обнаружена уже запущенная копия OiHelper. Завершение работы.")
        sys.exit(0)

    app = QApplication(sys.argv)

    splash = SplashScreen()
    splash.show()

    main_window = MainWindow(splash)

    sys.exit(app.exec())
