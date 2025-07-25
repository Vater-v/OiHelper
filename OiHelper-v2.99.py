# -*- coding: utf-8 -*-
"""
OiHelper - приложение для автоматизации работы с покерными клиентами.
Версия: 2.99

Зависимости:
- PyQt6
- pywin32
- requests
- pyautogui
- win32com.client (опционально, часть pywin32)

Для установки зависимостей:
pip install PyQt6 pywin32 requests pyautogui
"""
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
from ctypes import wintypes
import queue
import random
import math


try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
    logging.warning("PyAutoGUI не найден. Установите его ('pip install pyautogui') для надежной эмуляции кликов.")


from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel,
    QProgressBar, QGridLayout, QHBoxLayout, QGraphicsDropShadowEffect, QFrame, QSizePolicy
)
from PyQt6.QtCore import QObject, QTimer, pyqtSignal, Qt, QPropertyAnimation, QRect, QEasingCurve, pyqtProperty
from PyQt6.QtGui import QIcon, QColor, QPixmap, QFont, QPainter, QBrush, QPen
from PyQt6.QtSvg import QSvgRenderer

# ===================================================================
# 0. НАСТРОЙКА ЛОГИРОВАНИЯ
# ===================================================================
try:
    log_dir = os.path.join(os.getenv('APPDATA'), 'OiHelper')
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, 'app.log')
    file_handler = logging.FileHandler(log_file_path, encoding='utf-8', mode='a')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    root_logger = logging.getLogger()
    root_logger.addHandler(file_handler)
    root_logger.setLevel(logging.INFO)
except Exception as e:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.error(f"Не удалось настроить логирование в файл: {e}")

# ===================================================================
# 1. КОНФИГУРАЦИЯ И СТИЛИ
# ===================================================================

class AppConfig:
    """Централизованная конфигурация приложения."""
    CURRENT_VERSION = "2.99"
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
    
    DEFAULT_WIDTH = 380
    DEFAULT_HEIGHT = 220
    GG_UI_WIDTH = 750
    GG_UI_HEIGHT = 150
    WINDOW_MARGIN = 5
    
    APP_TITLE = f"OiHelper v{CURRENT_VERSION}"
    MSG_PANEL_TITLE = "Панель: {}"
    STATUS_MSG_PLAYER_SEARCH = "Статус: Поиск плеера..."
    STATUS_MSG_PLAYER_NOT_FOUND = "Статус: Плеер не найден"
    STATUS_MSG_PRESS_START = "Статус: Ожидание старта..."
    STATUS_MSG_OK = "Статус: Все системы в норме"
    MSG_ARRANGE_TABLES = "Расставить столы"
    MSG_ARRANGE_SYSTEM = "Системные окна" 
    MSG_CLICK_COMMAND = "Sit-Out всем"
    MSG_PROGRESS_LABEL = "Перезапись через: {}"
    MSG_LIMIT_REACHED = "<b>Лимит!</b>"
    MSG_UPDATE_CHECK = "Проверка обновлений..."
    MSG_UPDATE_FAIL = "Ошибка обновления. Работа в оффлайн-режиме."
    MSG_UPTIME_WARNING = "Компьютер не перезагружался более 5 дней."
    MSG_ADMIN_WARNING = "Нет прав администратора. Функции могут быть ограничены."
    MSG_ARRANGE_TABLES_NOT_FOUND = "Столы для расстановки не найдены."
    
    PLAYER_CHECK_INTERVAL = 5000
    PLAYER_AUTOSTART_INTERVAL = 7000
    AUTO_RECORD_INTERVAL = 5000
    AUTO_ARRANGE_INTERVAL = 3000
    RECORDER_CHECK_INTERVAL = 15000 
    NOTIFICATION_DURATION = 4000
    STATUS_MESSAGE_DURATION = 2000

class ColorPalette:
    BACKGROUND = "#23272E"
    SURFACE = "#343944"
    PRIMARY = "#0090FF" 
    PRIMARY_HOVER = "#33A5FF"
    PRIMARY_PRESSED = "#0078D7"
    SECONDARY = "#4A4D52"
    SECONDARY_HOVER = "#5A5D62"
    SECONDARY_PRESSED = "#3A3D42"
    GREEN = "#28A745"
    RED = "#DC3545"
    AMBER = "#FFC107"
    TEXT_PRIMARY = "#EAEAEA"
    TEXT_SECONDARY = "#9A9A9A"
    BORDER = "#404040"

class StyleSheet:
    MAIN_WINDOW = f"background-color: {ColorPalette.BACKGROUND};"
    TITLE_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 16px; font-weight: bold; color: {ColorPalette.TEXT_PRIMARY};"
    STATUS_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 13px; color: {ColorPalette.TEXT_SECONDARY};"
    PROGRESS_BAR_LABEL = f"font-family: 'Segoe UI', 'Roboto'; font-size: 12px; color: {ColorPalette.TEXT_SECONDARY};"
    
    @staticmethod
    def get_button_style(primary=True):
        if primary:
            bg, hover, pressed = (ColorPalette.PRIMARY, ColorPalette.PRIMARY_HOVER, ColorPalette.PRIMARY_PRESSED)
        else:
            bg, hover, pressed = (ColorPalette.SURFACE, ColorPalette.SECONDARY_HOVER, ColorPalette.SECONDARY_PRESSED)
        return f"""
            QPushButton {{ 
                background-color: {bg}; 
                color: {ColorPalette.TEXT_PRIMARY}; 
                font-family: 'Segoe UI', 'Roboto'; 
                font-size: 13px; 
                font-weight: bold; 
                border: 1px solid {ColorPalette.BORDER}; 
                border-radius: 6px; 
                padding: 8px 12px; 
            }} 
            QPushButton:hover {{ background-color: {hover}; }} 
            QPushButton:pressed {{ background-color: {pressed}; }} 
            QPushButton:disabled {{ background-color: {ColorPalette.SURFACE}; color: #666; border: 1px solid #404040; }}
        """

PROJECT_CONFIGS = {
    "GG": { "TABLE": { "FIND_METHOD": "RATIO", "W": 557, "H": 424, "TOLERANCE": 0.035 }, "LOBBY": { "FIND_METHOD": "RATIO", "W": 333, "H": 623, "TOLERANCE": 0.07, "X": 1580, "Y": 140 }, "PLAYER": { "W": 700, "H": 365, "X": 1385, "Y": 0 }, "TABLE_SLOTS": [(-5, 0), (271, 423), (816, 0), (1086, 423)], "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe"], "SESSION_MAX_DURATION_S": 4 * 3600, "SESSION_WARN_TIME_S": 3.5 * 3600, "ARRANGE_MINIMIZED_TABLES": False },
    "QQ": { "TABLE": { "FIND_METHOD": "TITLE_AND_SIZE", "TITLE": "QQPK", "W": 400, "H": 700, "TOLERANCE": 2 }, "LOBBY": { "FIND_METHOD": "RATIO", "W": 400, "H": 700, "TOLERANCE": 0.07, "X": 1418, "Y": 0 }, "CV_SERVER": { "FIND_METHOD": "TITLE", "TITLE": "OpenCv", "X": 1789, "Y": 367, "W": 993, "H": 605 }, "PLAYER": { "X": 1418, "Y": 942, "W": 724, "H": 370 }, "TABLE_SLOTS": [(0, 0), (401, 0), (802, 0), (1203, 0)], "TABLE_SLOTS_5": [(0, 0), (346, 0), (692, 0), (1038, 0), (1384, 0)], "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "OpenCvServer.exe"], "SESSION_MAX_DURATION_S": 3 * 3600, "SESSION_WARN_TIME_S": -1, "ARRANGE_MINIMIZED_TABLES": True }
}

# ===================================================================
# 2. СИСТЕМА ИНТЕРФЕЙСА (УВЕДОМЛЕНИЯ, ИНДИКАТОРЫ, ПЕРЕКЛЮЧАТЕЛИ)
# ===================================================================
class ToggleSwitch(QPushButton):
    """Кастомный виджет переключателя с анимацией."""
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

    def update_state(self):
        self.animation.setStartValue(self._circle_pos)
        self.animation.setEndValue(20 if self.isChecked() else 3)
        self.animation.start()

    def setChecked(self, checked):
        super().setChecked(checked)
        self.update_state()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        bg_color = QColor(ColorPalette.PRIMARY) if self.isChecked() else QColor(ColorPalette.SURFACE)
        painter.setBrush(QBrush(bg_color))
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 11, 11)
        painter.setBrush(QBrush(QColor(ColorPalette.TEXT_PRIMARY)))
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
        container.setStyleSheet(f"""
            QFrame {{ 
                background-color: {ColorPalette.SURFACE}; 
                border-radius: 6px; 
                border-left: 5px solid {color};
            }}
        """)
        
        shadow_layout = QVBoxLayout(self)
        shadow_layout.setContentsMargins(5, 5, 5, 5)
        shadow_layout.addWidget(container)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 160))
        shadow.setOffset(0, 2)
        container.setGraphicsEffect(shadow)

        layout = QHBoxLayout(container)
        layout.setContentsMargins(15, 10, 15, 10)
        
        text_label = QLabel(message)
        text_label.setFont(QFont("Segoe UI", 12))
        text_label.setWordWrap(True)
        text_label.setStyleSheet(f"background: transparent; border: none; color: {ColorPalette.TEXT_PRIMARY};")
        layout.addWidget(text_label, 1)

        self.setFixedWidth(320)

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
        if self.opacity <= 0.0:
            self.fade_out_timer.stop()
            self.close()

    def show_animation(self):
        self.show()
        self.fade_in_timer.start(20)

    def start_fade_out(self):
        if self.is_closing: return
        self.is_closing = True
        self.fade_in_timer.stop()
        self.fade_out_timer.start(20)

    def closeEvent(self, event):
        self.closed.emit(self)
        super().closeEvent(event)

class NotificationManager(QObject):
    def __init__(self):
        super().__init__()
        self.notifications = []

    def show(self, message, message_type):
        logging.info(f"Уведомление [{message_type}]: {message}")
        
        # Не показывать дубликаты
        if any(n.message == message for n in self.notifications):
            return

        # Ограничение на 2 уведомления
        if len(self.notifications) >= 2:
            self.notifications.pop(0).start_fade_out()

        notification = Notification(message, message_type)
        notification.closed.connect(self.on_notification_closed)
        self.notifications.append(notification)
        self.reposition_all()
        notification.show_animation()

    def on_notification_closed(self, notification):
        if notification in self.notifications:
            self.notifications.remove(notification)
        self.reposition_all()

    def reposition_all(self):
        try:
            screen_geo = QApplication.primaryScreen().availableGeometry()
        except AttributeError:
            screen_geo = QApplication.screens()[0].availableGeometry()
        
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
    """Кастомный анимированный прогресс-бар с динамической сменой цвета."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._value = 0.0
        self._maximum = 14400.0
        
        self._animation = QPropertyAnimation(self, b"progress_value", self)
        self._animation.setEasingCurve(QEasingCurve.Type.InOutCubic)

        # Для мигания в режиме "Лимит"
        self._alert_mode = False
        self._alert_visible = True
        self._alert_timer = QTimer(self)
        self._alert_timer.timeout.connect(self._toggle_alert_visibility)

    def _toggle_alert_visibility(self):
        self._alert_visible = not self._alert_visible
        self.update()

    def get_progress_value(self):
        return self._value

    def set_progress_value(self, value):
        self._value = value
        self.update()

    progress_value = pyqtProperty(float, fget=get_progress_value, fset=set_progress_value)

    def setValue(self, value):
        value = max(0, min(float(value), self._maximum))
        
        # Запускаем/обновляем анимацию
        if self._animation.state() == QPropertyAnimation.State.Running:
            self._animation.stop()
        self._animation.setStartValue(self.progress_value)
        self._animation.setEndValue(value)
        self._animation.setDuration(1000) # Анимация длится 1 секунду
        self._animation.start()
        
        # Проверяем режим "Лимит"
        is_limit_reached = value >= self._maximum
        if is_limit_reached and not self._alert_mode:
            self._alert_mode = True
            self._alert_timer.start(400) # Мигание каждые 400 мс
        elif not is_limit_reached and self._alert_mode:
            self._alert_mode = False
            self._alert_timer.stop()
            self._alert_visible = True
            self.update()

    def setMaximum(self, value):
        self._maximum = float(value)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Фон прогресс-бара
        bg_rect = self.rect()
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(ColorPalette.SURFACE))
        painter.drawRoundedRect(bg_rect, self.height() // 2, self.height() // 2)

        # Рассчитываем процент и ширину
        percent = (self.progress_value / self._maximum) if self._maximum > 0 else 0
        progress_width = percent * bg_rect.width()

        # Определяем цвет на основе процента
        if self._alert_mode:
            color = QColor(ColorPalette.RED) if self._alert_visible else QColor(ColorPalette.SURFACE)
        elif percent > 0.85:
            color = QColor(ColorPalette.RED)
        elif percent > 0.50:
            color = QColor(ColorPalette.AMBER)
        else:
            color = QColor(ColorPalette.PRIMARY)
            
        # Рисуем сам прогресс
        if progress_width > 0:
            progress_rect = QRect(0, 0, int(progress_width), self.height())
            painter.setBrush(color)
            painter.drawRoundedRect(progress_rect, self.height() // 2, self.height() // 2)

class ClickIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint); self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground); self.setFixedSize(30, 30); self.timer = QTimer(self); self.timer.setSingleShot(True); self.timer.timeout.connect(self.hide)
    def paintEvent(self, event): painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.setBrush(QColor(255, 0, 0, 150)); painter.setPen(Qt.PenStyle.NoPen); painter.drawEllipse(self.rect())
    def show_at(self, x, y): self.move(x - self.width() // 2, y - self.height() // 2); self.show(); self.timer.start(200)

class WindowManager(QObject):
    log_request = pyqtSignal(str, str)
    click_visual_request = pyqtSignal(int, int)
    
    def __init__(self):
        super().__init__()
        # CTYPES-структуры для SendInput
        self.INPUT_MOUSE = 0
        self.MOUSEEVENTF_LEFTDOWN = 0x0002
        self.MOUSEEVENTF_LEFTUP = 0x0004
        self.MOUSEEVENTF_MOVE = 0x0001
        self.MOUSEEVENTF_ABSOLUTE = 0x8000
        
        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", wintypes.LONG),
                        ("dy", wintypes.LONG),
                        ("mouseData", wintypes.DWORD),
                        ("dwFlags", wintypes.DWORD),
                        ("time", wintypes.DWORD),
                        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", wintypes.DWORD),
                        ("mi", MOUSEINPUT)]

        self.INPUT_STRUCT = INPUT
        self.screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    def _send_mouse_input(self, flags, x=0, y=0):
        """Вспомогательная функция для отправки событий мыши через SendInput."""
        if flags & self.MOUSEEVENTF_ABSOLUTE:
            x = (x * 65535) // self.screen_width
            y = (y * 65535) // self.screen_height

        mouse_input = self.INPUT_STRUCT(
            type=self.INPUT_MOUSE,
            mi=self.INPUT_STRUCT._fields_[1][1](
                dx=x, dy=y, mouseData=0, dwFlags=flags, time=0, dwExtraInfo=None
            )
        )
        ctypes.windll.user32.SendInput(1, ctypes.byref(mouse_input), ctypes.sizeof(mouse_input))

    def humanized_click(self, x, y):
        """
        Выполняет "очеловеченный" клик: движение по кривой Безье,
        случайные задержки, дрожание и финальный клик.
        """
        try:
            self.click_visual_request.emit(x, y)
            start_x, start_y = win32api.GetCursorPos()

            # 1. Рассчитываем траекторию по кубической кривой Безье
            dist = math.hypot(x - start_x, y - start_y)
            duration_ms = max(150, min(500, int(dist * 1.2))) # Длительность зависит от расстояния
            num_steps = max(10, int(duration_ms / 15))

            # Контрольные точки для кривой
            offset = int(dist * 0.2)
            control1_x = start_x + random.randint(-offset, offset)
            control1_y = start_y + random.randint(-offset, offset)
            control2_x = x + random.randint(-offset, offset)
            control2_y = y + random.randint(-offset, offset)

            points = []
            for i in range(num_steps + 1):
                t = i / num_steps
                inv_t = 1 - t
                b_x = (inv_t**3 * start_x +
                       3 * inv_t**2 * t * control1_x +
                       3 * inv_t * t**2 * control2_x +
                       t**3 * x)
                b_y = (inv_t**3 * start_y +
                       3 * inv_t**2 * t * control1_y +
                       3 * inv_t * t**2 * control2_y +
                       t**3 * y)
                points.append((int(b_x), int(b_y)))

            # 2. Движение по траектории
            for point_x, point_y in points:
                self._send_mouse_input(self.MOUSEEVENTF_MOVE | self.MOUSEEVENTF_ABSOLUTE, point_x, point_y)
                time.sleep(random.uniform(0.003, 0.01))

            # 3. "Дрожание" у цели
            for _ in range(random.randint(1, 3)):
                jitter_x = x + random.randint(-2, 2)
                jitter_y = y + random.randint(-2, 2)
                self._send_mouse_input(self.MOUSEEVENTF_MOVE | self.MOUSEEVENTF_ABSOLUTE, jitter_x, jitter_y)
                time.sleep(random.uniform(0.005, 0.015))

            # 4. Финальный клик
            self._send_mouse_input(self.MOUSEEVENTF_LEFTDOWN | self.MOUSEEVENTF_ABSOLUTE, x, y)
            time.sleep(random.uniform(0.04, 0.09))
            self._send_mouse_input(self.MOUSEEVENTF_LEFTUP | self.MOUSEEVENTF_ABSOLUTE, x, y)
            
            logging.info(f"Выполнен 'очеловеченный' клик по ({x},{y})")

        except Exception as e:
            self.log_request.emit(f"Ошибка 'очеловеченного' клика по ({x},{y}): {e}", "error")

    def find_windows_by_config(self, config, config_key, main_window_hwnd):
        window_config = config.get(config_key, {})
        if not window_config: return []
        find_method = window_config.get("FIND_METHOD"); EXCLUDED_TITLES = config.get("EXCLUDED_TITLES", []); EXCLUDED_PROCESSES = config.get("EXCLUDED_PROCESSES", []); arrange_minimized = config.get("ARRANGE_MINIMIZED_TABLES", False) and config_key == "TABLE"
        found_windows = []
        def enum_windows_callback(hwnd, _):
            if hwnd == main_window_hwnd or not win32gui.IsWindowVisible(hwnd) or (not arrange_minimized and win32gui.IsIconic(hwnd)): return True
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd); h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid); process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0)); win32api.CloseHandle(h_process)
                if process_name.lower() in EXCLUDED_PROCESSES: return True
                title = win32gui.GetWindowText(hwnd)
                if any(excluded.lower() in title.lower() for excluded in EXCLUDED_TITLES): return True
                rect = win32gui.GetWindowRect(hwnd); w, h = rect[2] - rect[0], rect[3] - rect[1]
                if w == 0 or h == 0: return True
                match = False
                if find_method == "RATIO":
                    if h != 0 and (window_config["W"] / window_config["H"]) * (1 - window_config["TOLERANCE"]) <= (w / h) <= (window_config["W"] / window_config["H"]) * (1 + window_config["TOLERANCE"]): match = True
                elif find_method == "TITLE_AND_SIZE":
                    if window_config["TITLE"].lower() in title.lower() and abs(w - window_config["W"]) <= window_config["TOLERANCE"] and abs(h - window_config["H"]) <= window_config["TOLERANCE"]: match = True
                elif find_method == "TITLE":
                    if window_config["TITLE"].lower() in title.lower(): match = True
                if match: found_windows.append(hwnd)
            except Exception: pass
            return True
        try: win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e: logging.error(f"Критическая ошибка в EnumWindows: {e}", exc_info=True)
        try: found_windows.sort(key=lambda hwnd: win32gui.GetWindowRect(hwnd)[0])
        except Exception as e: self.log_request.emit(f"Ошибка сортировки окон: {e}", "warning")
        return found_windows

    def find_first_window_by_title(self, text_in_title, exact_match=False):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try: title = win32gui.GetWindowText(hwnd)
                except Exception: return
                if (exact_match and text_in_title == title) or (not exact_match and text_in_title.lower() in title.lower()): hwnds.append(hwnd)
        try: win32gui.EnumWindows(callback, None)
        except Exception as e: logging.error(f"Критическая ошибка в EnumWindows (поиск по заголовку): {e}", exc_info=True)
        return hwnds[0] if hwnds else None
    def is_process_running(self, process_name_to_find):
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
    def find_first_window_by_process_name(self, process_name_to_find):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd); h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid); process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0)); win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in process_name.lower(): hwnds.append(hwnd)
                except Exception: pass
        try: win32gui.EnumWindows(callback, None)
        except Exception as e: logging.error(f"Критическая ошибка в EnumWindows (поиск по процессу): {e}", exc_info=True)
        return hwnds[0] if hwnds else None
    def press_key(self, key_code):
        try: win32api.keybd_event(key_code, 0, 0, 0); time.sleep(0.05); win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e: self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

class TelegramNotifier(QObject):
    def __init__(self, token, chat_id):
        super().__init__(); self.token = token; self.chat_id = chat_id; self.api_url = f"https://api.telegram.org/bot{self.token}/sendMessage"; self.message_queue = queue.Queue(); self.worker_thread = threading.Thread(target=self._worker, daemon=True); self.worker_thread.start()
    def send_message(self, message):
        if not self.token or not self.chat_id: logging.warning("Токен или ID чата Telegram не настроены."); return
        self.message_queue.put(message)
    def _worker(self):
        while True:
            message = self.message_queue.get()
            try: requests.post(self.api_url, data={'chat_id': self.chat_id, 'text': message}, timeout=10).raise_for_status(); logging.info("Сообщение в Telegram успешно отправлено.")
            except requests.RequestException as e: logging.error(f"Не удалось отправить сообщение в Telegram: {e}")
            self.message_queue.task_done()

class UpdateManager(QObject):
    log_request = pyqtSignal(str, str); check_finished = pyqtSignal()
    def __init__(self): super().__init__(); self.update_info = {}
    
    def is_new_version_available(self, current_v_str, latest_v_str):
        """ИСПРАВЛЕНО: Корректное сравнение версий."""
        try:
            current = [int(p) for p in current_v_str.lstrip('v').split('.')]
            latest = [int(p) for p in latest_v_str.lstrip('v').split('.')]
            max_len = max(len(current), len(latest))
            current += [0] * (max_len - len(current))
            latest += [0] * (max_len - len(latest))
            return latest > current
        except Exception as e:
            logging.error(f"Ошибка сравнения версий: {e}.")
            return latest_v_str > current_v_str

    def check_for_updates(self):
        self.log_request.emit(AppConfig.MSG_UPDATE_CHECK, "info")
        try:
            api_url = f"https://api.github.com/repos/{AppConfig.GITHUB_REPO}/releases/latest"; response = requests.get(api_url, timeout=10); response.raise_for_status(); latest_release = response.json()
            if (latest_version := latest_release.get("tag_name")) and self.is_new_version_available(AppConfig.CURRENT_VERSION, latest_version):
                self.log_request.emit(f"Доступна новая версия: {latest_version}. Обновление...", "info"); self.update_info = latest_release; threading.Thread(target=self.apply_update, daemon=True).start()
            else: self.log_request.emit("Вы используете последнюю версию.", "info"); self.check_finished.emit()
        except requests.RequestException as e: self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error"); self.check_finished.emit()
        except Exception as e: logging.error(f"Неожиданная ошибка при проверке обновлений: {e}", exc_info=True); self.check_finished.emit()
    def apply_update(self):
        download_url = next((asset["browser_download_url"] for asset in self.update_info.get("assets", []) if asset["name"] == AppConfig.ASSET_NAME), None)
        if not download_url: self.log_request.emit("Не удалось найти ZIP-архив в релизе.", "error"); return
        self.download_and_run_updater(download_url)
    def download_and_run_updater(self, url):
        update_zip_name = "update.zip"
        try:
            self.log_request.emit("Скачивание обновления...", "info"); response = requests.get(url, stream=True, timeout=60); response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
            self.log_request.emit("Распаковка архива...", "info"); update_folder = "update_temp"
            if os.path.isdir(update_folder): import shutil; shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref: zip_ref.extractall(update_folder)
            self.log_request.emit("Обновление скачано. Перезапуск...", "info"); updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]); current_dir = os.path.dirname(current_exe_path); exe_name = os.path.basename(current_exe_path)
            script_content = f'@echo off\nchcp 65001 > NUL\necho Waiting for OiHelper to close...\ntimeout /t 2 /nobreak > NUL\ntaskkill /pid {os.getpid()} /f > NUL\necho Waiting for process to terminate...\ntimeout /t 3 /nobreak > NUL\necho Moving new files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\nrd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\ndel "{current_dir}\\{update_zip_name}"\necho Starting new version...\nstart "" "{exe_name}"\n(goto) 2>nul & del "%~f0"'
            with open(updater_script_path, "w", encoding="cp866") as f: f.write(script_content)
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE); QApplication.instance().quit()
        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error"); logging.error(f"Update error: {e}", exc_info=True)
            if os.path.exists(update_zip_name):
                try: os.remove(update_zip_name)
                except OSError as err: logging.error(f"Не удалось удалить временный файл обновления: {err}")
            self.log_request.emit(AppConfig.MSG_UPDATE_FAIL, "warning"); self.check_finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.notification_manager = NotificationManager(); self.window_manager = WindowManager(); self.update_manager = UpdateManager(); self.telegram_notifier = TelegramNotifier(AppConfig.TELEGRAM_BOT_TOKEN, AppConfig.TELEGRAM_CHAT_ID); self.click_indicator = ClickIndicator()
        self.current_project = None; self.is_auto_record_enabled = True; self.is_automation_enabled = True; self.last_table_count = 0; self.recording_start_time = 0; self.is_sending_logs = False; self.player_was_open = False; self.is_record_stopping = False
        self.shell = win32com.client.Dispatch("WScript.Shell") if WIN32COM_AVAILABLE else None
        self.status_message_timer = QTimer(self)
        self.status_message_timer.setSingleShot(True)
        self.status_message_timer.timeout.connect(self.clear_status_message)
        self.init_ui(); self.connect_signals(); self.init_timers(); self.init_startup_checks(); self.sync_ui_state()

    def closeEvent(self, event): logging.info("Приложение закрывается, остановка таймеров...");_ = [timer.stop() for timer in self.timers.values()]; super().closeEvent(event)
    def get_current_screen(self):
        try: return self.screen() or QApplication.primaryScreen()
        except Exception: return QApplication.primaryScreen()
    def log(self, message, message_type): self.notification_manager.show(message, message_type);_ = self.telegram_notifier.send_message(f"OiHelper Критическая ошибка: {message}") if message_type == 'error' else None
    
    def set_status_message(self, message, is_persistent=False):
        if hasattr(self, 'status_label'):
            self.status_label.setText(message)
            if not is_persistent:
                self.status_message_timer.start(AppConfig.STATUS_MESSAGE_DURATION)

    def clear_status_message(self):
        if hasattr(self, 'status_label'):
            self.status_label.setText(AppConfig.STATUS_MSG_OK)

    def init_ui(self):
        """Инициализация и компоновка нового интерфейса."""
        self.setWindowTitle(AppConfig.APP_TITLE)
        if os.path.exists(AppConfig.ICON_PATH): self.setWindowIcon(QIcon(AppConfig.ICON_PATH))
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet(StyleSheet.MAIN_WINDOW)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.build_ui_for_project()

    def build_ui_for_project(self):
        if self.central_widget.layout() is not None:
            QWidget().setLayout(self.central_widget.layout())

        if self.current_project == "GG":
            self.build_gg_panel_ui()
        else:
            self.build_default_ui()
        
        self.connect_signals()

    def build_default_ui(self):
        container = self.central_widget
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(15, 10, 15, 15)
        main_layout.setSpacing(12)

        # --- Зона заголовка ---
        self.project_label = QLabel()
        self.project_label.setStyleSheet(StyleSheet.TITLE_LABEL)
        main_layout.addWidget(self.project_label)
        
        separator = QFrame(); separator.setFrameShape(QFrame.Shape.HLine); separator.setFrameShadow(QFrame.Shadow.Sunken); separator.setStyleSheet(f"color: {ColorPalette.BORDER};")
        main_layout.addWidget(separator)
        
        # --- Зона управления ---
        controls_frame = QFrame()
        controls_layout = QGridLayout(controls_frame)
        controls_layout.setContentsMargins(0, 5, 0, 5)
        controls_layout.setSpacing(10)

        self.automation_toggle = ToggleSwitch(); self.auto_record_toggle = ToggleSwitch()
        automation_label = QLabel("Автоматика"); automation_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        auto_record_label = QLabel("Автозапись"); auto_record_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        
        controls_layout.addWidget(automation_label, 0, 0); controls_layout.addWidget(self.automation_toggle, 0, 1)
        controls_layout.addWidget(auto_record_label, 1, 0); controls_layout.addWidget(self.auto_record_toggle, 1, 1)

        self.arrange_tables_button = QPushButton(AppConfig.MSG_ARRANGE_TABLES); self.arrange_tables_button.setStyleSheet(StyleSheet.get_button_style(primary=True))
        self.arrange_system_button = QPushButton(AppConfig.MSG_ARRANGE_SYSTEM); self.arrange_system_button.setStyleSheet(StyleSheet.get_button_style(False))
        self.sit_out_button = QPushButton(AppConfig.MSG_CLICK_COMMAND); self.sit_out_button.setStyleSheet(StyleSheet.get_button_style(False))
        
        controls_layout.addWidget(self.arrange_tables_button, 0, 2, 2, 1, Qt.AlignmentFlag.AlignVCenter)
        controls_layout.addWidget(self.arrange_system_button, 0, 3)
        controls_layout.addWidget(self.sit_out_button, 1, 3)
        controls_layout.setColumnStretch(2, 1)

        main_layout.addWidget(controls_frame)
        
        main_layout.addStretch(1)

        # --- Зона статуса ---
        status_frame = QFrame()
        status_layout = QVBoxLayout(status_frame)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(5)

        self.status_label = QLabel(AppConfig.STATUS_MSG_OK)
        self.status_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        status_layout.addWidget(self.status_label)

        self.progress_frame = QFrame()
        progress_bar_layout = QHBoxLayout(self.progress_frame)
        progress_bar_layout.setContentsMargins(0, 0, 0, 0)
        progress_bar_layout.setSpacing(10)
        
        self.progress_bar = AnimatedProgressBar()
        self.progress_bar.setFixedHeight(10)
        self.progress_bar_label = QLabel()
        self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)
        
        progress_bar_layout.addWidget(self.progress_bar, 2)
        progress_bar_layout.addWidget(self.progress_bar_label, 1)
        
        status_layout.addWidget(self.progress_frame)
        main_layout.addWidget(status_frame)

    def build_gg_panel_ui(self):
        container = self.central_widget
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(15, 10, 15, 10)
        main_layout.setSpacing(8)

        # --- Зона заголовка ---
        self.project_label = QLabel()
        self.project_label.setStyleSheet(StyleSheet.TITLE_LABEL)
        main_layout.addWidget(self.project_label)
        
        # --- Зона управления ---
        controls_frame = QFrame()
        controls_layout = QHBoxLayout(controls_frame)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(10)

        toggles_layout = QGridLayout()
        toggles_layout.setSpacing(5)
        self.automation_toggle = ToggleSwitch()
        self.auto_record_toggle = ToggleSwitch()
        automation_label = QLabel("Автоматика"); automation_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        auto_record_label = QLabel("Автозапись"); auto_record_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        toggles_layout.addWidget(automation_label, 0, 0); toggles_layout.addWidget(self.automation_toggle, 0, 1)
        toggles_layout.addWidget(auto_record_label, 1, 0); toggles_layout.addWidget(self.auto_record_toggle, 1, 1)
        controls_layout.addLayout(toggles_layout)

        separator = QFrame(); separator.setFrameShape(QFrame.Shape.VLine); separator.setFrameShadow(QFrame.Shadow.Sunken); separator.setStyleSheet(f"color: {ColorPalette.BORDER};")
        controls_layout.addWidget(separator)
        
        self.arrange_tables_button = QPushButton(AppConfig.MSG_ARRANGE_TABLES); self.arrange_tables_button.setStyleSheet(StyleSheet.get_button_style(primary=True))
        self.arrange_system_button = QPushButton(AppConfig.MSG_ARRANGE_SYSTEM); self.arrange_system_button.setStyleSheet(StyleSheet.get_button_style(False))
        self.sit_out_button = QPushButton(AppConfig.MSG_CLICK_COMMAND); self.sit_out_button.setStyleSheet(StyleSheet.get_button_style(False))
        controls_layout.addWidget(self.arrange_tables_button)
        controls_layout.addWidget(self.arrange_system_button)
        controls_layout.addWidget(self.sit_out_button)
        controls_layout.addStretch(1)
        main_layout.addWidget(controls_frame)

        # --- Зона статуса ---
        status_frame = QFrame()
        status_layout = QVBoxLayout(status_frame)
        status_layout.setContentsMargins(0, 5, 0, 0)
        status_layout.setSpacing(5)

        self.status_label = QLabel(AppConfig.STATUS_MSG_OK)
        self.status_label.setStyleSheet(StyleSheet.STATUS_LABEL)
        status_layout.addWidget(self.status_label)

        self.progress_frame = QFrame()
        progress_bar_layout = QHBoxLayout(self.progress_frame)
        progress_bar_layout.setContentsMargins(0, 0, 0, 0)
        progress_bar_layout.setSpacing(10)
        
        self.progress_bar = AnimatedProgressBar()
        self.progress_bar.setFixedHeight(8)
        self.progress_bar_label = QLabel()
        self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)
        
        progress_bar_layout.addWidget(self.progress_bar, 2)
        progress_bar_layout.addWidget(self.progress_bar_label, 1)
        
        status_layout.addWidget(self.progress_frame)
        main_layout.addWidget(status_frame)
        main_layout.addStretch(1)

    def sync_ui_state(self):
        self.build_ui_for_project() # Перестраиваем UI
        
        if self.current_project == "GG":
            self.setFixedSize(AppConfig.GG_UI_WIDTH, AppConfig.GG_UI_HEIGHT)
            self.project_label.setText(AppConfig.MSG_PANEL_TITLE.format("GG"))
            self.position_window_gg_default()
        elif self.current_project == "QQ":
            self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
            self.project_label.setText(AppConfig.MSG_PANEL_TITLE.format("QQ"))
            self.position_window_top_right()
        else:
            self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
            self.project_label.setText("Панель: OiHelper")
            self.position_window_default()
        
        self.automation_toggle.setChecked(self.is_automation_enabled)
        self.auto_record_toggle.setChecked(self.is_auto_record_enabled)
        
        is_recording = self.recording_start_time > 0
        self.progress_frame.setVisible(is_recording)
        
        is_project_active = self.current_project is not None
        self.arrange_tables_button.setEnabled(is_project_active)
        self.arrange_system_button.setEnabled(is_project_active)
        self.sit_out_button.setVisible(self.current_project == "GG")
        self.sit_out_button.setEnabled(is_project_active)

    def connect_signals(self):
        self.window_manager.log_request.connect(self.log); self.window_manager.click_visual_request.connect(self.click_indicator.show_at)
        self.update_manager.log_request.connect(self.log); self.update_manager.check_finished.connect(self.start_main_logic)
        
        self.automation_toggle.clicked.connect(self.toggle_automation)
        self.auto_record_toggle.clicked.connect(self.toggle_auto_record)
        self.arrange_tables_button.clicked.connect(self.arrange_tables)
        self.arrange_system_button.clicked.connect(self.arrange_other_windows)
        self.sit_out_button.clicked.connect(self.perform_special_clicks)

    def init_timers(self):
        self.timers = { "player_check": QTimer(self), "recorder_check": QTimer(self), "auto_record": QTimer(self), "auto_arrange": QTimer(self), "session": QTimer(self), "player_start": QTimer(self), "record_cooldown": QTimer(self) }
        self.timers["player_check"].timeout.connect(self.check_for_player); self.timers["recorder_check"].timeout.connect(self.check_for_recorder); self.timers["auto_record"].timeout.connect(self.check_auto_record_logic); self.timers["auto_arrange"].timeout.connect(self.check_for_new_tables); self.timers["session"].timeout.connect(self.update_session_progress); self.timers["player_start"].timeout.connect(self.attempt_player_start_click); self.timers["record_cooldown"].setSingleShot(True); self.timers["record_cooldown"].timeout.connect(lambda: setattr(self, 'is_record_stopping', False))
    
    def init_startup_checks(self):
        self.sync_ui_state()
        self.set_status_message(AppConfig.MSG_UPDATE_CHECK, is_persistent=True)
        self.arrange_tables_button.setEnabled(False)
        self.arrange_system_button.setEnabled(False)
        self.sit_out_button.setEnabled(False)
        threading.Thread(target=self.update_manager.check_for_updates, daemon=True).start()
    
    def start_main_logic(self):
        self.sync_ui_state()
        self.check_for_player()
        self.check_for_recorder()
        self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        QTimer.singleShot(1000, self.initial_recorder_sync_check)
        if AppConfig.TELEGRAM_REPORT_LEVEL == 'all':
            self.telegram_notifier.send_message(f"OiHelper {AppConfig.CURRENT_VERSION} запущен.")
        self.check_system_uptime()
        self.check_admin_rights()
    
    def on_project_changed(self, new_project_name):
        if self.current_project == new_project_name: return
        self.current_project = new_project_name
        self.last_table_count = 0
        self.timers["auto_arrange"].stop()
        if new_project_name:
            self.timers["auto_arrange"].start(AppConfig.AUTO_ARRANGE_INTERVAL)
        
        self.sync_ui_state()

        if self.is_automation_enabled:
            if new_project_name == "QQ": self.check_and_launch_opencv_server()
            elif new_project_name == "GG": self.minimize_injector_window()
            if new_project_name: self.arrange_other_windows()
        
    def check_for_player(self):
        if self.is_sending_logs: return
        player_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if not player_hwnd:
            self.timers["player_start"].stop(); self.on_project_changed(None)
            if self.player_was_open:
                self.player_was_open = False; self.handle_player_close()
            else:
                self.set_status_message(AppConfig.STATUS_MSG_PLAYER_NOT_FOUND)
                if self.is_automation_enabled: self.check_and_launch_player()
        else:
            self.player_was_open = True; project_name = None
            try:
                title = win32gui.GetWindowText(player_hwnd)
                project_name = next((short for full, short in {"QQPoker": "QQ", "ClubGG": "GG"}.items() if f"[{full}]" in title), None)
            except Exception: project_name = None
            if project_name:
                self.timers["player_start"].stop()
                if self.timers["player_start"].isActive(): self.log("Авто-старт плеера успешно завершен.", "info")
                self.on_project_changed(project_name)
            else:
                self.on_project_changed(None)
                self.set_status_message(AppConfig.STATUS_MSG_PRESS_START)
                if self.is_automation_enabled and not self.timers["player_start"].isActive():
                    self.log("Лаунчер найден. Включаю попытки авто-старта...", "info"); self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL); self.attempt_player_start_click()
        if not self.timers["player_check"].isActive(): self.timers["player_check"].start(AppConfig.PLAYER_CHECK_INTERVAL)
    
    def check_admin_rights(self):
        try:
            if ctypes.windll.shell32.IsUserAnAdmin() == 0: self.log(AppConfig.MSG_ADMIN_WARNING, "warning")
        except Exception as e: logging.error(f"Не удалось проверить права администратора: {e}")
    
    def check_system_uptime(self):
        try:
            if (ctypes.windll.kernel32.GetTickCount64() / (1000 * 60 * 60 * 24)) > 5: self.log(AppConfig.MSG_UPTIME_WARNING, "warning")
        except Exception as e: logging.error(f"Не удалось проверить время работы системы: {e}")
    
    def handle_player_close(self):
        if not self.is_automation_enabled: return
        self.is_sending_logs = True; self.log("Плеер закрыт. Запускаю отправку логов...", "info"); desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop'); launched_count = 0
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
        self.log(f"Перезапуск плеера через {AppConfig.PLAYER_RELAUNCH_DELAY_S} секунд...", "info"); QTimer.singleShot(AppConfig.PLAYER_RELAUNCH_DELAY_S * 1000, lambda: self.wait_for_logs_to_finish(time.monotonic()))
    
    def wait_for_logs_to_finish(self, start_time):
        if (time.monotonic() - start_time) > AppConfig.LOG_SENDER_TIMEOUT_S: self.log("Тайм-аут ожидания отправки логов. Возобновление работы.", "error"); self.is_sending_logs = False; self.check_for_player(); return
        if any(self.window_manager.find_first_window_by_title(k) for k in AppConfig.LOG_SENDER_KEYWORDS): self.log("Ожидание завершения отправки логов...", "info"); QTimer.singleShot(3000, lambda: self.wait_for_logs_to_finish(start_time))
        else: self.log("Отправка логов завершена.", "info"); self.is_sending_logs = False; self.check_for_player()
    
    def check_and_launch_player(self):
        if self.is_sending_logs: return
        if self.window_manager.find_first_window_by_title(AppConfig.PLAYER_GAME_LAUNCHER_TITLE) or self.window_manager.find_first_window_by_title("launch"): self.log("Обнаружен процесс запуска/обновления плеера. Ожидание...", "info"); return
        self.log("Плеер не найден, ищу ярлык 'launch'...", "warning"); desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'launch' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: self.log("Найден ярлык плеера. Запускаю...", "info"); os.startfile(shortcut_path); self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
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
            self.window_manager.humanized_click(x, y)
        except Exception as e: self.log(f"Не удалось активировать окно плеера: {e}", "error")
    
    def check_and_launch_opencv_server(self):
        if not self.is_automation_enabled: return
        config = PROJECT_CONFIGS.get("QQ");
        if not config or self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId()): return
        self.log("Сервер OpenCV не найден, ищу ярлык...", "warning"); desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'opencv' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: self.log("Найден ярлык OpenCV. Запускаю...", "info"); os.startfile(shortcut_path); QTimer.singleShot(4000, self.arrange_other_windows)
            else: self.log("Ярлык для OpenCV на рабочем столе не найден.", "error")
        except Exception as e: self.log(f"Ошибка при поиске/запуске ярлыка OpenCV: {e}", "error")
    
    def check_for_recorder(self):
        if self.window_manager.is_process_running(AppConfig.CAMTASIA_PROCESS_NAME):
            if self.timers["recorder_check"].isActive(): self.timers["recorder_check"].stop()
            return
        if self.is_automation_enabled:
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            try:
                shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.CAMTASIA_PROCESS_NAME in f.lower() and f.lower().endswith('.lnk')), None)
                if shortcut_path: self.log("Camtasia не найдена. Запускаю...", "warning"); os.startfile(shortcut_path); self.timers["recorder_check"].stop(); QTimer.singleShot(3000, self.check_for_recorder)
                else: self.log("Ярлык для Camtasia на рабочем столе не найден.", "error")
            except Exception as e: self.log(f"Ошибка при поиске/запуске Camtasia: {e}", "error")
        if not self.timers["recorder_check"].isActive(): self.timers["recorder_check"].start(AppConfig.RECORDER_CHECK_INTERVAL)
    
    def initial_recorder_sync_check(self):
        if self.window_manager.find_first_window_by_title("Recording..."): self.log("Обнаружена активная запись. Перезапускаю для синхронизации...", "warning"); self.stop_recording_session(); QTimer.singleShot(2000, self.check_auto_record_logic)
    
    def toggle_auto_record(self):
        self.is_auto_record_enabled = not self.is_auto_record_enabled
        self.log(f"Автозапись {'включена' if self.is_auto_record_enabled else 'выключена'}.", "info")
        if self.is_auto_record_enabled:
            self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        else:
            self.timers["auto_record"].stop()
        self.auto_record_toggle.setChecked(self.is_auto_record_enabled)
    
    def toggle_automation(self):
        self.is_automation_enabled = not self.is_automation_enabled
        self.log(f"Автоматика {'включена' if self.is_automation_enabled else 'выключена'}.", "info")
        if self.is_automation_enabled:
            self.check_for_player()
            self.check_for_recorder()
        self.automation_toggle.setChecked(self.is_automation_enabled)
    
    def check_auto_record_logic(self):
        if not self.is_auto_record_enabled or not self.current_project or self.is_record_stopping: return
        config = PROJECT_CONFIGS.get(self.current_project);
        if not config: return
        
        try:
            should_be_recording = False
            if self.current_project == "GG":
                # Логика для GG: ClubGG.exe ИЛИ chrome.exe
                should_be_recording = (self.window_manager.is_process_running("ClubGG.exe") or 
                                       self.window_manager.is_process_running("chrome.exe"))
            else: # Логика для QQ и других проектов
                should_be_recording = (self.window_manager.find_windows_by_config(config, "LOBBY", self.winId()) or 
                                       self.window_manager.find_windows_by_config(config, "TABLE", self.winId()))

            is_recording = self.window_manager.find_first_window_by_title("Recording...") is not None
            is_paused = self.window_manager.find_first_window_by_title("Paused...") is not None

            if should_be_recording and not (is_recording or is_paused):
                self.log("Начинаю автозапись...", "info")
                self.start_recording_session()
            elif should_be_recording and is_paused:
                self.log("Запись на паузе. Возобновляю...", "warning")
                self.perform_camtasia_action(win32con.VK_F9, "возобновление записи")
            elif not should_be_recording and (is_recording or is_paused):
                self.log("Активность завершена. Останавливаю запись...", "info")
                self.stop_recording_session()
        except Exception as e:
            self.log(f"Ошибка в логике автозаписи: {e}", "error")

    
    def start_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F9, "старт записи")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        
        self.recording_start_time = time.monotonic()
        self.timers["session"].start(1000)
        
        self.progress_bar.setMaximum(config["SESSION_MAX_DURATION_S"])
        self.progress_bar.setValue(0)
        
        self.sync_ui_state()
    
    def stop_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F10, "остановку записи"); self.timers["session"].stop(); self.recording_start_time = 0; self.is_record_stopping = True; self.timers["record_cooldown"].start(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000); self.sync_ui_state()
    
    def perform_camtasia_action(self, key_code, action_name):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME)
        if not recorder_hwnd:
            self.log(f"Не удалось выполнить '{action_name}': окно Camtasia не найдено.", "error")
            return
        self.log(f"Выполняю '{action_name}' для Camtasia...", "info")
        try:
            self.focus_window(recorder_hwnd)
            time.sleep(0.7)
            if action_name == "старт записи" and not self.window_manager.find_first_window_by_title("Recording..."):
                rect = win32gui.GetWindowRect(recorder_hwnd)
                x, y = rect[0] + 55, rect[1] + 80
                self.window_manager.humanized_click(x, y)
                self.log("Full Screen был нажат.", "info")
                time.sleep(0.5)
            self.window_manager.press_key(key_code)
            QTimer.singleShot(500, self.position_recorder_window)
        except Exception as e:
            self.log(f"Ошибка при взаимодействии с Camtasia: {e}", "error")
    
    def update_session_progress(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "SESSION_MAX_DURATION_S" not in config: return
        
        elapsed = time.monotonic() - self.recording_start_time
        self.progress_bar.setValue(elapsed)

        if elapsed >= config["SESSION_MAX_DURATION_S"]:
            self.progress_bar_label.setText(AppConfig.MSG_LIMIT_REACHED)
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL + f" color: {ColorPalette.RED}; font-weight: bold;")
            
            # Логика перезапуска остается, но выносим ее за пределы обновления UI
            QTimer.singleShot(100, self.handle_session_limit_reached)
        else:
            remaining_s = max(0, config["SESSION_MAX_DURATION_S"] - elapsed)
            self.progress_bar_label.setText(AppConfig.MSG_PROGRESS_LABEL.format(time.strftime('%H:%M:%S', time.gmtime(remaining_s))))
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)

    def handle_session_limit_reached(self):
        if self.recording_start_time == 0: return # Предотвращаем двойной вызов
        config = PROJECT_CONFIGS.get(self.current_project)
        self.log(f"{config['SESSION_MAX_DURATION_S']/3600:.0f} часа записи истекли. Перезапуск...", "info")
        self.stop_recording_session()
        QTimer.singleShot(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000 + 1000, self.check_auto_record_logic)
    
    def check_for_new_tables(self):
        if not self.is_automation_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project);
        if not config: return
        if self.current_project == "GG" and not self.window_manager.is_process_running("ClubGG.exe"): self.last_table_count = 0; return
        current_tables = self.window_manager.find_windows_by_config(config, "TABLE", self.winId()); current_count = len(current_tables)
        if current_count != self.last_table_count: self.log(f"Изменилось количество столов: {self.last_table_count} -> {current_count}. Перерасстановка...", "info"); QTimer.singleShot(500, self.arrange_tables)
        self.last_table_count = current_count
    
    def arrange_tables(self):
        if not self.current_project: self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error"); return
        config = PROJECT_CONFIGS.get(self.current_project);
        if not config or "TABLE" not in config: self.log(f"Нет конфига столов для {self.current_project}.", "warning"); return
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        
        if found_windows:
            titles = [win32gui.GetWindowText(hwnd) for hwnd in found_windows]
            logging.info(f"Найдены окна для расстановки ({len(titles)}): {titles}")
        else:
            self.log(AppConfig.MSG_ARRANGE_TABLES_NOT_FOUND, "warning"); return
            
        if self.current_project == "QQ" and len(found_windows) > 4:
            self.arrange_dynamic_qq_tables(found_windows, config)
            return

        slots_key = "TABLE_SLOTS_5" if self.current_project == "QQ" and len(found_windows) >= 5 else "TABLE_SLOTS"; SLOTS = config[slots_key]; arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS) or not win32gui.IsWindow(hwnd): continue
            x, y = SLOTS[i];_ = win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) if win32gui.IsIconic(hwnd) else None
            try:
                if self.current_project == "GG": win32gui.MoveWindow(hwnd, x, y, config["TABLE"]["W"], config["TABLE"]["H"], True)
                else: rect = win32gui.GetWindowRect(hwnd); win32gui.MoveWindow(hwnd, x, y, rect[2] - rect[0], rect[3] - rect[1], True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось разместить стол {i+1}: {e}", "error")
        if arranged_count > 0: self.log(f"Расставлено столов: {arranged_count}", "info")

    def arrange_dynamic_qq_tables(self, found_windows, config):
        """Динамически расставляет столы QQ, если их больше 4."""
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
            try:
                win32gui.MoveWindow(hwnd, x, y, new_width, new_height, True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось динамически разместить стол {i+1}: {e}", "error")

        if arranged_count > 0: self.log(f"Динамически расставлено столов: {arranged_count}", "info")
    
    def arrange_other_windows(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            self.log(f"Нет конфига для {self.current_project}.", "warning")
            return
        self.position_player_window(config)
        self.position_recorder_window()
        if self.current_project == "GG":
            self.position_lobby_window(config)
        elif self.current_project == "QQ":
            self.position_cv_server_window(config)
        self.log("Системные окна расставлены.", "info")
    
    def position_window(self, hwnd, x, y, w, h, log_fail):
        if hwnd and win32gui.IsWindow(hwnd): win32gui.MoveWindow(hwnd, x, y, w, h, True)
        else: self.log(log_fail, "warning")
    
    def position_player_window(self, config): player_config = config.get("PLAYER", {});_ = self.position_window(self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE), player_config["X"], player_config["Y"], player_config["W"], player_config["H"], "Плеер не найден.") if player_config else None
    
    def position_lobby_window(self, config): lobbies = self.window_manager.find_windows_by_config(config, "LOBBY", self.winId()); cfg = config.get("LOBBY", {});_ = self.position_window(lobbies[0] if lobbies else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "Лобби не найдено.") if cfg else None
    
    def position_cv_server_window(self, config): cv_windows = self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId()); cfg = config.get("CV_SERVER", {});_ = self.position_window(cv_windows[0] if cv_windows else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "CV Сервер не найден.") if cfg else None
    
    def position_recorder_window(self):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME)
        if recorder_hwnd and win32gui.IsWindow(recorder_hwnd):
            try:
                # ИЗМЕНЕНО: Разворачиваем окно, если оно свернуто, перед позиционированием
                if win32gui.IsIconic(recorder_hwnd):
                    win32gui.ShowWindow(recorder_hwnd, win32con.SW_RESTORE)
                    time.sleep(0.2) # Даем окну время на восстановление

                screen_rect = self.get_current_screen().availableGeometry()
                rect = win32gui.GetWindowRect(recorder_hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                x, y = screen_rect.left() + (screen_rect.width() - w) // 2, screen_rect.bottom() - h
                win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
            except Exception as e: self.log(f"Ошибка позиционирования Camtasia: {e}", "error")
    
    def minimize_injector_window(self):
        injector_hwnd = self.window_manager.find_first_window_by_title(AppConfig.INJECTOR_WINDOW_TITLE)
        if injector_hwnd and win32gui.IsWindow(injector_hwnd):
            try: win32gui.ShowWindow(injector_hwnd, win32con.SW_MINIMIZE); self.log("Окно 'injector' свернуто.", "info")
            except Exception as e: self.log(f"Не удалось свернуть окно 'injector': {e}", "error")
    
    def focus_window(self, hwnd):
        try:
            if self.shell: self.shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd); time.sleep(0.2)
        except Exception as e: logging.error(f"Не удалось сфокусировать окно {hwnd}: {e}")
    
    def perform_special_clicks(self):
        if self.current_project != "GG": return
        config = PROJECT_CONFIGS.get("GG");
        if not config: return
        self.log("Выполняю команду SIT-OUT для столов...", "info"); found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        if not found_windows: self.log("Столы для выполнения команды не найдены.", "warning"); return
        click_count = 0
        for hwnd in found_windows:
            if not win32gui.IsWindow(hwnd): continue
            try: 
                self.focus_window(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                x, y = rect[0] + 25, rect[1] + 410
                self.window_manager.humanized_click(x, y)
                time.sleep(0.1)
                click_count += 1
            except Exception as e: logging.error(f"Не удалось выполнить клик для окна {hwnd}: {e}")
        if click_count > 0: self.log(f"Команда SIT-OUT выполнена для {click_count} столов.", "info")
    
    def position_window_gg_default(self):
        try: screen = self.get_current_screen(); geo = screen.availableGeometry(); margin = AppConfig.WINDOW_MARGIN; self.move(geo.left() + margin, geo.top() + margin)
        except Exception as e: logging.error(f"Could not position window for GG: {e}")
    
    def position_window_top_right(self):
        try: screen = self.get_current_screen(); geo = screen.availableGeometry(); margin = AppConfig.WINDOW_MARGIN; self.move(geo.right() - self.frameGeometry().width() - margin, geo.top() + margin)
        except Exception as e: logging.error(f"Could not position window top-right: {e}")
    
    def position_window_default(self):
        try: screen = self.get_current_screen(); geo = screen.availableGeometry(); margin = AppConfig.WINDOW_MARGIN; self.move(geo.left() + margin, geo.bottom() - self.frameGeometry().height() - margin)
        except Exception as e: logging.error(f"Could not position window default: {e}")

# ===================================================================
# 6. ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ===================================================================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
