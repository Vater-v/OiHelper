# -*- coding: utf-8 -*-
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

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, 
    QProgressBar, QGridLayout, QHBoxLayout, QGraphicsDropShadowEffect
)
from PyQt6.QtCore import QObject, QPropertyAnimation, Qt, QTimer, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QPalette, QColor, QPixmap, QFont, QPainter
from PyQt6.QtSvg import QSvgRenderer

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class SvgIcon:
    TABLES = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="3" y1="9" x2="21" y2="9"></line><line x1="3" y1="15" x2="21" y2="15"></line><line x1="9" y1="3" x2="9" y2="21"></line><line x1="15" y1="3" x2="15" y2="21"></line></svg>"""
    OTHER = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>"""
    RECORD_ON = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="3" fill="red" stroke="none"></circle></svg>"""
    RECORD_OFF = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><rect x="9" y="9" width="6" height="6"></rect></svg>"""
    AUTOMATION_ON = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z"></path><path d="M12 8v4l2 2"></path></svg>"""
    AUTOMATION_OFF = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="1" y1="1" x2="23" y2="23"></line></svg>"""
    NOTIF_INFO = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>"""
    NOTIF_WARN = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>"""
    NOTIF_ERROR = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>"""
    CLICK = """<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3l7.07 16.97 2.51-7.39 7.39-2.51L3 3z"></path><path d="M13 13l6 6"></path></svg>"""

class AppConfig:
    # ИЗМЕНЕНО: Версия увеличена до 2.61
    CURRENT_VERSION = "v2.61"
    GITHUB_REPO = "Vater-v/OiHelper"
    ASSET_NAME = "OiHelper.zip"
    ICON_PATH = 'icon.ico'
    TELEGRAM_BOT_TOKEN = '8087291675:AAHC-s2MKTCqmd4J-sW3iwgqvzj3YhBMhs' 
    TELEGRAM_CHAT_ID = '747883453' 
    PLAYER_LAUNCHER_TITLE = "Holdem"
    PLAYER_GAME_LAUNCHER_TITLE = "Game"
    TELEGRAM_REPORT_LEVEL = 'all'
    LOG_SENDER_KEYWORDS = ["endsess", "logbot"]
    PLAYER_RELAUNCH_DELAY_S = 15
    DEFAULT_WIDTH = 440
    DEFAULT_HEIGHT = 310 
    GG_UI_WIDTH = 620
    GG_UI_HEIGHT = 160
    FLASHING_HEIGHT_V = 370
    FLASHING_HEIGHT_H = 220
    WINDOW_MARGIN = 20 
    APP_TITLE = f"OiHelper {CURRENT_VERSION}"
    MSG_PLAYER_SEARCH = "Поиск плеера..."
    MSG_PROJECT_UNDEFINED = "Проект не определен"
    MSG_PLAYER_NOT_FOUND = "Плеер не найден"
    MSG_PRESS_START_ON_PLAYER = "Нажмите 'Start' на плеере!"
    MSG_RUN_PLAYER = "Запустите плеер"
    MSG_AUTORECORD_ON = "Автозапись: ВКЛ"
    MSG_AUTORECORD_OFF = "Автозапись: ВЫКЛ"
    MSG_AUTOMATION_ON = "Автоматика: ВКЛ"
    MSG_AUTOMATION_OFF = "Автоматика: ВЫКЛ"
    MSG_STOP_FLASHING = "Остановить мигание"
    MSG_ARRANGE_TABLES = "Расставить столы"
    MSG_ARRANGE_OTHER = "Прочее"
    MSG_CLICK_COMMAND = "SIT-OUT"
    MSG_PROGRESS_LABEL = "Перезапуск через: {}"
    MSG_UPDATE_CHECK = "Проверка обновлений..."
    MSG_UPTIME_WARNING = "Компьютер не перезагружался более 5 дней. Рекомендуется перезагрузка для стабильной работы."
    MSG_ARRANGE_TABLES_NOT_FOUND = "Столы для расстановки не найдены."
    AUTO_RECORD_INTERVAL = 3000
    AUTO_ARRANGE_INTERVAL = 2500
    PLAYER_CHECK_INTERVAL = 5000 
    RECORDER_CHECK_INTERVAL = 10000
    FLASH_INTERVAL = 500
    NOTIFICATION_DURATION = 7000
    NOTIFICATION_FADE_DURATION = 300
    CAMTASIA_ACTION_DELAY = 0.2
    PLAYER_START_DELAY = 4000
    PLAYER_AUTOSTART_INTERVAL = 7000

PROJECT_CONFIGS = {
    "GG": { 
        "TABLE": { "FIND_METHOD": "RATIO", "W": 557, "H": 424, "TOLERANCE": 0.035 }, 
        "LOBBY": { "FIND_METHOD": "RATIO", "W": 333, "H": 623, "TOLERANCE": 0.07, "X": 1580, "Y": 140 }, 
        "PLAYER": { "W": 700, "H": 365, "X": 1385, "Y": 0 }, 
        "TABLE_SLOTS": [(-5, 0), (271, 423), (816, 0), (1086, 423)], 
        "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], 
        "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe"], 
        "SESSION_MAX_DURATION_S": 4 * 3600, 
        "SESSION_WARN_TIME_S": 3.5 * 3600, 
        "ARRANGE_MINIMIZED_TABLES": False 
    },
    "QQ": { 
        "TABLE": { "FIND_METHOD": "TITLE_AND_SIZE", "TITLE": "QQPK", "W": 400, "H": 700, "TOLERANCE": 2 }, 
        "CV_SERVER": { "FIND_METHOD": "TITLE", "TITLE": "OpenCv", "X": 1789, "Y": 367, "W": 993, "H": 605 }, 
        "PLAYER": { "X": 1418, "Y": 942, "W": 724, "H": 370 }, 
        "TABLE_SLOTS": [(0, 0), (401, 0), (802, 0), (1203, 0)], 
        "TABLE_SLOTS_5": [(0, 0), (346, 0), (692, 0), (1038, 0), (1384, 0)], 
        "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], 
        "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "OpenCvServer.exe"], 
        "SESSION_MAX_DURATION_S": 3 * 3600, 
        "SESSION_WARN_TIME_S": -1, 
        "ARRANGE_MINIMIZED_TABLES": True 
    }
}
# ===================================================================
# 2. СИСТЕМА ИНТЕРФЕЙСА (УВЕДОМЛЕНИЯ, ИНДИКАТОРЫ)
# ===================================================================
class Notification(QWidget):
    """Всплывающее уведомление с кастомным дизайном."""
    closed = pyqtSignal(QWidget)
    COLORS = {"info": ("#3498DB", "#2E86C1"), "warning": ("#F39C12", "#D35400"), "error": ("#E74C3C", "#C0392B")}
    ICONS = {"info": SvgIcon.NOTIF_INFO, "warning": SvgIcon.NOTIF_WARN, "error": SvgIcon.NOTIF_ERROR}

    def __init__(self, message, message_type):
        super().__init__()
        self.is_closing = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        
        container = QWidget(self)
        stop1, stop2 = self.COLORS.get(message_type, self.COLORS["info"])
        container.setStyleSheet(f"QWidget {{ background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {stop1}, stop:1 {stop2}); border-radius: 10px; border: 1px solid rgba(255, 255, 255, 0.1); }}")
        
        shadow_layout = QVBoxLayout(self)
        shadow_layout.setContentsMargins(5, 5, 5, 5)
        shadow_layout.addWidget(container)
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 160))
        shadow.setOffset(0, 3)
        container.setGraphicsEffect(shadow)
        
        layout = QHBoxLayout(container)
        layout.setContentsMargins(15, 15, 20, 15)
        layout.setSpacing(15)
        
        icon_label = QLabel()
        icon_svg = self.ICONS.get(message_type, self.ICONS["info"])
        renderer = QSvgRenderer(icon_svg.encode('utf-8'))
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        icon_label.setPixmap(pixmap)
        layout.addWidget(icon_label)
        
        text_label = QLabel(message)
        text_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        text_label.setWordWrap(True)
        text_label.setStyleSheet("background: transparent; border: none; color: white;")
        layout.addWidget(text_label, 1)
        
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(AppConfig.NOTIFICATION_FADE_DURATION)
        QTimer.singleShot(AppConfig.NOTIFICATION_DURATION, self.hide_animation)

    def show_animation(self):
        self.setWindowOpacity(0.0)
        self.show()
        self.animation.setStartValue(0.0)
        self.animation.setEndValue(1.0)
        self.animation.start()

    def hide_animation(self):
        if self.is_closing: return
        self.is_closing = True
        self.animation.setStartValue(self.windowOpacity())
        self.animation.setEndValue(0.0)
        self.animation.finished.connect(self.close)
        self.animation.start()

    def closeEvent(self, event):
        self.closed.emit(self)
        super().closeEvent(event)

class NotificationManager(QObject):
    """Управляет очередью и позиционированием уведомлений."""
    def __init__(self):
        super().__init__()
        self.notifications = []

    def show(self, message, message_type):
        logging.info(f"Уведомление [{message_type}]: {message}")
        if len(self.notifications) >= 5:
            self.notifications[0].hide_animation()
        
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
        
        margin = 20
        total_height = 0
        for n in reversed(self.notifications):
            n.adjustSize()
            width, height = n.width(), n.height()
            x = screen_geo.right() - width - margin
            y = screen_geo.bottom() - height - margin - total_height
            n.move(x, y)
            total_height += height + 10

class ClickIndicator(QWidget):
    """Визуальный индикатор клика мыши. Появляется в месте клика и исчезает."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(30, 30)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(255, 0, 0, 150))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(self.rect())

    def show_at(self, x, y):
        self.move(x - self.width() // 2, y - self.height() // 2)
        self.show()
        self.timer.start(200)

# ===================================================================
# 3. ЛОГИКА РАБОТЫ С ОКНАМИ (WINAPI)
# ===================================================================
class WindowManager(QObject):
    """Инкапсулирует всю работу с WinAPI: поиск окон, клики, нажатия клавиш."""
    log_request = pyqtSignal(str, str)
    click_visual_request = pyqtSignal(int, int)

    def find_first_window_by_title(self, text_in_title, exact_match=False):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if (exact_match and text_in_title == title) or \
                       (not exact_match and text_in_title.lower() in title.lower()):
                        hwnds.append(hwnd)
                except Exception as e:
                    logging.warning(f"Не удалось получить заголовок для hwnd {hwnd}: {e}")
        try:
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            logging.error(f"Критическая ошибка в EnumWindows при поиске по заголовку: {e}")
        return hwnds[0] if hwnds else None
    
    def is_process_running(self, process_name_to_find):
        """Проверяет, запущен ли процесс с указанным именем."""
        pids = win32process.EnumProcesses()
        for pid in pids:
            if pid == 0: continue
            try:
                h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                if h_process:
                    process_name = win32process.GetModuleFileNameEx(h_process, 0)
                    win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in os.path.basename(process_name).lower():
                        return True
            except Exception:
                continue
        return False

    def find_first_window_by_process_name(self, process_name_to_find):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    process_name = win32process.GetModuleFileNameEx(h_process, 0)
                    win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in os.path.basename(process_name).lower():
                        hwnds.append(hwnd)
                except Exception:
                    pass
        try:
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            logging.error(f"Критическая ошибка в EnumWindows при поиске по процессу: {e}")
        return hwnds[0] if hwnds else None

    def find_windows_by_config(self, config, config_key, main_window_hwnd):
        window_config = config.get(config_key, {})
        if not window_config: return []
        
        find_method = window_config.get("FIND_METHOD")
        EXCLUDED_TITLES = config.get("EXCLUDED_TITLES", [])
        EXCLUDED_PROCESSES = config.get("EXCLUDED_PROCESSES", [])
        arrange_minimized = config.get("ARRANGE_MINIMIZED_TABLES", False) and config_key == "TABLE"
        
        found_windows = []

        def enum_windows_callback(hwnd, _):
            if hwnd == main_window_hwnd: return True
            if not win32gui.IsWindowVisible(hwnd): return True
            if not arrange_minimized and win32gui.IsIconic(hwnd): return True
            
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
                    ratio = w / h
                    TARGET_RATIO = window_config["W"] / window_config["H"]
                    TOLERANCE = window_config["TOLERANCE"]
                    if TARGET_RATIO * (1 - TOLERANCE) <= ratio <= TARGET_RATIO * (1 + TOLERANCE):
                        match = True
                elif find_method == "TITLE_AND_SIZE":
                    TOLERANCE = window_config["TOLERANCE"]
                    if window_config["TITLE"].lower() in title.lower() and \
                       abs(w - window_config["W"]) <= TOLERANCE and \
                       abs(h - window_config["H"]) <= TOLERANCE:
                        match = True
                elif find_method == "TITLE":
                    if window_config["TITLE"].lower() in title.lower():
                        match = True
                
                if match:
                    found_windows.append(hwnd)

            except Exception as e:
                logging.debug(f"Пропущено окно {hwnd} (не удалось получить информацию): {e}")
            
            return True

        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            logging.error(f"Критическая ошибка в EnumWindows при поиске по конфигу: {e}")
        
        return found_windows

    def press_key(self, key_code):
        try:
            win32api.keybd_event(key_code, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e:
            self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

    def click_at_pos(self, x, y):
        try:
            self.click_visual_request.emit(x, y)
            win32api.SetCursorPos((x, y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            time.sleep(0.05)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
        except Exception as e:
            self.log_request.emit(f"Ошибка клика по координатам ({x},{y}): {e}", "error")

# ===================================================================
# 4. ВНЕШНИЕ СЕРВИСЫ (ТЕЛЕГРАМ, ОБНОВЛЕНИЯ)
# ===================================================================
class TelegramNotifier(QObject):
    """Отправляет сообщения в Telegram в отдельном потоке."""
    def __init__(self, token, chat_id):
        super().__init__()
        self.token = token
        self.chat_id = chat_id
        self.api_url = f"https://api.telegram.org/bot{self.token}/sendMessage"

    def send_message(self, message):
        if not self.token or not self.chat_id:
            logging.warning("Токен или ID чата Telegram не настроены.")
            return
        payload = {'chat_id': self.chat_id, 'text': message}
        threading.Thread(target=self._send_request, args=(payload,), daemon=True).start()

    def _send_request(self, payload):
        try:
            response = requests.post(self.api_url, data=payload, timeout=10)
            response.raise_for_status()
            logging.info("Сообщение в Telegram успешно отправлено.")
        except requests.RequestException as e:
            logging.error(f"Не удалось отправить сообщение в Telegram: {e}")

class UpdateManager(QObject):
    """Проверяет и применяет обновления с GitHub."""
    log_request = pyqtSignal(str, str)
    check_finished = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.update_info = {}

    def check_for_updates(self):
        self.log_request.emit("Проверка обновлений...", "info")
        try:
            api_url = f"https://api.github.com/repos/{AppConfig.GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release.get("tag_name")

            if latest_version and latest_version > AppConfig.CURRENT_VERSION:
                self.log_request.emit(f"Доступна новая версия: {latest_version}. Обновление...", "info")
                self.update_info = latest_release
                threading.Thread(target=self.apply_update, daemon=True).start()
            else:
                self.log_request.emit("Вы используете последнюю версию.", "info")
                self.check_finished.emit()
        except requests.RequestException as e:
            self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error")
            self.check_finished.emit()
        except Exception as e:
            logging.error(f"Неожиданная ошибка при проверке обновлений: {e}")
            self.check_finished.emit()

    def apply_update(self):
        assets = self.update_info.get("assets", [])
        download_url = None
        for asset in assets:
            if asset["name"] == AppConfig.ASSET_NAME:
                download_url = asset["browser_download_url"]
                break
        
        if not download_url:
            self.log_request.emit("Не удалось найти ZIP-архив в релизе.", "error")
            return
        
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url):
        try:
            self.log_request.emit("Скачивание обновления...", "info")
            update_zip_name = "update.zip"
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            self.log_request.emit("Распаковка архива...", "info")
            update_folder = "update_temp"
            if os.path.isdir(update_folder):
                import shutil
                shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref:
                zip_ref.extractall(update_folder)
            
            self.log_request.emit("Обновление скачано. Перезапуск...", "info")
            updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            exe_name = os.path.basename(current_exe_path)
            
            script_content = f'@echo off\nchcp 65001 > NUL\necho Waiting for OiHelper to close...\ntimeout /t 2 /nobreak > NUL\ntaskkill /pid {os.getpid()} /f > NUL\necho Waiting for process to terminate...\ntimeout /t 3 /nobreak > NUL\necho Moving new files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\nrd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\ndel "{current_dir}\\{update_zip_name}"\necho Starting new version...\nstart "" "{exe_name}"\n(goto) 2>nul & del "%~f0"'
            
            with open(updater_script_path, "w", encoding="cp866") as f:
                f.write(script_content)
            
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()

        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error")
            logging.error(f"Update error: {e}")

# ===================================================================
# 5. ГЛАВНОЕ ОКНО ПРИЛОЖЕНИЯ
# ===================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.notification_manager = NotificationManager()
        self.window_manager = WindowManager()
        self.update_manager = UpdateManager()
        self.telegram_notifier = TelegramNotifier(AppConfig.TELEGRAM_BOT_TOKEN, AppConfig.TELEGRAM_CHAT_ID)
        self.click_indicator = ClickIndicator()
        self.current_project = None
        self.is_auto_record_enabled = True
        self.is_automation_enabled = True
        self.is_flashing = False
        self.flash_state = False
        self.last_table_count = 0
        self.recording_start_time = 0
        self.flashing_disabled_for_session = False
        self.is_sending_logs = False
        self.player_was_open = False
        self.create_widgets()
        self.connect_signals()
        self.rebuild_ui(None)
        self.init_timers()
        self.init_startup_checks()

    def log(self, message, message_type):
        self.notification_manager.show(message, message_type)
        if message_type == 'error':
            self.telegram_notifier.send_message(f"OiHelper Критическая ошибка: {message}")

    def get_button_style(self, color1, color2, hover_color, pressed_color):
        base_style = "font-size: 16px; font-weight: bold; color: white; border-radius: 8px; padding: 5px 15px 5px 15px; border: 1px solid rgba(0,0,0,0.3); text-align: center;"
        template = "QPushButton {{ {base} background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {color1}, stop:1 {color2}); }} QPushButton:hover {{ background-color: {hover_color}; }} QPushButton:disabled {{ background-color: #5D6D7E; color: #BDC3C7; border: 1px solid #444;}} QPushButton:pressed {{ background-color: {pressed_color}; border: 1px solid rgba(255,255,255,0.2); }}"
        return template.format(base=base_style, color1=color1, color2=color2, hover_color=hover_color, pressed_color=pressed_color)

    def get_svg_icon(self, svg_data):
        renderer = QSvgRenderer(svg_data.encode('utf-8'))
        pixmap = QPixmap(24, 24)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)

    def create_widgets(self):
        self.project_label = QLabel(AppConfig.MSG_PLAYER_SEARCH)
        self.project_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.project_label.setStyleSheet("font-size: 28px; font-weight: bold; color: white; background: transparent;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("QProgressBar { border: 1px solid #555; border-radius: 5px; text-align: center; height: 10px; background-color: #222;} QProgressBar::chunk { background-color: #05B8CC; border-radius: 5px;}")
        self.progress_bar_label = QLabel()
        self.progress_bar_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar_label.setStyleSheet("font-size: 13px; color: #aaa; background: transparent;")
        self.progress_bar_label.setVisible(False)
        self.stop_flash_button = QPushButton(AppConfig.MSG_STOP_FLASHING)
        self.stop_flash_button.setVisible(False)
        self.auto_record_toggle_button = QPushButton()
        self.automation_toggle_button = QPushButton()
        self.arrange_tables_button = QPushButton(AppConfig.MSG_ARRANGE_TABLES)
        self.arrange_other_button = QPushButton(AppConfig.MSG_ARRANGE_OTHER)
        self.click_command_button = QPushButton(AppConfig.MSG_CLICK_COMMAND)
        self.arrange_tables_button.setStyleSheet(self.get_button_style("#007BFF", "#0056b3", "#006fe6", "#004a99"))
        self.arrange_other_button.setStyleSheet(self.get_button_style("#6c757d", "#5a6268", "#7c858d", "#4a5258"))
        self.stop_flash_button.setStyleSheet(self.get_button_style("#ffc107", "#e0a800", "#ffca2b", "#d39e00"))
        self.click_command_button.setStyleSheet(self.get_button_style("#17a2b8", "#138496", "#19b4cc", "#117a8b"))
        self.arrange_tables_button.setIcon(self.get_svg_icon(SvgIcon.TABLES))
        self.arrange_other_button.setIcon(self.get_svg_icon(SvgIcon.OTHER))
        self.click_command_button.setIcon(self.get_svg_icon(SvgIcon.CLICK))
        icon_size = QSize(20, 20)
        self.auto_record_toggle_button.setIconSize(icon_size)
        self.automation_toggle_button.setIconSize(icon_size)
        self.arrange_tables_button.setIconSize(icon_size)
        self.arrange_other_button.setIconSize(icon_size)
        self.click_command_button.setIconSize(icon_size)
        for btn in [self.auto_record_toggle_button, self.automation_toggle_button, self.arrange_tables_button, self.arrange_other_button, self.stop_flash_button, self.click_command_button]:
            btn.setMinimumHeight(42)
        self.update_auto_record_button_style()
        self.update_automation_button_style()

    def rebuild_ui(self, project_name):
        container = QWidget()
        if project_name == "GG":
            main_layout = QHBoxLayout(container)
            main_layout.setContentsMargins(15, 10, 15, 10)
            main_layout.setSpacing(15)
            left_vbox = QVBoxLayout()
            left_vbox.addWidget(self.arrange_tables_button)
            left_vbox.addWidget(self.arrange_other_button)
            left_vbox.addWidget(self.click_command_button)
            center_vbox = QVBoxLayout()
            center_vbox.addWidget(self.project_label)
            center_vbox.addWidget(self.progress_bar)
            center_vbox.addWidget(self.progress_bar_label)
            center_vbox.addWidget(self.stop_flash_button)
            center_vbox.addStretch(1)
            right_vbox = QVBoxLayout()
            right_vbox.addWidget(self.automation_toggle_button)
            right_vbox.addWidget(self.auto_record_toggle_button)
            right_vbox.addStretch(1)
            main_layout.addLayout(left_vbox, 1)
            main_layout.addLayout(center_vbox, 2)
            main_layout.addLayout(right_vbox, 1)
            self.setFixedSize(AppConfig.GG_UI_WIDTH, AppConfig.GG_UI_HEIGHT)
        else:
            main_layout = QVBoxLayout(container)
            main_layout.setContentsMargins(15, 10, 15, 10)
            main_layout.setSpacing(5)
            main_layout.addWidget(self.project_label)
            main_layout.addWidget(self.progress_bar)
            main_layout.addWidget(self.progress_bar_label)
            main_layout.addWidget(self.stop_flash_button)
            main_layout.addStretch(1)
            button_grid = QGridLayout()
            button_grid.setSpacing(10)
            button_grid.addWidget(self.automation_toggle_button, 0, 0, 1, 2)
            button_grid.addWidget(self.auto_record_toggle_button, 1, 0, 1, 2)
            button_grid.addWidget(self.arrange_tables_button, 2, 0)
            button_grid.addWidget(self.arrange_other_button, 2, 1)
            main_layout.addLayout(button_grid)
            self.setFixedSize(AppConfig.DEFAULT_WIDTH, AppConfig.DEFAULT_HEIGHT)
        self.setCentralWidget(container)
        self.setStyleSheet("QMainWindow { background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4c4c4c, stop:1 #2c2c2c); font-family: 'Segoe UI'; }")

    def connect_signals(self):
        self.window_manager.log_request.connect(self.log)
        self.window_manager.click_visual_request.connect(self.click_indicator.show_at)
        self.update_manager.log_request.connect(self.log)
        self.update_manager.check_finished.connect(self.start_main_logic)

        self.stop_flash_button.clicked.connect(self.stop_flashing)
        self.auto_record_toggle_button.clicked.connect(self.toggle_auto_record)
        self.automation_toggle_button.clicked.connect(self.toggle_automation)
        self.arrange_tables_button.clicked.connect(self.arrange_tables)
        self.arrange_other_button.clicked.connect(self.arrange_other_windows)
        self.click_command_button.clicked.connect(self.perform_special_clicks)

    def update_auto_record_button_style(self):
        if self.is_auto_record_enabled:
            self.auto_record_toggle_button.setText(AppConfig.MSG_AUTORECORD_ON)
            style = self.get_button_style("#28a745", "#218838", "#32b14f", "#1e7e34")
            self.auto_record_toggle_button.setIcon(self.get_svg_icon(SvgIcon.RECORD_ON))
        else:
            self.auto_record_toggle_button.setText(AppConfig.MSG_AUTORECORD_OFF)
            style = self.get_button_style("#dc3545", "#c82333", "#e04b59", "#b52b39")
            self.auto_record_toggle_button.setIcon(self.get_svg_icon(SvgIcon.RECORD_OFF))
        self.auto_record_toggle_button.setStyleSheet(style)

    def update_automation_button_style(self):
        if self.is_automation_enabled:
            self.automation_toggle_button.setText(AppConfig.MSG_AUTOMATION_ON)
            style = self.get_button_style("#28a745", "#218838", "#32b14f", "#1e7e34")
            self.automation_toggle_button.setIcon(self.get_svg_icon(SvgIcon.AUTOMATION_ON))
        else:
            self.automation_toggle_button.setText(AppConfig.MSG_AUTOMATION_OFF)
            style = self.get_button_style("#dc3545", "#c82333", "#e04b59", "#b52b39")
            self.automation_toggle_button.setIcon(self.get_svg_icon(SvgIcon.AUTOMATION_OFF))
        self.automation_toggle_button.setStyleSheet(style)
    
    def init_timers(self):
        self.player_check_timer = QTimer(self)
        self.player_check_timer.timeout.connect(self.check_for_player)
        self.recorder_check_timer = QTimer(self)
        self.recorder_check_timer.timeout.connect(self.check_for_recorder)
        self.auto_record_timer = QTimer(self)
        self.auto_record_timer.timeout.connect(self.check_auto_record_logic)
        self.auto_arrange_timer = QTimer(self)
        self.auto_arrange_timer.timeout.connect(self.check_for_new_tables)
        self.session_timer = QTimer(self)
        self.session_timer.timeout.connect(self.update_session_progress)
        self.flash_timer = QTimer(self)
        self.flash_timer.timeout.connect(self.toggle_window_flash)
        self.player_start_attempt_timer = QTimer(self)
        self.player_start_attempt_timer.timeout.connect(self.attempt_player_start_click)

    def init_startup_checks(self):
        self.setWindowTitle(AppConfig.APP_TITLE)
        if os.path.exists(AppConfig.ICON_PATH):
            self.setWindowIcon(QIcon(AppConfig.ICON_PATH))
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        self.project_label.setText(AppConfig.MSG_UPDATE_CHECK)
        self.arrange_tables_button.setEnabled(False)
        self.arrange_other_button.setEnabled(False)
        self.auto_record_toggle_button.setEnabled(False)
        self.automation_toggle_button.setEnabled(False)
        self.click_command_button.setEnabled(False)

        threading.Thread(target=self.update_manager.check_for_updates, daemon=True).start()

    def start_main_logic(self):
        """Запускает основную логику приложения после проверки обновлений."""
        self.arrange_tables_button.setEnabled(True)
        self.arrange_other_button.setEnabled(True)
        self.click_command_button.setEnabled(True)
        self.automation_toggle_button.setEnabled(True)
        self.check_for_player()
        self.check_for_recorder()
        self.auto_record_timer.start(AppConfig.AUTO_RECORD_INTERVAL)
        QTimer.singleShot(1000, self.initial_recorder_sync_check)
        if AppConfig.TELEGRAM_REPORT_LEVEL == 'all':
            self.telegram_notifier.send_message(f"OiHelper {AppConfig.CURRENT_VERSION} запущен.")
        self.check_system_uptime()

    def check_system_uptime(self):
        try:
            uptime_ms = ctypes.windll.kernel32.GetTickCount64()
            days = uptime_ms / (1000 * 60 * 60 * 24)
            if days > 5:
                self.log(AppConfig.MSG_UPTIME_WARNING, "warning")
        except Exception as e:
            logging.error(f"Не удалось проверить время работы системы: {e}")
    
    def check_for_player(self):
        if self.is_sending_logs:
            return

        player_hwnd = self.window_manager.find_first_window_by_title("holdem")

        if not player_hwnd:
            if self.player_start_attempt_timer.isActive():
                self.player_start_attempt_timer.stop()
            
            if self.current_project is not None:
                self.current_project = None
                self.rebuild_ui(None)
                self.position_window_default() 

            if self.player_was_open:
                self.player_was_open = False
                self.handle_player_close()
                return

            self.project_label.setText(AppConfig.MSG_PLAYER_NOT_FOUND)
            if self.is_automation_enabled:
                self.check_and_launch_player()
        else:
            self.player_was_open = True
            is_project_running = False
            project_name = None
            try:
                title = win32gui.GetWindowText(player_hwnd)
                project_map = {"QQPoker": "QQ", "ClubGG": "GG"}
                for full_name, short_name in project_map.items():
                    if f"[{full_name}]" in title:
                        project_name = short_name
                        is_project_running = True
                        break
            except Exception:
                is_project_running = False

            if is_project_running:
                if self.player_start_attempt_timer.isActive():
                    self.player_start_attempt_timer.stop()
                    self.log("Авто-старт плеера успешно завершен.", "info")
                
                if project_name != self.current_project:
                    self.current_project = project_name
                    self.rebuild_ui(project_name)
                    self.last_table_count = 0
                    self.project_label.setText(f"{project_name} - Панель")
                    self.auto_arrange_timer.start(AppConfig.AUTO_ARRANGE_INTERVAL)
                    
                    if project_name == "QQ":
                        self.position_window_top_right()
                        if self.is_automation_enabled:
                            self.check_and_launch_opencv_server()
                    elif project_name == "GG":
                        self.position_window_gg_default()
                        if self.is_automation_enabled:
                            self.minimize_injector_window()
                    
                    if self.is_automation_enabled:
                        self.arrange_other_windows()
            else:
                if self.current_project is not None:
                    self.current_project = None
                    self.rebuild_ui(None)
                    self.position_window_default()

                self.project_label.setText(AppConfig.MSG_PRESS_START_ON_PLAYER)
                if self.is_automation_enabled and not self.player_start_attempt_timer.isActive():
                    self.log("Лаунчер найден. Включаю циклические попытки авто-старта...", "info")
                    self.player_start_attempt_timer.start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
                    self.attempt_player_start_click()

        if not self.player_check_timer.isActive():
            self.player_check_timer.start(AppConfig.PLAYER_CHECK_INTERVAL)

    def handle_player_close(self):
        if not self.is_automation_enabled: return
        self.is_sending_logs = True
        self.log("Плеер закрыт. Запускаю отправку логов...", "info")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        launched_count = 0
        shortcuts_to_launch = []
        try:
            for filename in os.listdir(desktop_path):
                if filename.lower().endswith('.lnk'):
                    for keyword in AppConfig.LOG_SENDER_KEYWORDS:
                        if keyword in filename.lower():
                            shortcuts_to_launch.append((keyword, os.path.join(desktop_path, filename)))
                            break
        except Exception as e:
            self.log(f"Ошибка при поиске скриптов логов: {e}", "error")
        if not shortcuts_to_launch:
            self.log("Скрипты для отправки логов не найдены.", "warning")
        else:
            for keyword, path in shortcuts_to_launch:
                if self.window_manager.find_first_window_by_title(keyword):
                    self.log(f"Скрипт '{keyword}' уже запущен. Пропускаю.", "info")
                else:
                    try:
                        os.startfile(path)
                        launched_count += 1
                        self.log(f"Запущен скрипт отправки логов '{keyword}'.", "info")
                    except Exception as e:
                        self.log(f"Не удалось запустить ярлык для '{keyword}': {e}", "error")
        if launched_count > 0:
            self.log(f"Всего запущено скриптов: {launched_count}", "info")
        self.log(f"Перезапуск плеера через {AppConfig.PLAYER_RELAUNCH_DELAY_S} секунд...", "info")
        QTimer.singleShot(AppConfig.PLAYER_RELAUNCH_DELAY_S * 1000, self.wait_for_logs_to_finish)

    def wait_for_logs_to_finish(self):
        still_running = any(self.window_manager.find_first_window_by_title(keyword) for keyword in AppConfig.LOG_SENDER_KEYWORDS)
        if still_running:
            self.log("Ожидание завершения отправки логов...", "info")
            QTimer.singleShot(3000, self.wait_for_logs_to_finish)
        else:
            self.log("Отправка логов завершена.", "info")
            self.is_sending_logs = False
            self.check_for_player()

    def check_and_launch_player(self):
        if self.is_sending_logs: return
        if self.window_manager.find_first_window_by_title(AppConfig.PLAYER_GAME_LAUNCHER_TITLE) or self.window_manager.find_first_window_by_title("launch"):
            self.log("Обнаружен процесс запуска/обновления плеера. Ожидание...", "info")
            return
        self.log("Плеер не найден, ищу ярлык 'launch'...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            for filename in os.listdir(desktop_path):
                if 'launch' in filename.lower() and filename.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, filename)
                    self.log("Найден ярлык плеера. Запускаю...", "info")
                    try:
                        os.startfile(shortcut_path)
                        if not self.player_start_attempt_timer.isActive():
                           self.player_start_attempt_timer.start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
                        return
                    except Exception as e:
                        self.log(f"Не удалось запустить ярлык плеера: {e}", "error")
                        break
            self.log("Ярлык плеера 'launch' на рабочем столе не найден.", "error")
        except Exception as e:
            self.log(f"Ошибка при поиске на рабочем столе: {e}", "error")

    def attempt_player_start_click(self):
        if not self.is_automation_enabled:
            if self.player_start_attempt_timer.isActive():
                self.player_start_attempt_timer.stop()
            return
            
        launcher_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if not launcher_hwnd:
            self.log(f"Не удалось найти окно лаунчера для авто-старта. Отключаю попытки.", "warning")
            if self.player_start_attempt_timer.isActive():
                self.player_start_attempt_timer.stop()
            return
        try:
            if win32gui.IsIconic(launcher_hwnd):
                win32gui.ShowWindow(launcher_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(launcher_hwnd)
            time.sleep(0.5)
            rect = win32gui.GetWindowRect(launcher_hwnd)
            x = rect[0] + 50
            y = rect[1] + 50
            self.log("Попытка авто-старта плеера...", "info")
            self.window_manager.click_at_pos(x, y)
        except Exception as e:
            self.log(f"Не удалось активировать окно плеера: {e}", "error")

    def check_and_launch_opencv_server(self):
        if not self.is_automation_enabled: return
        config = PROJECT_CONFIGS.get("QQ")
        if not config: return
        if self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId()):
            self.log("OpenCV сервер уже запущен.", "info")
            return
        self.log("Сервер OpenCV не найден, ищу ярлык...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            for filename in os.listdir(desktop_path):
                if 'opencv' in filename.lower() and filename.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, filename)
                    self.log("Найден ярлык OpenCV. Запускаю...", "info")
                    try:
                        os.startfile(shortcut_path)
                        QTimer.singleShot(4000, self.arrange_other_windows)
                        return
                    except Exception as e:
                        self.log(f"Не удалось запустить ярлык OpenCV: {e}", "error")
                        break
            self.log("Ярлык для OpenCV на рабочем столе не найден.", "error")
        except Exception as e:
            self.log(f"Ошибка при поиске на рабочем столе: {e}", "error")

    def check_for_recorder(self):
        if self.window_manager.is_process_running("recorder"):
            self.auto_record_toggle_button.setEnabled(True)
            if self.recorder_check_timer.isActive():
                self.recorder_check_timer.stop()
            return
        self.auto_record_toggle_button.setEnabled(False)
        if self.is_automation_enabled:
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            try:
                for filename in os.listdir(desktop_path):
                    if 'recorder' in filename.lower() and filename.lower().endswith('.lnk'):
                        shortcut_path = os.path.join(desktop_path, filename)
                        self.log("Camtasia не найдена. Запускаю...", "warning")
                        try:
                            os.startfile(shortcut_path)
                            if self.recorder_check_timer.isActive(): self.recorder_check_timer.stop()
                            QTimer.singleShot(3000, self.check_for_recorder)
                            return
                        except Exception as e:
                            self.log(f"Не удалось запустить Camtasia: {e}", "error")
                            break
                self.log("Ярлык для Camtasia на рабочем столе не найден.", "error")
            except Exception as e:
                self.log(f"Ошибка при поиске на рабочем столе: {e}", "error")
        if not self.recorder_check_timer.isActive():
            self.recorder_check_timer.start(AppConfig.RECORDER_CHECK_INTERVAL)

    def initial_recorder_sync_check(self):
        if self.window_manager.find_first_window_by_title("Recording..."):
            self.log("Обнаружена активная запись. Перезапускаю для синхронизации...", "warning")
            self.stop_recording_session()
            QTimer.singleShot(2000, self.check_auto_record_logic)

    def toggle_auto_record(self):
        self.is_auto_record_enabled = not self.is_auto_record_enabled
        self.update_auto_record_button_style()
        if self.is_auto_record_enabled:
            self.log("Автозапись включена.", "info")
            self.auto_record_timer.start(AppConfig.AUTO_RECORD_INTERVAL)
            self.check_auto_record_logic()
        else:
            self.log("Автозапись выключена. Ручной режим.", "warning")
            self.auto_record_timer.stop()
            self.progress_bar.setVisible(False)
            self.progress_bar_label.setVisible(False)
            if self.is_flashing: self.stop_flashing()

    def toggle_automation(self):
        self.is_automation_enabled = not self.is_automation_enabled
        self.update_automation_button_style()
        if self.is_automation_enabled:
            self.log("Автоматика включена.", "info")
            self.check_for_player()
            self.check_for_recorder()
        else:
            self.log("Автоматика выключена. Все авто-действия остановлены.", "warning")

    def check_auto_record_logic(self):
        if not self.is_auto_record_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        try:
            should_be_recording = False
            if self.current_project == "GG":
                if self.window_manager.is_process_running("ClubGG.exe"):
                    should_be_recording = True
            else:
                lobby_windows = self.window_manager.find_windows_by_config(config, "LOBBY", self.winId())
                table_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
                if lobby_windows or table_windows:
                    should_be_recording = True
            is_recording = self.window_manager.find_first_window_by_title("Recording...") is not None
            is_paused = self.window_manager.find_first_window_by_title("Paused...") is not None
            is_currently_active = is_recording or is_paused
            if should_be_recording and not is_currently_active:
                self.log("Начинаю автозапись...", "info")
                self.start_recording_session()
            elif should_be_recording and is_paused:
                self.log("Запись на паузе. Возобновляю...", "warning")
                self.perform_camtasia_action(win32con.VK_F9, "возобновление записи")
            elif not should_be_recording and is_currently_active:
                self.log("Активность завершена. Останавливаю запись...", "info")
                self.stop_recording_session()
        except Exception as e:
            self.log(f"Ошибка в логике автозаписи: {e}", "error")

    def start_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F9, "старт записи")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        self.recording_start_time = time.monotonic()
        self.session_timer.start(1000)
        self.progress_bar.setMaximum(config["SESSION_MAX_DURATION_S"])
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.progress_bar_label.setVisible(True)
        self.flashing_disabled_for_session = False

    def stop_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F10, "остановку записи")
        if self.session_timer.isActive():
            self.session_timer.stop()
            self.recording_start_time = 0
            self.progress_bar.setVisible(False)
            self.progress_bar_label.setVisible(False)
            self.stop_flashing()

    # ИЗМЕНЕНО: Новая логика взаимодействия с Camtasia
    def perform_camtasia_action(self, key_code, action_name):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name("recorder")
        if not recorder_hwnd:
            self.log(f"Не удалось выполнить '{action_name}': окно Camtasia не найдено.", "error")
            return
        self.log(f"Выполняю '{action_name}' для Camtasia...", "info")
        try:
            # Сначала позиционируем окно рекордера, чтобы знать его точные координаты
            self.position_recorder_window()
            time.sleep(0.1)

            # Получаем координаты окна
            rect = win32gui.GetWindowRect(recorder_hwnd)
            
            # Клик по кнопке "Full Screen" (координаты 55, 80 от левого верхнего угла)
            x = rect[0] + 55
            y = rect[1] + 80
            self.window_manager.click_at_pos(x, y)
            self.log("Full Screen был нажат.", "info")

            # Ждем 0.5 секунды перед нажатием хоткея
            time.sleep(0.5)
            
            # Нажимаем горячую клавишу (F9 или F10)
            self.window_manager.press_key(key_code)
        except Exception as e:
            self.log(f"Ошибка при взаимодействии с Camtasia: {e}", "error")

    def update_session_progress(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "SESSION_MAX_DURATION_S" not in config: return
        elapsed = time.monotonic() - self.recording_start_time
        self.progress_bar.setValue(int(elapsed))
        remaining_s = config["SESSION_MAX_DURATION_S"] - elapsed
        if remaining_s < 0: remaining_s = 0
        formatted_time = time.strftime('%H:%M:%S', time.gmtime(remaining_s))
        self.progress_bar_label.setText(AppConfig.MSG_PROGRESS_LABEL.format(formatted_time))
        progress_percent = elapsed / config["SESSION_MAX_DURATION_S"]
        if progress_percent < 0.5: color = "#2ECC71"
        elif progress_percent < 0.85: color = "#F1C40F"
        else: color = "#E74C3C"
        self.progress_bar.setStyleSheet(f"QProgressBar::chunk {{ background-color: {color}; border-radius: 5px;}} QProgressBar {{ border: 1px solid #555; border-radius: 5px; text-align: center; height: 12px; background-color: #222;}}")
        if self.current_project == "GG" and config["SESSION_WARN_TIME_S"] > 0 and elapsed >= config["SESSION_WARN_TIME_S"] and not self.is_flashing and not self.flashing_disabled_for_session:
            self.start_flashing()
        if elapsed >= config["SESSION_MAX_DURATION_S"]:
            self.log(f"{config['SESSION_MAX_DURATION_S']/3600:.0f} часа записи истекли. Перезапуск...", "info")
            self.stop_recording_session()
            QTimer.singleShot(2000, self.check_auto_record_logic)

    def check_for_new_tables(self):
        if not self.is_automation_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        
        if self.current_project == "GG":
            if not self.window_manager.is_process_running("ClubGG.exe"):
                if self.last_table_count != 0: self.last_table_count = 0
                return
        
        current_tables = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        current_count = len(current_tables)
        if current_count != self.last_table_count:
            QTimer.singleShot(500, self.arrange_tables)
        self.last_table_count = current_count

    def arrange_tables(self):
        if not self.current_project:
            self.log("Проект не определен.", "error")
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "TABLE" not in config:
            self.log(f"Нет конфига столов для {self.current_project}.", "warning")
            return
            
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        if not found_windows:
            self.log(AppConfig.MSG_ARRANGE_TABLES_NOT_FOUND, "warning")
            return
        
        slots_key = "TABLE_SLOTS_5" if self.current_project == "QQ" and len(found_windows) >= 5 else "TABLE_SLOTS"
        SLOTS = config[slots_key]
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS): break
            x, y = SLOTS[i]
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            if self.current_project == "GG":
                win32gui.MoveWindow(hwnd, x, y, config["TABLE"]["W"], config["TABLE"]["H"], True)
            else:
                rect = win32gui.GetWindowRect(hwnd)
                win32gui.MoveWindow(hwnd, x, y, rect[2] - rect[0], rect[3] - rect[1], True)
            arranged_count += 1
            
        if arranged_count > 0:
            self.log(f"Расставлено столов: {arranged_count}", "info")

    def arrange_other_windows(self):
        if not self.current_project:
            self.log("Проект не определен.", "error")
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            self.log(f"Нет конфига для {self.current_project}.", "warning")
            return
            
        self.position_player_window(config)
        self.position_recorder_window()
        if self.current_project == "GG":
            self.position_lobby_window(config)
            self.position_window_gg_default()
        elif self.current_project == "QQ":
            self.position_cv_server_window(config)
            self.position_window_top_right()
        self.log("Вспомогательные окна расставлены.", "info")

    def position_window(self, hwnd, x, y, w, h, log_success, log_fail):
        if hwnd:
            win32gui.MoveWindow(hwnd, x, y, w, h, True)
        else:
            self.log(log_fail, "warning")

    def position_player_window(self, config):
        player_config = config.get("PLAYER", {})
        if not player_config: return
        player_hwnd = self.window_manager.find_first_window_by_title("holdem")
        self.position_window(player_hwnd, player_config["X"], player_config["Y"], player_config["W"], player_config["H"], "Плеер на месте.", "Плеер не найден.")

    def position_lobby_window(self, config):
        lobbies = self.window_manager.find_windows_by_config(config, "LOBBY", self.winId())
        lobby_hwnd = lobbies[0] if lobbies else None
        cfg = config.get("LOBBY", {})
        if not cfg: return
        self.position_window(lobby_hwnd, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "Лобби на месте.", "Лобби не найдено.")

    def position_cv_server_window(self, config):
        cv_windows = self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winId())
        cv_hwnd = cv_windows[0] if cv_windows else None
        cfg = config.get("CV_SERVER", {})
        if not cfg: return
        self.position_window(cv_hwnd, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "CV Сервер на месте.", "CV Сервер не найден.")

    def position_recorder_window(self):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name("recorder")
        if not recorder_hwnd: return
        try:
            screen_rect = QApplication.primaryScreen().availableGeometry()
            rect = win32gui.GetWindowRect(recorder_hwnd)
            w, h = rect[2] - rect[0], rect[3] - rect[1]
            x = screen_rect.left() + (screen_rect.width() - w) // 2
            y = screen_rect.bottom() - h
            win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
        except Exception as e:
            self.log(f"Ошибка позиционирования Camtasia: {e}", "error")

    def minimize_injector_window(self):
        injector_hwnd = self.window_manager.find_first_window_by_title("injector")
        if injector_hwnd:
            try:
                win32gui.ShowWindow(injector_hwnd, win32con.SW_MINIMIZE)
                self.log("Окно 'injector' свернуто.", "info")
            except Exception as e:
                self.log(f"Не удалось свернуть окно 'injector': {e}", "error")

    def perform_special_clicks(self):
        if self.current_project != "GG": return
        config = PROJECT_CONFIGS.get("GG")
        if not config: return
        
        self.log("Выполняю команду SIT-OUT для столов...", "info")
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winId())
        if not found_windows:
            self.log("Столы для выполнения команды не найдены.", "warning")
            return
            
        click_count = 0
        for hwnd in found_windows:
            try:
                if win32gui.IsWindow(hwnd):
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.1)
                    
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3) 

                    rect = win32gui.GetWindowRect(hwnd)
                    x = rect[0] + 25
                    y = rect[1] + 410
                    self.window_manager.click_at_pos(x, y)
                    time.sleep(0.1) 
                    click_count += 1
            except Exception as e:
                logging.error(f"Не удалось выполнить клик для окна {hwnd}: {e}")
        
        if click_count > 0:
            self.log(f"Команда SIT-OUT выполнена для {click_count} столов.", "info")

    def start_flashing(self):
        if self.is_flashing: return
        height = AppConfig.FLASHING_HEIGHT_H if self.current_project == "GG" else AppConfig.FLASHING_HEIGHT_V
        width = AppConfig.GG_UI_WIDTH if self.current_project == "GG" else AppConfig.DEFAULT_WIDTH
        self.setFixedSize(width, height)
        self.is_flashing = True
        self.stop_flash_button.setVisible(True)
        self.flash_timer.start(AppConfig.FLASH_INTERVAL)

    def stop_flashing(self):
        if not self.is_flashing: return
        self.flashing_disabled_for_session = True
        self.is_flashing = False
        self.rebuild_ui(self.current_project) 
        self.stop_flash_button.setVisible(False)
        self.flash_timer.stop()
        self.setStyleSheet("QMainWindow { background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4c4c4c, stop:1 #2c2c2c); font-family: 'Segoe UI'; }")

    def toggle_window_flash(self):
        if self.flash_state:
            self.setStyleSheet("QMainWindow { border: 3px solid red; background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4c4c4c, stop:1 #2c2c2c); font-family: 'Segoe UI'; }")
        else:
            self.setStyleSheet("QMainWindow { border: none; background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4c4c4c, stop:1 #2c2c2c); font-family: 'Segoe UI'; }")
        self.flash_state = not self.flash_state

    def position_window_gg_default(self):
        """Позиционирует окно в левом нижнем углу, но ниже и левее стандарта."""
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry()
                margin_x = 10 
                margin_y = 40
                x = available_geometry.left() + margin_x
                y = available_geometry.bottom() - self.frameGeometry().height() - margin_y
                self.move(x, y)
        except Exception as e:
            logging.error(f"Could not position window for GG: {e}")
            
    def position_window_top_right(self):
        """Позиционирует окно в правом верхнем углу."""
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry()
                margin = AppConfig.WINDOW_MARGIN
                x = available_geometry.right() - self.frameGeometry().width() - margin
                y = available_geometry.top() + margin
                self.move(x, y)
        except Exception as e:
            logging.error(f"Could not position window top-right: {e}")

    def position_window_default(self):
        """Позиционирует окно по умолчанию (левый нижний угол)."""
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry()
                margin = AppConfig.WINDOW_MARGIN
                x = available_geometry.left() + margin
                y = available_geometry.bottom() - self.frameGeometry().height() - margin
                self.move(x, y)
        except Exception as e:
            logging.error(f"Could not position window default: {e}")


# ===================================================================
# 6. ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ===================================================================

if __name__ == '__main__':
    app = QApplication(sys.argv)

    dark_palette = QPalette()
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
    dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
    dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
    dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
    app.setPalette(dark_palette)
    app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")

    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
