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
from concurrent.futures import ThreadPoolExecutor

# Попытка импортировать опциональные библиотеки и установить флаги доступности
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
    QComboBox, QProgressBar
)
from PyQt6.QtCore import QObject, QTimer, pyqtSignal, Qt, QPropertyAnimation, QRect, QEasingCurve, pyqtProperty
from PyQt6.QtGui import QIcon, QColor, QFont, QPainter, QBrush, QPen
from PyQt6.QtSvg import QSvgRenderer

# ===================================================================
# 1. КОНФИГУРАЦИЯ И СТИЛИ
# ===================================================================

class AppConfig:
    """Централизованная конфигурация приложения."""
    # --- Общие настройки приложения ---
    DEBUG_MODE = False
    CURRENT_VERSION = "7.10"
    MUTEX_NAME = "OiHelperMutex"
    APP_TITLE_TEMPLATE = "OiHelper v{version}"
    ICON_PATH = 'icon.ico'
    LOG_DIR_NAME = 'OiHelper'
    LOG_FILE_NAME = 'app.log'
    THREAD_POOL_WORKERS = 2

    # --- Настройки GitHub и обновлений ---
    GITHUB_REPO = "Vater-v/OiHelper"
    ASSET_NAME = "OiHelper.zip"
    UPDATE_ZIP_NAME = "update.zip"
    UPDATE_TEMP_FOLDER = "update_temp"
    UPDATER_SCRIPT_NAME = "updater.bat"
    UPDATE_CHECK_TIMEOUT_S = 10
    UPDATE_DOWNLOAD_TIMEOUT_S = 60
    UPDATE_CHUNK_SIZE = 8192

    # --- Настройки Telegram Bot ---
    TELEGRAM_BOT_TOKEN = os.environ.get("OIHELPER_TG_TOKEN", '')
    TELEGRAM_CHAT_ID = os.environ.get("OIHELPER_TG_CHAT_ID", '')
    TELEGRAM_REPORT_LEVEL = 'all'
    TELEGRAM_API_TIMEOUT_S = 10
    TELEGRAM_MAX_MSG_LEN = 4096
    TELEGRAM_TRUNCATE_SUFFIX = "\n[...]"

    # --- Заголовки окон и имена процессов ---
    PROJECT_GG = "GG"
    PROJECT_QQ = "QQ"
    PROJECT_WU = "WU"
    PLAYER_LAUNCHER_TITLE = "Holdem"
    PLAYER_GAME_LAUNCHER_TITLE = "Game"
    PLAYER_LAUNCHER_PROCESS_NAME = "launch"
    CAMTASIA_PROCESS_NAME = "recorder"
    INJECTOR_WINDOW_TITLE = "injector"
    OPENCV_SERVER_TITLE = "OpenCv"
    CHROME_PROCESS_NAME = "chrome.exe"
    CLUBGG_PROCESS_NAME = "clubgg.exe"
    CAMTASIA_WINDOW_TITLES = ["Camtasia", "Recording...", "Paused..."]
    CAMTASIA_WINDOW_TITLE_RECORDING = "recording..."
    CAMTASIA_WINDOW_TITLE_PAUSED = "paused..."

    # --- Временные интервалы (в миллисекундах, если не указано иное) ---
    PLAYER_CHECK_INTERVAL = 700
    PLAYER_AUTOSTART_INTERVAL = 500
    AUTO_RECORD_INTERVAL = 600
    AUTO_ARRANGE_INTERVAL = 800
    RECORDER_CHECK_INTERVAL = 800
    POPUP_CHECK_INTERVAL_FAST = 750
    POPUP_CHECK_INTERVAL_SLOW = 10000
    POPUP_FAST_SCAN_DURATION_S = 120
    NOTIFICATION_DURATION = 4500
    STATUS_MESSAGE_DURATION = 3500
    LOG_SENDER_TIMEOUT_S = 300
    PLAYER_RELAUNCH_DELAY_S = 10
    RECORD_RESTART_COOLDOWN_S = 5
    SESSION_PROGRESS_UPDATE_INTERVAL = 1000
    CAMTASIA_ACTION_RETRY_INTERVAL = 350
    CAMTASIA_HOTKEY_WAIT_INTERVAL = 400
    CAMTASIA_LAUNCH_WAIT_S = 2
    CAMTASIA_LAUNCH_POLL_INTERVAL_MS = 500
    CAMTASIA_SYNC_RESTART_DELAY = 400
    LOG_WAIT_RETRY_INTERVAL = 3000
    LAUNCHER_WINDOW_ACTIVATION_TIMEOUT = 5000
    OPENCV_LAUNCH_ARRANGE_DELAY = 1000
    TABLE_ARRANGE_ON_CHANGE_DELAY = 500
    INJECTOR_MINIMIZE_DELAY = 1000
    SESSION_LIMIT_HANDLER_DELAY = 100
    AUTO_STOP_RECORD_INACTIVITY_S = 300  # 5 минут

    # --- Размеры и расположение UI ---
    DEFAULT_WIDTH = 350
    DEFAULT_HEIGHT = 250
    GG_UI_WIDTH = 800
    GG_UI_HEIGHT = 100
    WINDOW_MARGIN = 1
    CLICK_INDICATOR_SIZE = 45
    CLICK_INDICATOR_DURATION = 250
    SPLASH_WIDTH = 300
    SPLASH_HEIGHT = 100
    SPLASH_PROGRESS_HEIGHT = 10
    NOTIFICATION_WIDTH = 420
    NOTIFICATION_MAX_COUNT = 5
    NOTIFICATION_FADE_INTERVAL = 20
    PROGRESS_BAR_HEIGHT = 6
    PROGRESS_BAR_ANIMATION_DURATION = 1000
    PROGRESS_BAR_ALERT_BLINK_INTERVAL = 400
    TOGGLE_SWITCH_WIDTH = 40
    TOGGLE_SWITCH_HEIGHT = 22
    TOGGLE_ANIMATION_DURATION = 200
    RECORDER_FIXED_WIDTH = 410
    RECORDER_FIXED_HEIGHT = 105
    RECORDER_BOTTOM_MARGIN = 30
    QQ_DYNAMIC_ARRANGE_THRESHOLD = 4
    QQ_DYNAMIC_TABLES_PER_ROW = 5
    QQ_DYNAMIC_SCALE_FACTOR = 0.6
    QQ_DYNAMIC_OVERLAP_FACTOR = 0.93

    # --- Сообщения и текст в UI ---
    SPLASH_MSG_LOADING = "Загрузка OiHelper..."
    SPLASH_MSG_UPDATING = "Обновление..."
    SPLASH_MSG_DOWNLOADING = "Скачивание"
    SPLASH_MSG_UNPACKING = "Распаковка"
    SPLASH_MSG_RESTARTING = "Перезапуск..."
    SPLASH_MSG_ERROR = "Ошибка"
    LOBBY_MSG_UPDATE_CHECK = "Проверка обновлений..."
    LOBBY_MSG_SEARCHING = "Поиск проекта..."
    LOBBY_MSG_PLAYER_NOT_FOUND = "Плеер не найден. Запустите его."
    LOBBY_MSG_WAITING_FOR_LAUNCHER = "Ожидание старта в лаунчере..."
    LOBBY_MSG_MANUAL_SELECT_PROMPT = "или выберите вручную:"
    LOBBY_MSG_COMBO_PLACEHOLDER = "Выберите проект..."
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
    MSG_CAMTASIA_AUTOMATION = "Автоматизация Camtasia...\nНе трогайте мышь и клавиатуру"
    MSG_BLOCKING_OVERLAY_DEFAULT = "Пожалуйста, не трогайте мышь и клавиатуру..."
    MSG_AUTO_RECORD_ON = "Автозапись включена."
    MSG_AUTO_RECORD_OFF = "Автозапись выключена."
    MSG_AUTOMATION_ON = "Автоматика включена."
    MSG_AUTOMATION_OFF = "Автоматика выключена."
    MSG_POPUP_CLOSER_ON = "Автозакрытие спама (beta) включено."
    MSG_POPUP_CLOSER_OFF = "Автозакрытие спама (beta) выключено."
    MSG_TG_HELPER_STARTED = f"OiHelper {CURRENT_VERSION} запущен."
    MSG_TG_CRITICAL_ERROR = "OiHelper Критическая ошибка: {}"

    # --- Системные и WinAPI константы ---
    ERROR_ALREADY_EXISTS = 183
    INPUT_MOUSE = 0
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    MOUSEEVENTF_MOVE = 0x0001
    MOUSEEVENTF_ABSOLUTE = 0x8000
    VK_F9 = win32con.VK_F9
    VK_F10 = win32con.VK_F10
    UPTIME_WARN_DAYS = 5

    # --- Параметры автоматизации и кликов ---
    ROBUST_CLICK_DELAY = 0.13
    ROBUST_CLICK_ACTIVATION_DELAY = 0.07
    ROBUST_CLICK_SET_CURSOR_DELAY = 0.07
    ROBUST_CLICK_MOUSE_DOWN_DELAY = 0.02
    DOUBLE_ROBUST_CLICK_INTERVAL = 0.24
    SENDINPUT_MOVE_DELAY = 0.03
    SENDINPUT_CLICK_INTERVAL_MIN = 0.05
    SENDINPUT_CLICK_INTERVAL_MAX = 0.1
    KEY_PRESS_DELAY = 0.03
    KEY_PRESS_WAIT_AFTER = 0.15
    CAMTASIA_MAX_ACTION_ATTEMPTS = 6
    CAMTASIA_RESUME_CHECK_ATTEMPTS = 6
    CAMTASIA_RESUME_CHECK_INTERVAL = 400
    DEFAULT_CV_CONFIDENCE = 0.73
    POPUP_ACTION_DELAY = 0.5
    BONUS_SPIN_WAIT_S = 4

    # --- Пути к файлам и ключевые слова ---
    LOG_SENDER_KEYWORDS = ["endsess", "logbot"]
    PLAYER_LAUNCHER_SHORTCUT_KEYWORD = 'launch'
    CAMTASIA_SHORTCUT_KEYWORD = 'recorder'
    OPENCV_SHORTCUT_KEYWORD = 'opencv'
    SHORTCUT_EXTENSION = '.lnk'
    TEMPLATES_DIR = 'templates'
    CAMTASIA_FULLSCREEN_TEMPLATE = 'camtasia_fullscreen.png'
    CAMTASIA_REC_TEMPLATE = 'camtasia_rec.png'
    CAMTASIA_STOP_TEMPLATE = 'camtasia_stop.png'
    CAMTASIA_RESUME_TEMPLATE = 'camtasia_resume.png'

    # --- Ключи для конфигурации проектов ---
    KEY_TABLE = "TABLE"
    KEY_LOBBY = "LOBBY"
    KEY_PLAYER = "PLAYER"
    KEY_CV_SERVER = "CV_SERVER"
    KEY_FIND_METHOD = "FIND_METHOD"
    KEY_TITLE = "TITLE"
    KEY_WIDTH = "W"
    KEY_HEIGHT = "H"
    KEY_TOLERANCE = "TOLERANCE"
    KEY_X = "X"
    KEY_Y = "Y"
    KEY_EXCLUDED_TITLES = "EXCLUDED_TITLES"
    KEY_EXCLUDED_PROCESSES = "EXCLUDED_PROCESSES"
    KEY_ARRANGE_MINIMIZED = "ARRANGE_MINIMIZED_TABLES"
    KEY_TABLE_SLOTS = "TABLE_SLOTS"
    KEY_TABLE_SLOTS_5 = "TABLE_SLOTS_5"
    KEY_SESSION_MAX_S = "SESSION_MAX_DURATION_S"
    KEY_SESSION_WARN_S = "SESSION_WARN_TIME_S"
    KEY_POPUPS = "POPUPS"
    KEY_CONFIDENCE = "CONFIDENCE"
    KEY_BUY_IN = "BUY_IN"
    KEY_SPAM = "SPAM"
    KEY_BONUS = "BONUS"
    KEY_POPUP_TRIGGER = "trigger"
    KEY_POPUP_MAX_BTN = "max_button"
    KEY_POPUP_CONFIRM_BTN = "confirm_button"

    FIND_METHOD_RATIO = "RATIO"
    FIND_METHOD_TITLE_AND_SIZE = "TITLE_AND_SIZE"
    FIND_METHOD_TITLE = "TITLE"
    FIND_METHOD_PROCESS_AND_TITLE = "PROCESS_AND_TITLE"


    # --- Действия Camtasia ---
    ACTION_START = "start"
    ACTION_STOP = "stop"
    ACTION_RESUME = "resume"

    # --- Проекты ---
    PROJECT_MAPPING = {"QQPoker": PROJECT_QQ, "ClubGG": PROJECT_GG, "WUPoker": PROJECT_WU}

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
    WHITE = "#FFFFFF"
    BLACK = "#000000"
    OVERLAY_BG = "rgba(0,0,0,120)"
    OVERLAY_CLOSE_BTN_BG = "rgba(255,255,255,0.92)"
    OVERLAY_CLOSE_BTN_HOVER_BG = "#FFEBEB"
    OVERLAY_CLOSE_BTN_HOVER_COLOR = "#B91C1C"
    OVERLAY_SHADOW = "rgba(0,0,0,0.14)"
    NOTIFICATION_SHADOW = "rgba(0,0,0,90)"
    BUTTON_SHADOW = "rgba(0,0,0,40)"

class StyleSheet:
    FONT_FAMILY = "'Segoe UI', 'Roboto'"
    FONT_FAMILY_SEMIBOLD = "'Segoe UI Semibold', 'Roboto'"

    MAIN_WINDOW = f"background-color: {ColorPalette.BACKGROUND};"
    SPLASH_LABEL = f"font-family: {FONT_FAMILY}; font-size: 14px; color: {ColorPalette.TEXT_SECONDARY};"
    LOBBY_LABEL = f"font-family: {FONT_FAMILY}; font-size: 14px; color: {ColorPalette.TEXT_SECONDARY};"
    STATUS_LABEL = f"font-family: {FONT_FAMILY_SEMIBOLD}; font-size: 13px; color: {ColorPalette.TEXT_PRIMARY};"
    PROGRESS_BAR_LABEL = f"font-family: {FONT_FAMILY}; font-size: 11px; color: {ColorPalette.TEXT_SECONDARY};"
    COMBO_BOX = f"""
        QComboBox {{
            font-family: {FONT_FAMILY};
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
            selection-color: {ColorPalette.WHITE};
        }}
    """

    @staticmethod
    def get_button_style(primary=True):
        base_style = f"QPushButton {{ font-family: {StyleSheet.FONT_FAMILY_SEMIBOLD}; font-size: 12px; border-radius: 6px; padding: 8px 12px; }}"
        if primary:
            return base_style + f"QPushButton {{ background-color: {ColorPalette.PRIMARY}; color: {ColorPalette.WHITE}; border: none; }} QPushButton:hover {{ background-color: {ColorPalette.PRIMARY_HOVER}; }} QPushButton:pressed {{ background-color: {ColorPalette.PRIMARY_PRESSED}; }} QPushButton:disabled {{ background-color: {ColorPalette.SECONDARY}; color: {ColorPalette.TEXT_SECONDARY}; }}"
        else:
            return base_style + f"QPushButton {{ background-color: {ColorPalette.SURFACE}; color: {ColorPalette.TEXT_PRIMARY}; border: 1px solid {ColorPalette.BORDER}; }} QPushButton:hover {{ background-color: {ColorPalette.BACKGROUND}; }} QPushButton:pressed {{ background-color: {ColorPalette.SECONDARY}; }} QPushButton:disabled {{ background-color: {ColorPalette.BACKGROUND}; color: {ColorPalette.TEXT_SECONDARY}; border-color: {ColorPalette.SECONDARY}; }}"

PROJECT_CONFIGS = {
    AppConfig.PROJECT_GG: {
        AppConfig.KEY_TABLE: {
            AppConfig.KEY_FIND_METHOD: AppConfig.FIND_METHOD_RATIO,
            AppConfig.KEY_WIDTH: 557,
            AppConfig.KEY_HEIGHT: 424,
            AppConfig.KEY_TOLERANCE: 0.035
        },
        AppConfig.KEY_LOBBY: {
            AppConfig.KEY_FIND_METHOD: AppConfig.FIND_METHOD_RATIO,
            AppConfig.KEY_WIDTH: 333,
            AppConfig.KEY_HEIGHT: 623,
            AppConfig.KEY_TOLERANCE: 0.07,
            AppConfig.KEY_X: 1580,
            AppConfig.KEY_Y: 140
        },
        AppConfig.KEY_PLAYER: {
            AppConfig.KEY_WIDTH: 700,
            AppConfig.KEY_HEIGHT: 365,
            AppConfig.KEY_X: 1385,
            AppConfig.KEY_Y: 0
        },
        AppConfig.KEY_TABLE_SLOTS: [(-5, 0), (271, 423), (816, 0), (1086, 423)],
        AppConfig.KEY_EXCLUDED_TITLES: ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"],
        AppConfig.KEY_EXCLUDED_PROCESSES: ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe"],
        AppConfig.KEY_SESSION_MAX_S: 4 * 3600,
        AppConfig.KEY_SESSION_WARN_S: 3.5 * 3600,
        AppConfig.KEY_ARRANGE_MINIMIZED: False,
        AppConfig.KEY_POPUPS: {
            AppConfig.KEY_CONFIDENCE: 0.85,
            AppConfig.KEY_BUY_IN: {
                AppConfig.KEY_POPUP_TRIGGER: "buyin_window.png",
                AppConfig.KEY_POPUP_MAX_BTN: "max_button.png",
                AppConfig.KEY_POPUP_CONFIRM_BTN: "buyin_button.png"
            }
        }
    },
    AppConfig.PROJECT_QQ: {
        AppConfig.KEY_TABLE: {AppConfig.KEY_FIND_METHOD: AppConfig.FIND_METHOD_TITLE_AND_SIZE, AppConfig.KEY_TITLE: "QQPK", AppConfig.KEY_WIDTH: 400, AppConfig.KEY_HEIGHT: 700, AppConfig.KEY_TOLERANCE: 2},
        AppConfig.KEY_LOBBY: {AppConfig.KEY_FIND_METHOD: AppConfig.FIND_METHOD_RATIO, AppConfig.KEY_WIDTH: 400, AppConfig.KEY_HEIGHT: 700, AppConfig.KEY_TOLERANCE: 0.07, AppConfig.KEY_X: 1418, AppConfig.KEY_Y: 0},
        AppConfig.KEY_CV_SERVER: {AppConfig.KEY_FIND_METHOD: AppConfig.FIND_METHOD_TITLE, AppConfig.KEY_TITLE: AppConfig.OPENCV_SERVER_TITLE, AppConfig.KEY_X: 1789, AppConfig.KEY_Y: 367, AppConfig.KEY_WIDTH: 993, AppConfig.KEY_HEIGHT: 605},
        AppConfig.KEY_PLAYER: {AppConfig.KEY_X: 1418, AppConfig.KEY_Y: 942, AppConfig.KEY_WIDTH: 724, AppConfig.KEY_HEIGHT: 370},
        AppConfig.KEY_TABLE_SLOTS: [(0, 0), (401, 0), (802, 0), (1203, 0)],
        AppConfig.KEY_TABLE_SLOTS_5: [(0, 0), (346, 0), (692, 0), (1038, 0), (1384, 0)],
        AppConfig.KEY_EXCLUDED_TITLES: ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandboxie", "Notepad", "Explorer"],
        AppConfig.KEY_EXCLUDED_PROCESSES: ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "OpenCvServer.exe"],
        AppConfig.KEY_SESSION_MAX_S: 3 * 3600,
        AppConfig.KEY_SESSION_WARN_S: -1,
        AppConfig.KEY_ARRANGE_MINIMIZED: True,
        AppConfig.KEY_POPUPS: {
            AppConfig.KEY_CONFIDENCE: 0.83,
            AppConfig.KEY_SPAM: [
                {"poster": "popup_poster_1.png", "close_button": "popup_close_btn_1.png"},
                {"poster": "popup_poster_2.png", "close_button": "popup_close_btn_2.png"},
                {"poster": "popup_poster_3.png", "close_button": "popup_close_button.png"},
                {"poster": "popup_poster_4.png", "close_button": "popup_close_button.png"}
            ],
            AppConfig.KEY_BONUS: {
                "wheel": "bonus_wheel.png",
                "spin_button": "bonus_spin_button.png",
                "close_button": "bonus_close_button.png"
            }
        }
    },
    AppConfig.PROJECT_WU: {
        AppConfig.KEY_PLAYER: {
            AppConfig.KEY_X: 1400,
            AppConfig.KEY_Y: 900,
            AppConfig.KEY_WIDTH: 724,
            AppConfig.KEY_HEIGHT: 370
        },
            AppConfig.KEY_EXCLUDED_TITLES: ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"],
            AppConfig.KEY_EXCLUDED_PROCESSES: ["explorer.exe", "svchost.exe"],
            AppConfig.KEY_SESSION_MAX_S: 3 * 3600,
            AppConfig.KEY_SESSION_WARN_S: -1,
            AppConfig.KEY_ARRANGE_MINIMIZED: True,
            AppConfig.KEY_POPUPS: {
                # Если нужно - добавь обработку попапов
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
        self.setFixedSize(AppConfig.SPLASH_WIDTH, AppConfig.SPLASH_HEIGHT)

        container = QFrame(self)
        container.setStyleSheet(f"background-color: {ColorPalette.SURFACE}; border-radius: 8px;")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(ColorPalette.BLACK))
        shadow.setOffset(0, 4)
        container.setGraphicsEffect(shadow)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(container)

        layout = QVBoxLayout(container)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.status_label = QLabel(AppConfig.SPLASH_MSG_LOADING)
        self.status_label.setStyleSheet(StyleSheet.SPLASH_LABEL)
        layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(AppConfig.SPLASH_PROGRESS_HEIGHT)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)
        self.progress_bar.hide()

    def update_status(self, message: str):
        self.status_label.setText(message)

    def set_progress(self, value: int, stage: str = None):
        self.progress_bar.setValue(value)
        if not self.progress_bar.isVisible():
            self.progress_bar.show()
        if stage:
            self.status_label.setText(f"{stage}... ({value}%)")
            if stage == AppConfig.SPLASH_MSG_ERROR:
                self.hide_progress()

    def hide_progress(self):
        self.progress_bar.hide()

CAMTASIA_SHORTCUT_KEYWORD = "camtasia"
SHORTCUT_EXTENSION = ".lnk"


def try_launch_camtasia_shortcut():
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    try:
        shortcut_path = next(
            (os.path.join(desktop_path, f) for f in os.listdir(desktop_path)
             if CAMTASIA_SHORTCUT_KEYWORD in f.lower() and f.lower().endswith(SHORTCUT_EXTENSION)),
            None)
        if shortcut_path:
            os.startfile(shortcut_path)
            logging.info("Запущен ярлык Camtasia: %s", shortcut_path)
            return True
        else:
            logging.warning("Ярлык Camtasia на рабочем столе не найден.")
    except Exception as e:
        logging.error(f"Ошибка при запуске ярлыка Camtasia: {e}")
    return False

def find_camtasia_window():
    result = []
    def callback(hwnd, _):
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
            win32api.CloseHandle(h_process)
            title = win32gui.GetWindowText(hwnd)
            # Критерий можно усилить/ослабить под свою сборку
            if "camtasia" in process_name.lower() and (
                "recorder" in title.lower() or "paused" in title.lower() or "recording" in title.lower()):
                result.append(hwnd)
        except Exception:
            pass
    win32gui.EnumWindows(callback, None)
    return result[0] if result else None

def focus_camtasia_window(max_retries=3, wait_launch=2.5, poll_interval=0.5):
    """
    Попытка сфокусировать Camtasia:
      - ищет окно
      - если не найдено — запускает ярлык, ждёт появления
      - разворачивает, активирует, логирует
      - повторяет max_retries раз
    """
    for attempt in range(max_retries):
        hwnd = find_camtasia_window()
        if hwnd and win32gui.IsWindow(hwnd):
            try:
                title = win32gui.GetWindowText(hwnd)
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    time.sleep(0.13)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.09)
                logging.info(f"[focus_camtasia_window] hwnd={hwnd}, title='{title}' успешно активировано")
                return True
            except Exception as e:
                logging.error(f"[focus_camtasia_window] Ошибка: {e}")
                return False
        if attempt == 0:
            logging.warning("Окно Camtasia не найдено. Пробую запустить ярлык.")
            try_launch_camtasia_shortcut()
            # даём время процессу появиться
            time.sleep(wait_launch)
        else:
            # Если уже пробовали — просто ждём чуть-чуть и ещё раз ищем
            time.sleep(poll_interval)
    logging.error("Не удалось найти или активировать окно Camtasia после повторных попыток.")
    return False



class BlockingOverlay(QWidget):
    def __init__(self, message=AppConfig.MSG_BLOCKING_OVERLAY_DEFAULT, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)
        self.setStyleSheet(f"background-color: {ColorPalette.OVERLAY_BG};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        close_btn = QPushButton("✕", self)
        close_btn.setFixedSize(56, 56)
        close_btn.setStyleSheet(
            f"QPushButton {{"
            f"  background: {ColorPalette.OVERLAY_CLOSE_BTN_BG};"
            f"  color: {ColorPalette.RED};"
            f"  font-size: 38px;"
            f"  font-weight: bold;"
            f"  border-radius: 28px;"
            f"  border: 2px solid {ColorPalette.RED};"
            f"  box-shadow: 0 6px 20px {ColorPalette.OVERLAY_SHADOW};"
            f"}}"
            f"QPushButton:hover {{"
            f"  background: {ColorPalette.OVERLAY_CLOSE_BTN_HOVER_BG};"
            f"  color: {ColorPalette.OVERLAY_CLOSE_BTN_HOVER_COLOR};"
            f"  border-color: {ColorPalette.OVERLAY_CLOSE_BTN_HOVER_COLOR};"
            f"}}"
        )
        close_btn.clicked.connect(self.close)
        close_btn.move(self.width() - close_btn.width() - 40, (self.height() - close_btn.height()) // 2)
        close_btn.raise_()

        msg_layout = QVBoxLayout()
        label = QLabel(message)
        label.setStyleSheet(f"color: {ColorPalette.WHITE}; font-size: 26px; font-weight: bold;")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        msg_layout.addStretch()
        msg_layout.addWidget(label)
        msg_layout.addStretch()
        layout.addLayout(msg_layout)

    def resizeEvent(self, event):
        btn = self.findChild(QPushButton)
        if btn:
            btn.move(self.width() - btn.width() - 40, (self.height() - btn.height()) // 2)
        super().resizeEvent(event)

class ToggleSwitch(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFixedSize(AppConfig.TOGGLE_SWITCH_WIDTH, AppConfig.TOGGLE_SWITCH_HEIGHT)
        self.clicked.connect(self.update_state)
        self._circle_pos = 3
        self.animation = QPropertyAnimation(self, b"circle_pos", self)
        self.animation.setDuration(AppConfig.TOGGLE_ANIMATION_DURATION)
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

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(ColorPalette.NOTIFICATION_SHADOW))
        shadow.setOffset(0, 4)
        container.setGraphicsEffect(shadow)

        layout = QHBoxLayout(container)
        layout.setContentsMargins(20, 15, 20, 15)
        text_label = QLabel(message)
        text_label.setFont(QFont(StyleSheet.FONT_FAMILY, 14))
        text_label.setWordWrap(True)
        text_label.setStyleSheet(f"background: transparent; border: none; color: {ColorPalette.TEXT_PRIMARY};")
        layout.addWidget(text_label, 1)
        self.setFixedWidth(AppConfig.NOTIFICATION_WIDTH)

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

    def show_animation(self):
        self.show()
        self.fade_in_timer.start(AppConfig.NOTIFICATION_FADE_INTERVAL)

    def start_fade_out(self):
        if self.is_closing: return
        self.is_closing = True
        self.fade_in_timer.stop()
        self.fade_out_timer.start(AppConfig.NOTIFICATION_FADE_INTERVAL)

    def closeEvent(self, event):
        self.closed.emit(self)
        super().closeEvent(event)

class NotificationManager(QObject):
    def __init__(self):
        super().__init__()
        self.notifications = []

    def show(self, message: str, message_type: str):
        logging.info(f"Уведомление [{message_type}]: {message}")
        if any(n.message == message for n in self.notifications): return
        if len(self.notifications) >= AppConfig.NOTIFICATION_MAX_COUNT:
            self.notifications.pop(0).start_fade_out()

        notification = Notification(message, message_type)
        notification.closed.connect(self.on_notification_closed)
        self.notifications.append(notification)
        self.reposition_all()
        notification.show_animation()

    def on_notification_closed(self, notification: QWidget):
        if notification in self.notifications: self.notifications.remove(notification)
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

    def _toggle_alert_visibility(self):
        self._alert_visible = not self._alert_visible
        self.update()

    def get_progress_value(self): return self._value
    def set_progress_value(self, value): self._value = value; self.update()
    progress_value = pyqtProperty(float, fget=get_progress_value, fset=set_progress_value)

    def setValue(self, value):
        value = max(0, min(float(value), self._maximum))
        if self._animation.state() == QPropertyAnimation.State.Running: self._animation.stop()
        self._animation.setStartValue(self.progress_value)
        self._animation.setEndValue(value)
        self._animation.setDuration(AppConfig.PROGRESS_BAR_ANIMATION_DURATION)
        self._animation.start()

        is_limit_reached = value >= self._maximum
        if is_limit_reached and not self._alert_mode:
            self._alert_mode = True
            self._alert_timer.start(AppConfig.PROGRESS_BAR_ALERT_BLINK_INTERVAL)
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
        bg_rect = self.rect()
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(ColorPalette.SECONDARY))
        painter.drawRoundedRect(bg_rect, self.height() // 2, self.height() // 2)

        percent = (self.progress_value / self._maximum) if self._maximum > 0 else 0
        progress_width = percent * bg_rect.width()

        if self._alert_mode:
            color = QColor(ColorPalette.RED) if self._alert_visible else QColor(ColorPalette.SECONDARY)
        elif percent > 0.85:
            color = QColor(ColorPalette.RED)
        elif percent > 0.50:
            color = QColor(ColorPalette.AMBER)
        else:
            color = QColor(ColorPalette.PRIMARY)

        if progress_width > 0:
            progress_rect = QRect(0, 0, int(progress_width), self.height())
            painter.setBrush(color)
            painter.drawRoundedRect(progress_rect, self.height() // 2, self.height() // 2)

class ClickIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(AppConfig.CLICK_INDICATOR_SIZE, AppConfig.CLICK_INDICATOR_SIZE)
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
        self.timer.start(AppConfig.CLICK_INDICATOR_DURATION)

class WindowManager(QObject):
    log_request = pyqtSignal(str, str)
    click_visual_request = pyqtSignal(int, int)

    def __init__(self):
        super().__init__()
        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
                        ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))]
        class INPUT(ctypes.Structure):
            _fields_ = [("type", wintypes.DWORD), ("mi", MOUSEINPUT)]
        self.INPUT_STRUCT = INPUT
        self.screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    def _send_mouse_input(self, flags, x=0, y=0):
        if flags & AppConfig.MOUSEEVENTF_ABSOLUTE:
            x = (x * 65535) // self.screen_width
            y = (y * 65535) // self.screen_height
        mi = self.INPUT_STRUCT._fields_[1][1](dx=x, dy=y, mouseData=0, dwFlags=flags, time=0, dwExtraInfo=None)
        mouse_input = self.INPUT_STRUCT(type=AppConfig.INPUT_MOUSE, mi=mi)
        ctypes.windll.user32.SendInput(1, ctypes.byref(mouse_input), ctypes.sizeof(mouse_input))

    def double_robust_click(self, x: int, y: int, hwnd: int = None, delay: float = AppConfig.ROBUST_CLICK_DELAY, interval: float = AppConfig.DOUBLE_ROBUST_CLICK_INTERVAL, log_prefix: str = "") -> bool:
        first = self.robust_click(x, y, hwnd=hwnd, delay=delay, log_prefix=log_prefix)
        time.sleep(interval)
        second = self.robust_click(x, y, hwnd=hwnd, delay=delay, log_prefix=log_prefix)
        return first and second

    def click_camtasia_fullscreen(self, hwnd):
        fs_path = os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_FULLSCREEN_TEMPLATE)
        fs_coords = self.find_template(fs_path)
        if fs_coords:
            self.log_request.emit("Нашёл кнопку Fullscreen. Жму...", "info")
            ok = self.double_robust_click(fs_coords[0], fs_coords[1], hwnd=hwnd)
            logging.info(f"[Camtasia] Double click on fullscreen at {fs_coords}, ok={ok}")
            time.sleep(AppConfig.ROBUST_CLICK_DELAY)
            return ok
        else:
            self.log_request.emit("Не нашёл кнопку Fullscreen!", "warning")
            logging.warning("[Camtasia] Не найден шаблон fullscreen")
            return False

    def find_template(self, template_path: str, confidence=AppConfig.DEFAULT_CV_CONFIDENCE) -> Optional[Tuple[int, int]]:
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

    def find_and_click_template(self, template_path: str, confidence=AppConfig.DEFAULT_CV_CONFIDENCE, hwnd: int = None, log_prefix: str = "") -> bool:
        coords = self.find_template(template_path, confidence)
        if coords:
            self.humanized_click(coords[0], coords[1], hwnd=hwnd, log_prefix=log_prefix)
            return True
        return False

    def find_windows_by_config(self, config: dict, config_key: str, main_window_hwnd: int) -> List[int]:
        window_config = config.get(config_key, {})
        if not window_config: return []
        find_method = window_config.get(AppConfig.KEY_FIND_METHOD)
        EXCLUDED_TITLES = config.get(AppConfig.KEY_EXCLUDED_TITLES, [])
        EXCLUDED_PROCESSES = config.get(AppConfig.KEY_EXCLUDED_PROCESSES, [])
        arrange_minimized = config.get(AppConfig.KEY_ARRANGE_MINIMIZED, False) and config_key == AppConfig.KEY_TABLE
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
                if find_method == AppConfig.FIND_METHOD_RATIO:
                    if h != 0 and (window_config[AppConfig.KEY_WIDTH] / window_config[AppConfig.KEY_HEIGHT]) * (1 - window_config[AppConfig.KEY_TOLERANCE]) <= (w / h) <= (window_config[AppConfig.KEY_WIDTH] / window_config[AppConfig.KEY_HEIGHT]) * (1 + window_config[AppConfig.KEY_TOLERANCE]): match = True
                elif find_method == AppConfig.FIND_METHOD_TITLE_AND_SIZE:
                    if window_config[AppConfig.KEY_TITLE].lower() in title.lower() and abs(w - window_config[AppConfig.KEY_WIDTH]) <= window_config[AppConfig.KEY_TOLERANCE] and abs(h - window_config[AppConfig.KEY_HEIGHT]) <= window_config[AppConfig.KEY_TOLERANCE]: match = True
                elif find_method == AppConfig.FIND_METHOD_TITLE:
                    if window_config[AppConfig.KEY_TITLE].lower() in title.lower(): match = True
                if match: found_windows.append(hwnd)
            except Exception as e:
                logging.debug(f"Ошибка при перечислении окна {hwnd}: {e}")
            return True

        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            logging.error("Критическая ошибка в EnumWindows", exc_info=True)
        try:
            found_windows.sort(key=lambda hwnd: win32gui.GetWindowRect(hwnd)[0])
        except Exception as e:
            self.log_request.emit(f"Ошибка сортировки окон: {e}", "warning")
        return found_windows

    def find_first_window_by_title(self, text_in_title: str, exact_match: bool = False) -> Optional[int]:
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                except Exception:
                    return
                if (exact_match and text_in_title == title) or (not exact_match and text_in_title.lower() in title.lower()):
                    hwnds.append(hwnd)
        try:
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            logging.error("Критическая ошибка в EnumWindows (поиск по заголовку)", exc_info=True)
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
                    if process_name_to_find.lower() in process_name.lower():
                        return True
            except Exception:
                continue
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
                    if process_name_to_find.lower() in process_name.lower():
                        hwnds.append(hwnd)
                except Exception:
                    pass
        try:
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            logging.error("Критическая ошибка в EnumWindows (поиск по процессу)", exc_info=True)
        return hwnds[0] if hwnds else None

    def press_key(self, key_code: int):
        try:
            win32api.keybd_event(key_code, 0, 0, 0)
            time.sleep(AppConfig.KEY_PRESS_DELAY)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e:
            logging.error("Ошибка эмуляции нажатия клавиши", exc_info=True)
            self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

    def close_window(self, hwnd: int):
        try:
            if win32gui.IsWindow(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                logging.info(f"Отправлено сообщение WM_CLOSE окну {hwnd}")
                return True
        except Exception as e:
            logging.error(f"Не удалось отправить WM_CLOSE окну {hwnd}", exc_info=True)
        return False

    def robust_click(self, x: int, y: int, hwnd: int = None, delay: float = AppConfig.ROBUST_CLICK_DELAY, log_prefix: str = "") -> bool:
        try:
            if hwnd:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(delay)
                fg = win32gui.GetForegroundWindow()
                if fg != hwnd:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(AppConfig.ROBUST_CLICK_ACTIVATION_DELAY)
                    fg = win32gui.GetForegroundWindow()
                if fg != hwnd:
                    logging.error(f"{log_prefix} robust_click: Не удалось активировать окно для клика (hwnd={hwnd}, fg={fg})")
                    return False
            old_pos = win32api.GetCursorPos()
            ctypes.windll.user32.SetCursorPos(x, y)
            time.sleep(AppConfig.ROBUST_CLICK_SET_CURSOR_DELAY)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            time.sleep(AppConfig.ROBUST_CLICK_MOUSE_DOWN_DELAY)
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
        if len(message) > AppConfig.TELEGRAM_MAX_MSG_LEN:
            truncate_len = AppConfig.TELEGRAM_MAX_MSG_LEN - len(AppConfig.TELEGRAM_TRUNCATE_SUFFIX)
            message = message[:truncate_len] + AppConfig.TELEGRAM_TRUNCATE_SUFFIX
        self.message_queue.put(message)

    def _worker(self):
        while True:
            message = self.message_queue.get()
            try:
                requests.post(self.api_url, data={'chat_id': self.chat_id, 'text': message}, timeout=AppConfig.TELEGRAM_API_TIMEOUT_S).raise_for_status()
                logging.info("Сообщение в Telegram успешно отправлено.")
            except requests.RequestException as e:
                logging.error(f"Не удалось отправить сообщение в Telegram: {e}")
            finally:
                self.message_queue.task_done()

def find_camtasia_window() -> Optional[int]:
    result = []
    def callback(hwnd, _):
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd): return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
            win32api.CloseHandle(h_process)
            if AppConfig.CAMTASIA_PROCESS_NAME not in process_name.lower(): return
            title = win32gui.GetWindowText(hwnd)
            if any(t.lower() in title.lower() for t in AppConfig.CAMTASIA_WINDOW_TITLES):
                result.append(hwnd)
        except Exception:
            pass
    win32gui.EnumWindows(callback, result)
    return result[0] if result else None

class UpdateManager(QObject):
    log_request = pyqtSignal(str, str)
    check_finished = pyqtSignal()
    status_update = pyqtSignal(str)
    progress_changed = pyqtSignal(int, str)

    def __init__(self):
        super().__init__()
        self.update_info = {}

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
            response = requests.get(api_url, timeout=AppConfig.UPDATE_CHECK_TIMEOUT_S)
            response.raise_for_status()
            latest_release = response.json()
            if (latest_version := latest_release.get("tag_name")) and self.is_new_version_available(AppConfig.CURRENT_VERSION, latest_version):
                self.status_update.emit(f"Найдена версия {latest_version}...")
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
            logging.error("Неожиданная ошибка при проверке обновлений", exc_info=True)
            self.check_finished.emit()

    def apply_update(self):
        download_url = next((asset["browser_download_url"] for asset in self.update_info.get("assets", []) if asset["name"] == AppConfig.ASSET_NAME), None)
        if not download_url:
            self.log_request.emit("Не удалось найти ZIP-архив в релизе.", "error")
            return
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url: str):
        try:
            self.status_update.emit(AppConfig.SPLASH_MSG_DOWNLOADING)
            response = requests.get(url, stream=True, timeout=AppConfig.UPDATE_DOWNLOAD_TIMEOUT_S)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            with open(AppConfig.UPDATE_ZIP_NAME, "wb") as f:
                for chunk in response.iter_content(chunk_size=AppConfig.UPDATE_CHUNK_SIZE):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        percent = int((downloaded / total_size) * 100) if total_size else 0
                        self.progress_changed.emit(percent, AppConfig.SPLASH_MSG_DOWNLOADING)

            self.status_update.emit(AppConfig.SPLASH_MSG_UNPACKING)
            self.progress_changed.emit(0, AppConfig.SPLASH_MSG_UNPACKING)
            if os.path.isdir(AppConfig.UPDATE_TEMP_FOLDER):
                import shutil
                shutil.rmtree(AppConfig.UPDATE_TEMP_FOLDER)
            with zipfile.ZipFile(AppConfig.UPDATE_ZIP_NAME, 'r') as zip_ref:
                filelist = zip_ref.infolist()
                total_files = len(filelist)
                for i, member in enumerate(filelist, 1):
                    zip_ref.extract(member, AppConfig.UPDATE_TEMP_FOLDER)
                    percent = int(i / total_files * 100)
                    self.progress_changed.emit(percent, AppConfig.SPLASH_MSG_UNPACKING)

            self.status_update.emit(AppConfig.SPLASH_MSG_RESTARTING)
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            exe_name = os.path.basename(current_exe_path)
            script_content = (
                f'@echo off\n'
                f'chcp 65001 > NUL\n'
                f'echo Waiting for OiHelper to close...\n'
                f'timeout /t 2 /nobreak > NUL\n'
                f'taskkill /pid {os.getpid()} /f > NUL\n'
                f'echo Waiting for process to terminate...\n'
                f'timeout /t 3 /nobreak > NUL\n'
                f'echo Moving new files...\n'
                f'robocopy "{current_dir}\\{AppConfig.UPDATE_TEMP_FOLDER}" "{current_dir}" /e /move /is > NUL\n'
                f'rd /s /q "{current_dir}\\{AppConfig.UPDATE_TEMP_FOLDER}"\n'
                f'echo Cleaning up...\n'
                f'del "{current_dir}\\{AppConfig.UPDATE_ZIP_NAME}"\n'
                f'echo Starting new version...\n'
                f'start "" "{exe_name}"\n'
                f'(goto) 2>nul & del "%~f0"\n'
            )
            with open(AppConfig.UPDATER_SCRIPT_NAME, "w", encoding="cp866") as f:
                f.write(script_content)
            subprocess.Popen([AppConfig.UPDATER_SCRIPT_NAME], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()

        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error")
            logging.error("Ошибка при обновлении", exc_info=True)
            if os.path.exists(AppConfig.UPDATE_ZIP_NAME):
                try:
                    os.remove(AppConfig.UPDATE_ZIP_NAME)
                except OSError as err:
                    logging.error(f"Не удалось удалить временный файл обновления: {err}")
            self.log_request.emit(AppConfig.MSG_UPDATE_FAIL, "warning")
            self.check_finished.emit()
            self.progress_changed.emit(0, AppConfig.SPLASH_MSG_ERROR)

class CamtasiaWorker(QObject):
    finished = pyqtSignal(bool)

    def __init__(self, wm, action):
        super().__init__()
        self.wm = wm
        self.action = action

    def run(self):
        hwnd = find_camtasia_window()
        if not hwnd:
            try_launch_camtasia_shortcut()
            wait_cycles = int(AppConfig.CAMTASIA_LAUNCH_WAIT_S * 1000 / AppConfig.CAMTASIA_LAUNCH_POLL_INTERVAL_MS)
            for _ in range(wait_cycles):
                time.sleep(AppConfig.CAMTASIA_LAUNCH_POLL_INTERVAL_MS / 1000)
                hwnd = find_camtasia_window()
                if hwnd: break
        if not hwnd:
            self.finished.emit(False)
            return

        if self.action == AppConfig.ACTION_START:
            self.wm.click_camtasia_fullscreen(hwnd)
            template_name = AppConfig.CAMTASIA_REC_TEMPLATE
        elif self.action == AppConfig.ACTION_STOP:
            template_name = AppConfig.CAMTASIA_STOP_TEMPLATE
        elif self.action == AppConfig.ACTION_RESUME:
            template_name = AppConfig.CAMTASIA_RESUME_TEMPLATE
        else:
            self.finished.emit(False)
            return

        coords = self.wm.find_template(os.path.join(AppConfig.TEMPLATES_DIR, template_name))
        if coords:
            success = self.wm.robust_click(coords[0], coords[1], hwnd=hwnd)
            self.finished.emit(success)
        else:
            self.finished.emit(False)

class CamtasiaHotkeyWorker(QObject):
    finished = pyqtSignal(bool)

    def __init__(self, wm, action):
        super().__init__()
        self.wm = wm
        self.action = action

    def run(self):
        key_map = {
            AppConfig.ACTION_START: AppConfig.VK_F9,
            AppConfig.ACTION_STOP: AppConfig.VK_F10,
            AppConfig.ACTION_RESUME: AppConfig.VK_F9
        }
        key_code = key_map.get(self.action)
        if key_code:
            self.wm.press_key(key_code)
            time.sleep(AppConfig.KEY_PRESS_WAIT_AFTER)
            if self.action in [AppConfig.ACTION_START, AppConfig.ACTION_RESUME]:
                ok = self.wm.find_first_window_by_title(AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING) is not None
            elif self.action == AppConfig.ACTION_STOP:
                ok = self.wm.find_first_window_by_title(AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING) is None
            else:
                ok = False
            self.finished.emit(ok)
        else:
            self.finished.emit(False)

class MainWindow(QMainWindow):
    def __init__(self, splash):
        super().__init__()
        self.splash = splash
        self.notification_manager = NotificationManager()
        self.window_manager = WindowManager()
        self.update_manager = UpdateManager()
        self.telegram_notifier = TelegramNotifier(AppConfig.TELEGRAM_BOT_TOKEN, AppConfig.TELEGRAM_CHAT_ID)
        self.click_indicator = ClickIndicator()
        self._executor = ThreadPoolExecutor(max_workers=AppConfig.THREAD_POOL_WORKERS)

        self.current_project = None
        self.is_auto_record_enabled = True
        self.is_automation_enabled = True
        self.is_auto_popup_closing_enabled = False
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
        try:
            return self.screen() or QApplication.primaryScreen()
        except Exception:
            return QApplication.primaryScreen()

    def log(self, message: str, message_type: str):
        self.notification_manager.show(message, message_type)
        if message_type == 'error':
            self.telegram_notifier.send_message(AppConfig.MSG_TG_CRITICAL_ERROR.format(message))

    def _clear_layout(self):
        old_widget = self.centralWidget()
        if old_widget is not None:
            old_widget.setParent(None)
            old_widget.deleteLater()
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

    def _check_resume_result(self):
        hwnd = find_camtasia_window()
        status = win32gui.GetWindowText(hwnd) if hwnd else "NotFound"
        if AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in status.lower():
            self.hide_blocking_overlay()
            self.log("Снято с паузы успешно!", "info")
            return

        if getattr(self, '_camtasia_attempt_idx', 0) < AppConfig.CAMTASIA_RESUME_CHECK_ATTEMPTS:
            self._camtasia_attempt_idx += 1
            self.log(f"Не удалось снять с паузы, повторная попытка ({self._camtasia_attempt_idx})...", "warning")
            hwnd = find_camtasia_window()
            if hwnd:
                self._do_camtasia_click(os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_RESUME_TEMPLATE), hwnd)
            QTimer.singleShot(AppConfig.CAMTASIA_RESUME_CHECK_INTERVAL, self._check_resume_result)
            return

        if not getattr(self, '_camtasia_hotkey_tried', False):
            self._camtasia_hotkey_tried = True
            self.log("Не удаётся снять с паузы по кнопке, пробую хоткей...", "warning")
            self.window_manager.press_key(AppConfig.VK_F9)
            QTimer.singleShot(AppConfig.CAMTASIA_HOTKEY_WAIT_INTERVAL, self._check_resume_result)
            return

        self.hide_blocking_overlay()
        self.log("Camtasia: не удалось снять с паузы!", "error")

    def perform_camtasia_action(self, action: str):
        #self.show_blocking_overlay(AppConfig.MSG_CAMTASIA_AUTOMATION)
        self._camtasia_action = action
        self._camtasia_attempt_idx = 0
        self._camtasia_hotkey_tried = False
        self._try_camtasia_action_step()

    def _try_camtasia_action_step(self):
        hwnd = find_camtasia_window()
        status = win32gui.GetWindowText(hwnd) if hwnd else "NotFound"
        logging.info(f"[Camtasia] Attempt {self._camtasia_attempt_idx+1}: window status = '{status}'")

        if AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in status.lower():
            if self._camtasia_action == AppConfig.ACTION_START:
                self.hide_blocking_overlay()
                self.log("Camtasia: запись активна!", "info")
                return

        if AppConfig.CAMTASIA_WINDOW_TITLE_PAUSED in status.lower():
            if self._camtasia_action != AppConfig.ACTION_STOP:
                self.log("Camtasia на паузе. Пробую Resume...", "warning")
                self._do_camtasia_click(os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_RESUME_TEMPLATE), hwnd)
                QTimer.singleShot(AppConfig.CAMTASIA_RESUME_CHECK_INTERVAL, self._check_resume_result)
                return

        if not hwnd:
            self.log("Camtasia не найдена. Пробую ярлык...", "warning")
            try_launch_camtasia_shortcut()
            QTimer.singleShot(AppConfig.CAMTASIA_SYNC_RESTART_DELAY, self._try_camtasia_action_step)
            return

        if self._camtasia_action == AppConfig.ACTION_START:
            focus_camtasia_window()
            self.window_manager.click_camtasia_fullscreen(hwnd)
            self._do_camtasia_click(os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_REC_TEMPLATE), hwnd)
        elif self._camtasia_action == AppConfig.ACTION_STOP:
            focus_camtasia_window()
            self._do_camtasia_click(os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_STOP_TEMPLATE), hwnd)
        elif self._camtasia_action == AppConfig.ACTION_RESUME:
            focus_camtasia_window()
            self._do_camtasia_click(os.path.join(AppConfig.TEMPLATES_DIR, AppConfig.CAMTASIA_RESUME_TEMPLATE), hwnd)

        self._camtasia_attempt_idx += 1

        if self._camtasia_attempt_idx < AppConfig.CAMTASIA_MAX_ACTION_ATTEMPTS:
            QTimer.singleShot(AppConfig.CAMTASIA_ACTION_RETRY_INTERVAL, self._try_camtasia_action_step)
        else:
            if not self._camtasia_hotkey_tried:
                #self.log("Camtasia: попытки исчерпаны. Пробую hotkey...", "warning")
                self._camtasia_hotkey_tried = True
                self._try_camtasia_hotkey()
            else:
                self.hide_blocking_overlay()
                self.log("Camtasia: не удалось выполнить действие!", "error")

    def _do_camtasia_click(self, template_path, hwnd):
        coords = self.window_manager.find_template(template_path)
        if coords:
            success = self.window_manager.robust_click(coords[0], coords[1], hwnd=hwnd)
            logging.info(f"[Camtasia] Click on '{os.path.basename(template_path)}' ({coords})")
            #self.log(f"Клик по '{os.path.basename(template_path)}' — координаты: {coords}, успех: {success}", "info")
        else:
            #self.log(f"Шаблон '{os.path.basename(template_path)}' не найден, клик невозможен.", "warning")
            pass

    def _try_camtasia_hotkey(self):
        key_map = {
            AppConfig.ACTION_START: AppConfig.VK_F9,
            AppConfig.ACTION_STOP: AppConfig.VK_F10,
            AppConfig.ACTION_RESUME: AppConfig.VK_F9
        }
        key_code = key_map.get(self._camtasia_action)
        if key_code:
            self.window_manager.press_key(key_code)
            #self.log("Горячая клавиша отправлена.", "warning")
            QTimer.singleShot(AppConfig.CAMTASIA_HOTKEY_WAIT_INTERVAL, self._try_camtasia_action_step)
        else:
            self.hide_blocking_overlay()
            #self.log("Неизвестное действие для hotkey!", "error")

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

            manual_select_label = QLabel(AppConfig.LOBBY_MSG_MANUAL_SELECT_PROMPT)
            manual_select_label.setStyleSheet(StyleSheet.LOBBY_LABEL)
            manual_select_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

            self.project_combo = QComboBox()
            self.project_combo.setStyleSheet(StyleSheet.COMBO_BOX)
            self.project_combo.addItem(AppConfig.LOBBY_MSG_COMBO_PLACEHOLDER)
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
            QTimer.singleShot(0, lambda: self.on_project_changed(project_name))

    def build_project_ui(self):
        if not self.current_project: return
        try:
            self._clear_layout()
            if self.current_project == AppConfig.PROJECT_GG:
                self.setFixedSize(AppConfig.GG_UI_WIDTH, AppConfig.GG_UI_HEIGHT)
                main_layout = QHBoxLayout(self.central_widget)
                main_layout.setContentsMargins(12, 10, 12, 10)
                main_layout.setSpacing(15)
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
            if self.current_project != AppConfig.PROJECT_GG:
                toggles_layout.setColumnStretch(0, 1)
            main_layout.addWidget(toggles_frame)

            if self.current_project == AppConfig.PROJECT_GG:
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
            buttons_layout.addWidget(self.arrange_tables_button)
            buttons_layout.addWidget(self.arrange_system_button)
            if self.current_project == AppConfig.PROJECT_GG:
                self.close_tables_button = QPushButton(AppConfig.MSG_CLICK_COMMAND)
                buttons_layout.addWidget(self.close_tables_button)
                main_layout.addWidget(buttons_frame, 1)
            else:
                main_layout.addWidget(buttons_frame)
                main_layout.addStretch(1)

            # --- Progress Bar ---
            self.progress_frame = QFrame()
            if self.current_project == AppConfig.PROJECT_GG:
                progress_layout = QVBoxLayout(self.progress_frame)
            else:
                progress_layout = QHBoxLayout(self.progress_frame)
            progress_layout.setContentsMargins(0,0,0,0)
            self.progress_bar = AnimatedProgressBar()
            self.progress_bar.setFixedHeight(AppConfig.PROGRESS_BAR_HEIGHT)
            self.progress_bar_label = QLabel()
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)
            if self.current_project == AppConfig.PROJECT_GG:
                progress_layout.addWidget(self.progress_bar)
                progress_layout.addWidget(self.progress_bar_label)
                main_layout.addWidget(self.progress_frame, 1)
            else:
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
            buttons_to_style = [self.arrange_tables_button, self.arrange_system_button]
            if hasattr(self, 'close_tables_button'):
                buttons_to_style.append(self.close_tables_button)

            self.arrange_tables_button.setStyleSheet(StyleSheet.get_button_style(primary=True))
            self.arrange_system_button.setStyleSheet(StyleSheet.get_button_style(primary=False))
            if hasattr(self, 'close_tables_button'):
                self.close_tables_button.setStyleSheet(StyleSheet.get_button_style(primary=False))

            for btn in buttons_to_style:
                if btn:
                    shadow = QGraphicsDropShadowEffect()
                    shadow.setBlurRadius(15)
                    shadow.setColor(QColor(ColorPalette.BUTTON_SHADOW))
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
        if hasattr(self, 'close_tables_button'):
            self.close_tables_button.setEnabled(is_project_active)

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
        if hasattr(self, 'close_tables_button'):
            self.close_tables_button.clicked.connect(self.close_all_tables)

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

    def show_blocking_overlay(self, message=AppConfig.MSG_BLOCKING_OVERLAY_DEFAULT):
        #self.overlay = BlockingOverlay(message)
        #self.overlay.show()
        #QApplication.processEvents()
        pass

    def hide_blocking_overlay(self):
        if hasattr(self, 'overlay') and self.overlay:
            self.overlay.close()
            self.overlay = None

    def show_main_window_and_start_logic(self):
        self.splash.close()
        self.build_lobby_ui()
        self.show()
        self.position_window_default()
        self.start_main_logic()
        QTimer.singleShot(AppConfig.CAMTASIA_SYNC_RESTART_DELAY, self.initial_recorder_sync_check)

    def start_main_logic(self):
        if hasattr(self, 'lobby_status_label'):
            self.lobby_status_label.setText(AppConfig.LOBBY_MSG_SEARCHING)
        self.check_for_player()
        self.check_for_recorder()
        self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        if AppConfig.TELEGRAM_REPORT_LEVEL == 'all':
            self.telegram_notifier.send_message(AppConfig.MSG_TG_HELPER_STARTED)
        self.check_system_uptime()
        self.check_admin_rights()

    def on_project_changed(self, new_project_name: Optional[str]):
        if self.current_project == new_project_name: return

        if self.current_project == AppConfig.PROJECT_GG and hasattr(self, 'close_tables_button'):
            del self.close_tables_button

        self.current_project = new_project_name
        self.last_table_count = 0
        self.timers["auto_arrange"].stop()
        self.timers["popup_check"].stop()
        self.update_window_title()

        if new_project_name:
            self.build_project_ui()
            self.timers["auto_arrange"].start(AppConfig.AUTO_ARRANGE_INTERVAL)

            if new_project_name == AppConfig.PROJECT_QQ:
                self.position_window_top_right()
                self.timers["popup_check"].start(AppConfig.POPUP_CHECK_INTERVAL_FAST)
            elif new_project_name == AppConfig.PROJECT_GG:
                self.position_gg_panel()
                QTimer.singleShot(AppConfig.INJECTOR_MINIMIZE_DELAY, self.minimize_injector_window)
                self.timers["popup_check"].start(AppConfig.POPUP_CHECK_INTERVAL_FAST)

            if self.is_automation_enabled:
                if new_project_name == AppConfig.PROJECT_QQ: self.check_and_launch_opencv_server()
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
            if self.player_was_open:
                self.player_was_open = False
                self.handle_player_close()
            elif self.is_automation_enabled:
                self.check_and_launch_player()
        else:
            self.player_was_open = True
            project_name = None
            try:
                title = win32gui.GetWindowText(player_hwnd)
                project_name = next((short for full, short in AppConfig.PROJECT_MAPPING.items() if f"[{full}]" in title), None)
            except Exception:
                project_name = None

            if project_name:
                if self.timers["player_start"].isActive():
                    self.timers["player_start"].stop()
                    self.log("Авто-старт плеера успешно завершен.", "info")
                self.on_project_changed(project_name)
            else:
                if self.current_project is not None: self.on_project_changed(None)
                if hasattr(self, 'lobby_status_label'): self.lobby_status_label.setText(AppConfig.LOBBY_MSG_WAITING_FOR_LAUNCHER)
                if self.is_automation_enabled and not self.timers["player_start"].isActive():
                    self.log("Лаунчер найден. Запускаюсь...", "info")
                    self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
                    self.attempt_player_start_click()
        if not self.timers["player_check"].isActive():
            self.timers["player_check"].start(AppConfig.PLAYER_CHECK_INTERVAL)

    def check_admin_rights(self):
        try:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                self.log(AppConfig.MSG_ADMIN_WARNING, "warning")
        except Exception as e:
            logging.error("Не удалось проверить права администратора", exc_info=True)

    def check_system_uptime(self):
        try:
            uptime_days = ctypes.windll.kernel32.GetTickCount64() / (1000 * 60 * 60 * 24)
            if uptime_days > AppConfig.UPTIME_WARN_DAYS:
                self.log(AppConfig.MSG_UPTIME_WARNING, "warning")
        except Exception as e:
            logging.error("Не удалось проверить время работы системы", exc_info=True)

    def handle_player_close(self):
        if not self.is_automation_enabled: return
        self.is_sending_logs = True
        self.log("Плеер закрыт. Отправляю логи...", "info")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        launched_count = 0
        try:
            shortcuts_to_launch = [(k, os.path.join(desktop_path, f)) for f in os.listdir(desktop_path) if f.lower().endswith(AppConfig.SHORTCUT_EXTENSION) for k in AppConfig.LOG_SENDER_KEYWORDS if k in f.lower()]
        except Exception as e:
            self.log(f"Ошибка при поиске батника логов: {e}", "error")
            shortcuts_to_launch = []

        if not shortcuts_to_launch:
            self.log("Батники для отправки логов не найдены.", "warning")
        else:
            for keyword, path in shortcuts_to_launch:
                if self.window_manager.find_first_window_by_title(keyword):
                    self.log(f"Скрипт '{keyword}' уже запущен.", "info")
                else:
                    try:
                        os.startfile(path)
                        launched_count += 1
                        self.log(f"Запущен батник отправки логов '{keyword}'.", "info")
                    except Exception as e:
                        self.log(f"Не удалось запустить ярлык для '{keyword}': {e}", "error")
        if launched_count > 0:
            #self.log(f"Всего запущено: {launched_count}", "info")
            pass
        self.log(f"Перезапуск плеера через {AppConfig.PLAYER_RELAUNCH_DELAY_S} секунд...", "info")
        QTimer.singleShot(AppConfig.PLAYER_RELAUNCH_DELAY_S * 1000, lambda: self.wait_for_logs_to_finish(time.monotonic()))

    def wait_for_logs_to_finish(self, start_time: float):
        if (time.monotonic() - start_time) > AppConfig.LOG_SENDER_TIMEOUT_S:
            self.log("Тайм-аут ожидания отправки логов. Возобновление работы.", "error")
            self.is_sending_logs = False
            self.check_for_player()
            return
        if any(self.window_manager.find_first_window_by_title(k) for k in AppConfig.LOG_SENDER_KEYWORDS):
            self.log("Ожидание завершения отправки логов...", "info")
            QTimer.singleShot(AppConfig.LOG_WAIT_RETRY_INTERVAL, lambda: self.wait_for_logs_to_finish(start_time))
        else:
            self.log("Отправка логов завершена.", "info")
            self.is_sending_logs = False
            self.check_for_player()

    def check_and_launch_player(self):
        if self.is_launching_player or self.is_sending_logs: return
        if self.window_manager.find_first_window_by_title(AppConfig.PLAYER_GAME_LAUNCHER_TITLE) or self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_PROCESS_NAME):
            self.log("Обнаружен процесс запуска/обновления плеера. Ожидание...", "info")
            return
        self.log(f"Плеер не найден, ищу ярлык '{AppConfig.PLAYER_LAUNCHER_SHORTCUT_KEYWORD}'...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.PLAYER_LAUNCHER_SHORTCUT_KEYWORD in f.lower() and f.lower().endswith(AppConfig.SHORTCUT_EXTENSION)), None)
            if shortcut_path:
                self.log("Найден ярлык плеера. Запускаю...", "info")
                os.startfile(shortcut_path)
                self.is_launching_player = True
                QTimer.singleShot(AppConfig.LAUNCHER_WINDOW_ACTIVATION_TIMEOUT, lambda: setattr(self, 'is_launching_player', False))
                self.timers["player_start"].start(AppConfig.PLAYER_AUTOSTART_INTERVAL)
            else:
                self.log(f"Ярлык плеера '{AppConfig.PLAYER_LAUNCHER_SHORTCUT_KEYWORD}' на рабочем столе не найден.", "error")
        except Exception as e:
            self.log(f"Ошибка при поиске/запуске ярлыка: {e}", "error")

    def attempt_player_start_click(self):
        if not self.is_automation_enabled:
            self.timers["player_start"].stop()
            return
        launcher_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if not launcher_hwnd:
            self.log("Не удалось найти окно лаунчера для авто-старта.", "warning")
            self.timers["player_start"].stop()
            return
        try:
            self.focus_window(launcher_hwnd)
            rect = win32gui.GetWindowRect(launcher_hwnd)
            x, y = rect[0] + 50, rect[1] + 50
            self.log("Попытка авто-старта плеера...", "info")
            self.window_manager.robust_click(x, y, hwnd=launcher_hwnd, log_prefix="PlayerStart")
        except Exception as e:
            self.log(f"Не удалось активировать окно плеера: {e}", "error")

    def check_and_launch_opencv_server(self):
        if not self.is_automation_enabled: return
        config = PROJECT_CONFIGS.get(AppConfig.PROJECT_QQ)
        if not config or self.window_manager.find_windows_by_config(config, AppConfig.KEY_CV_SERVER, self.winId()): return
        self.log("Сервер OpenCV не найден, ищу ярлык...", "warning")
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.OPENCV_SHORTCUT_KEYWORD in f.lower() and f.lower().endswith(AppConfig.SHORTCUT_EXTENSION)), None)
            if shortcut_path:
                self.log("Найден ярлык OpenCV. Запускаю...", "info")
                os.startfile(shortcut_path)
                QTimer.singleShot(AppConfig.OPENCV_LAUNCH_ARRANGE_DELAY, self.arrange_other_windows)
            else:
                self.log("Ярлык для OpenCV на рабочем столе не найден.", "error")
        except Exception as e:
            self.log(f"Ошибка при поиске/запуске ярлыка OpenCV: {e}", "error")

    def check_for_recorder(self):
        if self.is_launching_recorder or self.window_manager.is_process_running(AppConfig.CAMTASIA_PROCESS_NAME):
            if self.timers["recorder_check"].isActive(): self.timers["recorder_check"].stop()
            return
        if self.is_automation_enabled:
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            try:
                shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.CAMTASIA_SHORTCUT_KEYWORD in f.lower() and f.lower().endswith(AppConfig.SHORTCUT_EXTENSION)), None)
                if shortcut_path:
                    self.log("Camtasia не найдена. Запускаю...", "warning")
                    os.startfile(shortcut_path)
                    self.is_launching_recorder = True
                    QTimer.singleShot(10000, lambda: setattr(self, 'is_launching_recorder', False))
                    self.timers["recorder_check"].stop()
                    QTimer.singleShot(3000, self.check_for_recorder)
                else:
                    self.log("Ярлык для Camtasia на рабочем столе не найден.", "error")
            except Exception as e:
                self.log(f"Ошибка при поиске/запуске Camtasia: {e}", "error")
        if not self.timers["recorder_check"].isActive():
            self.timers["recorder_check"].start(AppConfig.RECORDER_CHECK_INTERVAL)

    def initial_recorder_sync_check(self):
        hwnd = find_camtasia_window()
        is_recording = False
        if hwnd:
            title = win32gui.GetWindowText(hwnd)
            is_recording = AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in title.lower()
        if is_recording:
            self.log("Обнаружена активная запись. Перезапускаю для синхронизации...", "warning")
            self.stop_recording_session()
            QTimer.singleShot(AppConfig.CAMTASIA_SYNC_RESTART_DELAY, self.check_auto_record_logic)

    def toggle_auto_record(self):
        self.is_auto_record_enabled = not self.is_auto_record_enabled
        self.log(AppConfig.MSG_AUTO_RECORD_ON if self.is_auto_record_enabled else AppConfig.MSG_AUTO_RECORD_OFF, "info")
        if self.is_auto_record_enabled:
            self.timers["auto_record"].start(AppConfig.AUTO_RECORD_INTERVAL)
        else:
            self.timers["auto_record"].stop()
        self.auto_record_toggle.setChecked(self.is_auto_record_enabled)

    def toggle_automation(self):
        self.is_automation_enabled = not self.is_automation_enabled
        self.log(AppConfig.MSG_AUTOMATION_ON if self.is_automation_enabled else AppConfig.MSG_AUTOMATION_OFF, "info")
        if self.is_automation_enabled:
            self.check_for_player()
            self.check_for_recorder()
        self.automation_toggle.setChecked(self.is_automation_enabled)

    def toggle_auto_popup_closing(self):
        self.is_auto_popup_closing_enabled = not self.is_auto_popup_closing_enabled
        self.log(AppConfig.MSG_POPUP_CLOSER_ON if self.is_auto_popup_closing_enabled else AppConfig.MSG_POPUP_CLOSER_OFF, "info")
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

            # === ClubGG: таймер выключения 5 минут ===
            if self.current_project == AppConfig.PROJECT_GG:
                is_clubgg_visible = False
                for hwnd in self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId()):
                    try:
                        if win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd):
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                            win32api.CloseHandle(h_process)
                            if process_name.lower() == AppConfig.CLUBGG_PROCESS_NAME:
                                is_clubgg_visible = True
                                break
                    except Exception:
                        continue

                is_chrome_visible = False
                chrome_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CHROME_PROCESS_NAME, check_visible=True)
                if chrome_hwnd and win32gui.IsWindowVisible(chrome_hwnd) and not win32gui.IsIconic(chrome_hwnd):
                    is_chrome_visible = True

                should_be_recording = is_clubgg_visible or is_chrome_visible
                if should_be_recording:
                    trigger_reason = []
                    if is_clubgg_visible:
                        trigger_reason.append("ClubGG (видимое окно)")
                    if is_chrome_visible:
                        trigger_reason.append("Chrome (видимое окно)")
                    trigger_reason = " и ".join(trigger_reason)

                # --- Логика с таймером выключения ---
                if not hasattr(self, "last_should_record_false_time"):
                    self.last_should_record_false_time = None

                hwnd = find_camtasia_window()
                is_recording, is_paused = False, False
                if hwnd:
                    title = win32gui.GetWindowText(hwnd).lower()
                    is_recording = AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in title
                    is_paused = AppConfig.CAMTASIA_WINDOW_TITLE_PAUSED in title

                if should_be_recording:
                    self.last_should_record_false_time = None
                    if not (is_recording or is_paused):
                        self.log(f"Начинаю автозапись (обнаружен {trigger_reason})...", "info")
                        self.start_recording_session()
                    elif is_paused:
                        self.log("Запись на паузе. Возобновляю...", "warning")
                        self.perform_camtasia_action(AppConfig.ACTION_RESUME)
                else:
                    now = time.monotonic()
                    if self.last_should_record_false_time is None:
                        self.last_should_record_false_time = now
                    if (now - self.last_should_record_false_time) >= AppConfig.AUTO_STOP_RECORD_INACTIVITY_S:
                        if is_recording or is_paused:
                            self.log(f"Нет активности >{AppConfig.AUTO_STOP_RECORD_INACTIVITY_S / 60:.0f} минут. Останавливаю запись...", "info")
                            self.stop_recording_session()

                self.chrome_was_visible = is_chrome_visible

            # === QQ: выключение моментально, если нет окон ===
            elif self.current_project == AppConfig.PROJECT_QQ:
                is_qq_visible = False
                for hwnd in self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId()):
                    try:
                        if win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd):
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                            win32api.CloseHandle(h_process)
                            is_qq_visible = True
                            break
                    except Exception:
                        continue

                is_chrome_visible = False
                chrome_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CHROME_PROCESS_NAME, check_visible=True)
                if chrome_hwnd and win32gui.IsWindowVisible(chrome_hwnd) and not win32gui.IsIconic(chrome_hwnd):
                    is_chrome_visible = True

                should_be_recording = is_qq_visible or is_chrome_visible
                trigger_reason = []
                if is_qq_visible:
                    trigger_reason.append("QQPK (видимый стол)")
                if is_chrome_visible:
                    trigger_reason.append("Chrome (видимое окно)")
                trigger_reason = " и ".join(trigger_reason) if trigger_reason else None

                hwnd = find_camtasia_window()
                is_recording, is_paused = False, False
                if hwnd:
                    title = win32gui.GetWindowText(hwnd).lower()
                    is_recording = AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in title
                    is_paused = AppConfig.CAMTASIA_WINDOW_TITLE_PAUSED in title

                if should_be_recording:
                    if not (is_recording or is_paused):
                        self.log(f"Начинаю автозапись (обнаружен {trigger_reason})...", "info")
                        self.start_recording_session()
                    elif is_paused:
                        self.log("Запись на паузе. Возобновляю...", "warning")
                        self.perform_camtasia_action(AppConfig.ACTION_RESUME)
                else:
                    if is_recording or is_paused:
                        self.log(f"Нет активных окон QQPK/Chrome. Останавливаю запись...", "info")
                        self.stop_recording_session()

            # === Другие проекты (оставить как раньше) ===
            else:
                if self.window_manager.find_windows_by_config(config, AppConfig.KEY_LOBBY, self.winId()) \
                        or self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId()):
                    should_be_recording = True
                    trigger_reason = f"активность {self.current_project}"

                hwnd = find_camtasia_window()
                is_recording, is_paused = False, False
                if hwnd:
                    title = win32gui.GetWindowText(hwnd).lower()
                    is_recording = AppConfig.CAMTASIA_WINDOW_TITLE_RECORDING in title
                    is_paused = AppConfig.CAMTASIA_WINDOW_TITLE_PAUSED in title

                if should_be_recording:
                    if not (is_recording or is_paused):
                        self.log(f"Начинаю автозапись (обнаружен {trigger_reason})...", "info")
                        self.start_recording_session()
                    elif is_paused:
                        self.log("Запись на паузе. Возобновляю...", "warning")
                        self.perform_camtasia_action(AppConfig.ACTION_RESUME)
                else:
                    if is_recording or is_paused:
                        self.log(f"Нет активности. Останавливаю запись...", "info")
                        self.stop_recording_session()

        except Exception as e:
            self.log(f"Ошибка в логике автозаписи: {e}", "error")
            logging.error("Ошибка в check_auto_record_logic", exc_info=True)


    def start_recording_session(self):
        self.perform_camtasia_action(AppConfig.ACTION_START)
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        self.recording_start_time = time.monotonic()
        self.timers["session"].start(AppConfig.SESSION_PROGRESS_UPDATE_INTERVAL)
        self.progress_bar.setMaximum(config[AppConfig.KEY_SESSION_MAX_S])
        self.progress_bar.setValue(0)
        self.update_project_ui_state()

    def stop_recording_session(self):
        self.perform_camtasia_action(AppConfig.ACTION_STOP)
        self.timers["session"].stop()
        self.recording_start_time = 0
        self.is_record_stopping = True
        self.timers["record_cooldown"].start(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000)
        self.update_project_ui_state()

    def update_session_progress(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or AppConfig.KEY_SESSION_MAX_S not in config: return
        elapsed = time.monotonic() - self.recording_start_time
        self.progress_bar.setValue(elapsed)
        if elapsed >= config[AppConfig.KEY_SESSION_MAX_S]:
            self.progress_bar_label.setText(AppConfig.MSG_LIMIT_REACHED)
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL + f" color: {ColorPalette.RED}; font-weight: bold;")
            QTimer.singleShot(AppConfig.SESSION_LIMIT_HANDLER_DELAY, self.handle_session_limit_reached)
        else:
            remaining_s = max(0, config[AppConfig.KEY_SESSION_MAX_S] - elapsed)
            self.progress_bar_label.setText(AppConfig.MSG_PROGRESS_LABEL.format(time.strftime('%H:%M:%S', time.gmtime(remaining_s))))
            self.progress_bar_label.setStyleSheet(StyleSheet.PROGRESS_BAR_LABEL)

    def handle_session_limit_reached(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        self.log(f"{config[AppConfig.KEY_SESSION_MAX_S]/3600:.0f} часа записи истекли. Перезапуск...", "info")
        self.stop_recording_session()
        QTimer.singleShot(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000 + 1000, self.check_auto_record_logic)

    def check_for_new_tables(self):
        if not self.is_automation_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        if self.current_project == AppConfig.PROJECT_GG and not self.window_manager.is_process_running(AppConfig.CLUBGG_PROCESS_NAME):
            self.last_table_count = 0
            return
        current_tables = self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId())
        current_count = len(current_tables)
        if current_count != self.last_table_count:
            self.log(f"Изменилось количество столов: {self.last_table_count} -> {current_count}. Перерасстановка...", "info")
            QTimer.singleShot(AppConfig.TABLE_ARRANGE_ON_CHANGE_DELAY, self.arrange_tables)
        self.last_table_count = current_count

    def arrange_tables(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return

        self.last_arrangement_time = time.monotonic()
        if self.current_project in [AppConfig.PROJECT_QQ, AppConfig.PROJECT_GG] and self.timers["popup_check"].interval() == AppConfig.POPUP_CHECK_INTERVAL_SLOW:
            self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_FAST)
            self.log("Ускорен поиск попапов после расстановки.", "info")

        if self.current_project == AppConfig.PROJECT_GG:
            self.arrange_gg_tables()
            return

        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or AppConfig.KEY_TABLE not in config:
            self.log(f"Нет конфига столов для {self.current_project}.", "warning")
            return
        found_windows = self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId())
        if not found_windows:
            self.log(AppConfig.MSG_ARRANGE_TABLES_NOT_FOUND, "warning")
            return

        titles = [win32gui.GetWindowText(hwnd) for hwnd in found_windows if win32gui.IsWindow(hwnd)]
        logging.info(f"Найдены окна для расстановки ({len(titles)}): {titles}")

        if self.current_project == AppConfig.PROJECT_QQ and len(found_windows) > AppConfig.QQ_DYNAMIC_ARRANGE_THRESHOLD:
            self.arrange_dynamic_qq_tables(found_windows, config)
            return

        slots_key = AppConfig.KEY_TABLE_SLOTS_5 if self.current_project == AppConfig.PROJECT_QQ and len(found_windows) >= 5 else AppConfig.KEY_TABLE_SLOTS
        SLOTS = config[slots_key]
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS) or not win32gui.IsWindow(hwnd): continue
            x, y = SLOTS[i]
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try:
                rect = win32gui.GetWindowRect(hwnd)
                win32gui.MoveWindow(hwnd, x, y, rect[2] - rect[0], rect[3] - rect[1], True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось разместить стол {i+1}: {e}", "error")
        if arranged_count > 0:
            self.log(f"Расставлено столов: {arranged_count}", "info")

    def arrange_gg_tables(self):
        config = PROJECT_CONFIGS.get(AppConfig.PROJECT_GG)
        if not config:
            self.log("GG конфиг не найден!", "error")
            return

        all_hwnds = self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId())
        gg_tables = []
        for hwnd in all_hwnds:
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
                win32api.CloseHandle(h_process)
                if process_name.lower() == AppConfig.CLUBGG_PROCESS_NAME:
                    gg_tables.append(hwnd)
            except Exception as e:
                logging.debug(f"Ошибка фильтрации clubgg.exe hwnd={hwnd}: {e}")
        found_windows = gg_tables[:4]

        slots = config.get(AppConfig.KEY_TABLE_SLOTS, [])
        base_w, base_h = config[AppConfig.KEY_TABLE][AppConfig.KEY_WIDTH], config[AppConfig.KEY_TABLE][AppConfig.KEY_HEIGHT]
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(slots): break
            x, y = slots[i]
            try:
                if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.MoveWindow(hwnd, x, y, base_w, base_h, True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось разместить GG стол {i+1}: {e}", "error")
        if arranged_count > 0:
            self.log(f"GG столы расставлены по фиксированным слотам ({arranged_count}).", "info")

    def arrange_dynamic_qq_tables(self, found_windows, config):
        max_tables = len(found_windows)
        screen_geo = self.get_current_screen().availableGeometry()
        base_width = config[AppConfig.KEY_TABLE][AppConfig.KEY_WIDTH]
        base_height = config[AppConfig.KEY_TABLE][AppConfig.KEY_HEIGHT]
        tables_per_row = min(max_tables, AppConfig.QQ_DYNAMIC_TABLES_PER_ROW)
        rows = (max_tables + tables_per_row - 1) // tables_per_row
        new_width = max(int(base_width * AppConfig.QQ_DYNAMIC_SCALE_FACTOR), screen_geo.width() // tables_per_row)
        new_height = max(int(base_height * AppConfig.QQ_DYNAMIC_SCALE_FACTOR), screen_geo.height() // rows)
        overlap_x = AppConfig.QQ_DYNAMIC_OVERLAP_FACTOR if tables_per_row > 1 else 1.0
        overlap_y = AppConfig.QQ_DYNAMIC_OVERLAP_FACTOR if rows > 1 else 1.0
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            row, col = i // tables_per_row, i % tables_per_row
            x = screen_geo.left() + int(col * new_width * overlap_x)
            y = screen_geo.top() + int(row * new_height * overlap_y)
            if not win32gui.IsWindow(hwnd): continue
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try:
                win32gui.MoveWindow(hwnd, x, y, new_width, new_height, True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось динамически разместить стол {i+1}: {e}", "error")
        if arranged_count > 0:
            self.log(f"Динамически расставлено столов: {arranged_count}", "info")

    def arrange_other_windows(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config:
            self.log(f"Нет конфига для {self.current_project}.", "warning")
            return

        self.position_player_window(config)
        if self.current_project == AppConfig.PROJECT_GG:
            self.minimize_injector_window()
        self.position_recorder_window()
        if self.current_project == AppConfig.PROJECT_GG:
            self.position_lobby_window(config)
        elif self.current_project == AppConfig.PROJECT_QQ:
            self.position_cv_server_window(config)
        self.log("Системные окна расставлены.", "info")

    def position_window(self, hwnd: Optional[int], x: int, y: int, w: int, h: int, log_fail: str):
        if hwnd and win32gui.IsWindow(hwnd):
            try:
                win32gui.MoveWindow(hwnd, x, y, w, h, True)
            except Exception as e:
                logging.error(f"Ошибка позиционирования окна: {log_fail}", exc_info=True)
        else:
            self.log(log_fail, "warning")

    def position_player_window(self, config):
        player_config = config.get(AppConfig.KEY_PLAYER, {})
        if player_config:
            hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
            self.position_window(hwnd, player_config[AppConfig.KEY_X], player_config[AppConfig.KEY_Y], player_config[AppConfig.KEY_WIDTH], player_config[AppConfig.KEY_HEIGHT], "Плеер не найден.")

    def position_lobby_window(self, config):
        cfg = config.get(AppConfig.KEY_LOBBY, {})
        if cfg:
            lobbies = self.window_manager.find_windows_by_config(config, AppConfig.KEY_LOBBY, self.winId())
            self.position_window(lobbies[0] if lobbies else None, cfg[AppConfig.KEY_X], cfg[AppConfig.KEY_Y], cfg[AppConfig.KEY_WIDTH], cfg[AppConfig.KEY_HEIGHT], "Лобби не найдено.")

    def position_cv_server_window(self, config):
        cfg = config.get(AppConfig.KEY_CV_SERVER, {})
        if cfg:
            cv_windows = self.window_manager.find_windows_by_config(config, AppConfig.KEY_CV_SERVER, self.winId())
            self.position_window(cv_windows[0] if cv_windows else None, cfg[AppConfig.KEY_X], cfg[AppConfig.KEY_Y], cfg[AppConfig.KEY_WIDTH], cfg[AppConfig.KEY_HEIGHT], "CV Сервер не найден.")

    def position_recorder_window(self):
        recorder_hwnd = find_camtasia_window()
        if recorder_hwnd and win32gui.IsWindow(recorder_hwnd):
            try:
                if win32gui.IsIconic(recorder_hwnd):
                    win32gui.ShowWindow(recorder_hwnd, win32con.SW_RESTORE)
                    time.sleep(AppConfig.ROBUST_CLICK_ACTIVATION_DELAY)
                screen_rect = self.get_current_screen().availableGeometry()
                rect = win32gui.GetWindowRect(recorder_hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                x = screen_rect.left() + (screen_rect.width() - w) // 2
                y = screen_rect.bottom() - h
                win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
            except Exception as e:
                self.log(f"Ошибка позиционирования Camtasia: {e}", "error")

    def minimize_injector_window(self):
        injector_hwnd = self.window_manager.find_first_window_by_title(AppConfig.INJECTOR_WINDOW_TITLE, exact_match=False)
        if injector_hwnd and win32gui.IsWindow(injector_hwnd):
            try:
                win32gui.ShowWindow(injector_hwnd, win32con.SW_MINIMIZE)
                self.log(f"Окно '{AppConfig.INJECTOR_WINDOW_TITLE}' свернуто.", "info")
            except Exception as e:
                self.log(f"Не удалось свернуть окно '{AppConfig.INJECTOR_WINDOW_TITLE}': {e}", "error")
        else:
            logging.warning(f"Окно '{AppConfig.INJECTOR_WINDOW_TITLE}' для сворачивания не найдено.")

    def focus_window(self, hwnd: int):
        try:
            if self.shell: self.shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(AppConfig.ROBUST_CLICK_DELAY)
        except Exception as e:
            logging.error(f"Не удалось сфокусировать окно {hwnd}", exc_info=True)

    def close_all_tables(self):
        if not self.current_project:
            self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error")
            return
        self.log(f"Закрываю все столы для проекта {self.current_project}...", "info")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        found_windows = self.window_manager.find_windows_by_config(config, AppConfig.KEY_TABLE, self.winId())
        if not found_windows:
            self.log("Столы для закрытия не найдены.", "warning")
            return
        closed_count = sum(1 for hwnd in found_windows if self.window_manager.close_window(hwnd))
        if closed_count > 0:
            self.log(f"Закрыто столов: {closed_count}", "info")

    def check_for_popups(self):
        if not self.is_automation_enabled or not self.current_project or not self.is_auto_popup_closing_enabled: return

        if self.last_arrangement_time > 0:
            elapsed = time.monotonic() - self.last_arrangement_time
            current_interval = self.timers["popup_check"].interval()
            if elapsed > AppConfig.POPUP_FAST_SCAN_DURATION_S:
                if current_interval != AppConfig.POPUP_CHECK_INTERVAL_SLOW:
                    self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_SLOW)
                    self.log(f"Замедлен поиск попапов для {self.current_project}.", "info")
            elif current_interval != AppConfig.POPUP_CHECK_INTERVAL_FAST:
                self.timers["popup_check"].setInterval(AppConfig.POPUP_CHECK_INTERVAL_FAST)

        if self.current_project == AppConfig.PROJECT_QQ:
            self._handle_qq_popups()
        elif self.current_project == AppConfig.PROJECT_GG:
            self._handle_gg_popups()

    def _handle_qq_popups(self):
        popup_config = PROJECT_CONFIGS.get(AppConfig.PROJECT_QQ, {}).get(AppConfig.KEY_POPUPS)
        if not popup_config: return
        templates_dir = self._get_templates_dir(AppConfig.PROJECT_QQ)
        if not templates_dir: return

        wm = self.window_manager
        confidence = popup_config.get(AppConfig.KEY_CONFIDENCE)

        for rule in popup_config.get(AppConfig.KEY_SPAM, []):
            poster_path = os.path.join(templates_dir, rule["poster"])
            if os.path.exists(poster_path) and wm.find_template(poster_path, confidence=confidence):
                self.log(f"Обнаружен спам '{rule['poster']}'. Ищу кнопку '{rule['close_button']}'...", "info")
                time.sleep(AppConfig.POPUP_ACTION_DELAY)
                close_btn_path = os.path.join(templates_dir, rule['close_button'])
                if wm.find_and_click_template(close_btn_path, confidence=confidence):
                    self.log(f"Успешно закрыт спам с помощью '{rule['close_button']}'.", "info")
                else:
                    self.log(f"Не удалось найти кнопку закрытия '{rule['close_button']}'.", "warning")
                return

        bonus_rule = popup_config.get(AppConfig.KEY_BONUS)
        if not bonus_rule: return
        wheel_path = os.path.join(templates_dir, bonus_rule['wheel'])
        if os.path.exists(wheel_path) and wm.find_template(wheel_path, confidence=confidence):
            self.log("Обнаружено бонусное колесо. Ищу кнопку вращения...", "info")
            time.sleep(AppConfig.POPUP_ACTION_DELAY)
            spin_btn_path = os.path.join(templates_dir, bonus_rule['spin_button'])
            if wm.find_and_click_template(spin_btn_path, confidence=confidence):
                self.log("Бонус: нажата кнопка вращения. Ожидание...", "info")
                time.sleep(AppConfig.BONUS_SPIN_WAIT_S)
                close_bonus_btn_path = os.path.join(templates_dir, bonus_rule['close_button'])
                if wm.find_and_click_template(close_bonus_btn_path, confidence=confidence):
                    self.log("Бонусное окно успешно закрыто.", "info")
                else:
                    self.log("Не удалось найти кнопку закрытия бонусного окна.", "warning")
            else:
                self.log("Не удалось найти кнопку вращения бонуса.", "warning")
            return

    def _handle_gg_popups(self):
        popup_config = PROJECT_CONFIGS.get(AppConfig.PROJECT_GG, {}).get(AppConfig.KEY_POPUPS)
        if not popup_config: return
        templates_dir = self._get_templates_dir(AppConfig.PROJECT_GG)
        if not templates_dir: return

        wm = self.window_manager
        confidence = popup_config.get(AppConfig.KEY_CONFIDENCE)
        buyin_rule = popup_config.get(AppConfig.KEY_BUY_IN)
        if not buyin_rule: return

        trigger_path = os.path.join(templates_dir, buyin_rule[AppConfig.KEY_POPUP_TRIGGER])
        if os.path.exists(trigger_path) and wm.find_template(trigger_path, confidence=confidence):
            self.log("Обнаружено окно Buy-in. Выполняю авто-нажатия...", "info")
            time.sleep(AppConfig.POPUP_ACTION_DELAY)

            max_btn_path = os.path.join(templates_dir, buyin_rule[AppConfig.KEY_POPUP_MAX_BTN])
            if wm.find_and_click_template(max_btn_path, confidence=confidence):
                self.log("Нажата кнопка 'Max'.", "info")
                time.sleep(AppConfig.POPUP_ACTION_DELAY)
                confirm_btn_path = os.path.join(templates_dir, buyin_rule[AppConfig.KEY_POPUP_CONFIRM_BTN])
                if wm.find_and_click_template(confirm_btn_path, confidence=confidence):
                    self.log("Нажата кнопка подтверждения Buy-in.", "info")
                else:
                    self.log("Не удалось найти кнопку подтверждения Buy-in.", "warning")
            else:
                self.log("Не удалось найти кнопку 'Max'.", "warning")
            return

    def _get_templates_dir(self, project_name: str) -> Optional[str]:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        templates_dir = os.path.join(base_path, AppConfig.TEMPLATES_DIR, project_name)
        if not os.path.isdir(templates_dir):
            if self.timers["popup_check"].isActive():
                self.log(f"Папка шаблонов {templates_dir} не найдена. Отключаю проверку.", "warning")
                self.timers["popup_check"].stop()
            return None
        return templates_dir

    def position_window_top_right(self):
        try:
            screen = self.get_current_screen()
            geo = screen.availableGeometry()
            self.move(geo.right() - self.frameGeometry().width() - AppConfig.WINDOW_MARGIN, geo.top() + AppConfig.WINDOW_MARGIN)
        except Exception as e:
            logging.error("Could not position window top-right", exc_info=True)

    def position_window_default(self):
        try:
            screen = self.get_current_screen()
            geo = screen.availableGeometry()
            self.move(geo.left() + AppConfig.WINDOW_MARGIN, geo.bottom() - self.frameGeometry().height() - AppConfig.WINDOW_MARGIN)
        except Exception as e:
            logging.error("Could not position window default", exc_info=True)

    def position_gg_panel(self):
        try:
            screen = self.get_current_screen()
            geo = screen.availableGeometry()
            self.move(geo.left() + AppConfig.WINDOW_MARGIN, geo.bottom() - self.frameGeometry().height() - AppConfig.WINDOW_MARGIN)
        except Exception as e:
            logging.error("Could not position GG panel", exc_info=True)

def try_launch_camtasia_shortcut():
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    try:
        shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.CAMTASIA_SHORTCUT_KEYWORD in f.lower() and f.lower().endswith(AppConfig.SHORTCUT_EXTENSION)), None)
        if shortcut_path:
            os.startfile(shortcut_path)
            logging.info("Запущен ярлык Camtasia.")
            return True
        else:
            logging.warning("Ярлык Camtasia на рабочем столе не найден.")
    except Exception as e:
        logging.error(f"Ошибка при запуске ярлыка Camtasia: {e}")
    return False

class SingleInstance:
    """Класс для проверки и удержания мьютекса, чтобы предотвратить запуск второй копии приложения."""
    def __init__(self, name):
        self.mutex = windll.kernel32.CreateMutexW(None, False, name)
        self.last_error = windll.kernel32.GetLastError()
    
    def is_already_running(self):
        return self.last_error == 183 # ERROR_ALREADY_EXISTS

# ===================================================================
# 6. ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ===================================================================
if __name__ == '__main__':
    instance = SingleInstance(AppConfig.MUTEX_NAME)
    if instance.is_already_running():
        logging.warning("Обнаружена уже запущенная копия OiHelper. Завершение работы.")
        sys.exit(0)

    app = QApplication(sys.argv)
    splash = SplashScreen()
    splash.show()

    def on_update_progress(value, stage):
        splash.set_progress(value, stage)
        if stage == AppConfig.SPLASH_MSG_ERROR:
            splash.hide_progress()

    main_window = MainWindow(splash)
    main_window.update_manager.progress_changed.connect(on_update_progress)
    sys.exit(app.exec())
