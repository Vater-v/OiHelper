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
import tkinter as tk
from tkinter import ttk, scrolledtext
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
    logging.warning("PyAutoGUI не найден. Установите его ('pip install pyautogui') для дополнительного метода клика.")

# ===================================================================
# 0. НАСТРОЙКА ЛОГИРОВАНИЯ
# ===================================================================
try:
    log_dir = os.path.join(os.getenv('APPDATA'), 'OiHelper')
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, 'app.log')
    file_handler = logging.FileHandler(log_file_path, encoding='utf-8', mode='w')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    root_logger = logging.getLogger()
    root_logger.addHandler(file_handler)
    root_logger.setLevel(logging.INFO)
except Exception as e:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.error(f"Не удалось настроить логирование в файл: {e}")

# ===================================================================
# 1. КОНФИГУРАЦИЯ
# ===================================================================

class AppConfig:
    CURRENT_VERSION = "4.3"
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
    LOG_SENDER_TIMEOUT_S = 120
    PLAYER_RELAUNCH_DELAY_S = 5
    RECORD_RESTART_COOLDOWN_S = 5

    DEFAULT_WIDTH = 450
    DEFAULT_HEIGHT = 350
    WINDOW_MARGIN = 10

    APP_TITLE = f"OiHelper v{CURRENT_VERSION}"
    STATUS_MSG_PLAYER_NOT_FOUND = "Статус: Плеер не найден"
    STATUS_MSG_PRESS_START = "Статус: Ожидание старта..."
    STATUS_MSG_OK = "Статус: Все системы в норме"
    MSG_ARRANGE_TABLES = "Расставить столы"
    MSG_ARRANGE_SYSTEM = "Системные окна"
    MSG_CLICK_COMMAND = "Sit-Out всем"
    MSG_PROGRESS_LABEL = "Перезапись через: {}"
    MSG_LIMIT_REACHED = "Лимит!"
    MSG_UPDATE_CHECK = "Проверка обновлений..."
    MSG_UPDATE_FAIL = "Ошибка обновления. Работа в оффлайн-режиме."
    MSG_UPTIME_WARNING = "Компьютер не перезагружался более 5 дней."
    MSG_ADMIN_WARNING = "Нет прав администратора. Функции могут быть ограничены."
    MSG_ARRANGE_TABLES_NOT_FOUND = "Столы для расстановки не найдены."
    MSG_PROJECT_UNDEFINED = "Проект не определен."

    PLAYER_CHECK_INTERVAL = 2000
    PLAYER_AUTOSTART_INTERVAL = 4000
    AUTO_RECORD_INTERVAL = 3000
    AUTO_ARRANGE_INTERVAL = 2000
    RECORDER_CHECK_INTERVAL = 10000
    STATUS_MESSAGE_DURATION = 3000

    FOCUS_DELAY = 0.25
    CLICK_DELAY_ACTION = 0.08


PROJECT_CONFIGS = {
    "GG": { "PROCESS_NAME": "clubgg.exe", "TABLE": { "FIND_METHOD": "RATIO", "W": 557, "H": 424, "TOLERANCE": 0.035 }, "LOBBY": { "FIND_METHOD": "RATIO", "W": 333, "H": 623, "TOLERANCE": 0.07, "X": 1580, "Y": 140 }, "PLAYER": { "W": 700, "H": 365, "X": 1385, "Y": 0 }, "TABLE_SLOTS": [(-5, 0), (271, 423), (816, 0), (1086, 423)], "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "notepad.exe"], "SESSION_MAX_DURATION_S": 4 * 3600, "SESSION_WARN_TIME_S": 3.5 * 3600, "ARRANGE_MINIMIZED_TABLES": False },
    "QQ": { "PROCESS_NAME": "qqpoker.exe", "TABLE": { "FIND_METHOD": "TITLE_AND_SIZE", "TITLE": "QQPK", "W": 400, "H": 700, "TOLERANCE": 2 }, "LOBBY": { "FIND_METHOD": "RATIO", "W": 400, "H": 700, "TOLERANCE": 0.07, "X": 1418, "Y": 0 }, "CV_SERVER": { "FIND_METHOD": "TITLE", "TITLE": "OpenCv", "X": 1789, "Y": 367, "W": 993, "H": 605 }, "PLAYER": { "X": 1418, "Y": 942, "W": 724, "H": 370 }, "TABLE_SLOTS": [(0, 0), (401, 0), (802, 0), (1203, 0)], "TABLE_SLOTS_5": [(0, 0), (346, 0), (692, 0), (1038, 0), (1384, 0)], "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"], "EXCLUDED_PROCESSES": ["explorer.exe", "svchost.exe", "cmd.exe", "powershell.exe", "Taskmgr.exe", "chrome.exe", "firefox.exe", "msedge.exe", "RuntimeBroker.exe", "ApplicationFrameHost.exe", "SystemSettings.exe", "NekoRay.exe", "nekobox.exe", "Sandbo.exe", "OpenCvServer.exe", "notepad.exe"], "SESSION_MAX_DURATION_S": 3 * 3600, "SESSION_WARN_TIME_S": -1, "ARRANGE_MINIMIZED_TABLES": True }
}

# ===================================================================
# 2. UI КОМПОНЕНТЫ И МЕНЕДЖЕРЫ
# ===================================================================

class ClickIndicator(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-transparentcolor", "white")
        self.config(bg="white")
        
        self.canvas = tk.Canvas(self, width=50, height=50, bg="white", highlightthickness=0)
        self.canvas.pack()
        self.canvas.create_oval(5, 5, 45, 45, fill="#FF0000", outline="")

    def show_at(self, x, y):
        self.geometry(f"+{x - 25}+{y - 25}")
        self.deiconify()
        self.after(400, self.withdraw)

class WindowManager:
    def __init__(self, app):
        self.app = app
        self.shell = win32com.client.Dispatch("WScript.Shell") if WIN32COM_AVAILABLE else None

    def get_process_name_from_hwnd(self, hwnd):
        if not hwnd or not win32gui.IsWindow(hwnd): return None
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            if not h_process: return None
            process_name = os.path.basename(win32process.GetModuleFileNameEx(h_process, 0))
            win32api.CloseHandle(h_process)
            return process_name.lower()
        except Exception:
            return None

    def force_foreground(self, hwnd):
        try:
            foreground = win32gui.GetForegroundWindow()
            if hwnd == foreground:
                return True
            fg_thread, _ = win32process.GetWindowThreadProcessId(foreground)
            target_thread, _ = win32process.GetWindowThreadProcessId(hwnd)
            ctypes.windll.user32.AttachThreadInput(fg_thread, target_thread, True)
            win32gui.SetForegroundWindow(hwnd)
            ctypes.windll.user32.AttachThreadInput(fg_thread, target_thread, False)
            return win32gui.GetForegroundWindow() == hwnd
        except Exception as e:
            self.app.log(f"Ошибка AttachThreadInput: {e}", "error")
            return False

    def focus_window(self, hwnd):
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
            if self.shell:
                self.shell.SendKeys('%')
                time.sleep(0.1)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(AppConfig.FOCUS_DELAY)
            if win32gui.GetForegroundWindow() == hwnd:
                return True
            self.app.log(f"Вторая попытка фокуса на hwnd={hwnd}", "info")
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            if win32gui.GetForegroundWindow() == hwnd:
                return True
            self.app.log(f"Не удалось подтвердить фокус на окне {hwnd}", "warning")
            return False
        except Exception as e:
            self.app.log(f"Критическая ошибка при фокусировке окна {hwnd}: {e}", "error")
            return False

    def _sendinput_click(self, screen_x, screen_y):
        try:
            screen_w, screen_h = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
            abs_x = int(screen_x * 65535 / (screen_w - 1))
            abs_y = int(screen_y * 65535 / (screen_h - 1))
            class MOUSEINPUT(ctypes.Structure): _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long), ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
            class _INPUT_UNION(ctypes.Union): _fields_ = [("mi", MOUSEINPUT)]
            class INPUT(ctypes.Structure): _fields_ = [("type", ctypes.c_ulong), ("u", _INPUT_UNION)]
            INPUT_MOUSE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP = 0, 0x8000, 0x0002, 0x0004
            mi_down = MOUSEINPUT(abs_x, abs_y, 0, MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN, 0, None)
            inp_down = INPUT(type=INPUT_MOUSE, u=_INPUT_UNION(mi=mi_down))
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(inp_down))
            time.sleep(AppConfig.CLICK_DELAY_ACTION)
            mi_up = MOUSEINPUT(abs_x, abs_y, 0, MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTUP, 0, None)
            inp_up = INPUT(type=INPUT_MOUSE, u=_INPUT_UNION(mi=mi_up))
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(inp_up))
            return True
        except Exception as e:
            self.app.log(f"SendInput не удался: {e}", "error")
            return False

    def robust_click(self, hwnd, client_x, client_y, check_success_func=None):
        self.app.log(f"НАЖИМАЮ (AttachThreadInput) по hwnd={hwnd} ({client_x},{client_y})", "info")
        
        # 1. Попытка жестко передать фокус окну:
        if not self.force_foreground(hwnd):
            self.app.log("Принудительная фокусировка не удалась, пробуем старый метод...", "warning")
            # Fallback to the old focus method
            if not self.focus_window(hwnd):
                 self.app.log("Фокусировка не удалась всеми методами, клик отменен.", "error")
                 return False
            
        try:
            rect = win32gui.GetWindowRect(hwnd)
            screen_x, screen_y = rect[0] + client_x, rect[1] + client_y
            self.app.click_indicator.show_at(screen_x, screen_y)
        except Exception as e:
            self.app.log(f"Не удалось получить геометрию окна: {e}", "error")
            return False

        # Явная задержка для отработки фокуса
        time.sleep(0.15)

        # Основной клик — SendInput
        click_ok = self._sendinput_click(screen_x, screen_y)
        time.sleep(0.25)
        
        if not click_ok:
            self.app.log("SendInput не удался, fallback mouse_event.", "warning")
            try:
                win32api.SetCursorPos((screen_x, screen_y))
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, screen_x, screen_y, 0, 0)
                time.sleep(0.11)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, screen_x, screen_y, 0, 0)
            except Exception as e:
                self.app.log(f"mouse_event тоже не удался: {e}", "error")

        # Проверка успеха
        if check_success_func:
            if check_success_func():
                self.app.log("Клик подтвержден!", "info")
                return True
            self.app.log("Клик не прошел (по функции проверки)", "error")
            return False
        else:
            self.app.log("Клик выполнен (без проверки)", "info")
            return True

    def find_windows_by_config(self, config, config_key, main_window_hwnd):
        window_config = config.get(config_key, {})
        if not window_config: return []
        find_method = window_config.get("FIND_METHOD"); EXCLUDED_TITLES = config.get("EXCLUDED_TITLES", []); EXCLUDED_PROCESSES = config.get("EXCLUDED_PROCESSES", []); arrange_minimized = config.get("ARRANGE_MINIMIZED_TABLES", False) and config_key == "TABLE"
        found_windows = []
        def enum_windows_callback(hwnd, _):
            if hwnd == main_window_hwnd or not win32gui.IsWindowVisible(hwnd) or (not arrange_minimized and win32gui.IsIconic(hwnd)): return True
            try:
                process_name = self.get_process_name_from_hwnd(hwnd)
                if process_name and process_name in EXCLUDED_PROCESSES: return True
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
        except Exception as e: self.app.log(f"Ошибка сортировки окон: {e}", "warning")
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
                    process_name = self.get_process_name_from_hwnd(hwnd)
                    if process_name and process_name_to_find.lower() in process_name: hwnds.append(hwnd)
                except Exception: pass
        try: win32gui.EnumWindows(callback, None)
        except Exception as e: logging.error(f"Критическая ошибка в EnumWindows (поиск по процессу): {e}", exc_info=True)
        return hwnds[0] if hwnds else None

    def press_key(self, key_code):
        try:
            win32api.keybd_event(key_code, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e:
            self.app.log(f"Ошибка эмуляции нажатия: {e}", "error")

class TelegramNotifier:
    def __init__(self, token, chat_id):
        self.token = token; self.chat_id = chat_id; self.api_url = f"https://api.telegram.org/bot{self.token}/sendMessage"; self.message_queue = queue.Queue(); self.worker_thread = threading.Thread(target=self._worker, daemon=True); self.worker_thread.start()
    def send_message(self, message):
        if not self.token or not self.chat_id: logging.warning("Токен или ID чата Telegram не настроены."); return
        self.message_queue.put(message)
    def _worker(self):
        while True:
            message = self.message_queue.get()
            try: requests.post(self.api_url, data={'chat_id': self.chat_id, 'text': message}, timeout=10).raise_for_status()
            except requests.RequestException as e: logging.error(f"Не удалось отправить сообщение в Telegram: {e}")
            self.message_queue.task_done()

class UpdateManager:
    def __init__(self, app):
        self.app = app
        self.update_info = {}

    def is_new_version_available(self, current_v_str, latest_v_str):
        try:
            current = [int(p) for p in current_v_str.lstrip('v').split('.')]
            latest = [int(p) for p in latest_v_str.lstrip('v').split('.')]
            max_len = max(len(current), len(latest))
            current += [0] * (max_len - len(current)); latest += [0] * (max_len - len(latest))
            return latest > current
        except Exception as e:
            logging.error(f"Ошибка сравнения версий: {e}.")
            return latest_v_str > current_v_str

    def check_for_updates(self):
        self.app.log(AppConfig.MSG_UPDATE_CHECK, "info")
        try:
            api_url = f"https://api.github.com/repos/{AppConfig.GITHUB_REPO}/releases/latest"; response = requests.get(api_url, timeout=10); response.raise_for_status(); latest_release = response.json()
            if (latest_version := latest_release.get("tag_name")) and self.is_new_version_available(AppConfig.CURRENT_VERSION, latest_version):
                self.app.log(f"Доступна новая версия: {latest_version}. Обновление...", "info"); self.update_info = latest_release; threading.Thread(target=self.apply_update, daemon=True).start()
            else:
                self.app.log("Вы используете последнюю версию.", "info")
                self.app.after(100, self.app.finish_startup)
        except requests.RequestException as e:
            self.app.log(f"Ошибка проверки обновлений: {e}", "error")
            self.app.after(100, self.app.handle_update_failure)
        except Exception as e:
            logging.error(f"Неожиданная ошибка при проверке обновлений: {e}", exc_info=True)
            self.app.after(100, self.app.handle_update_failure)

    def apply_update(self):
        download_url = next((asset["browser_download_url"] for asset in self.update_info.get("assets", []) if asset["name"] == AppConfig.ASSET_NAME), None)
        if not download_url: self.app.log("Не удалось найти ZIP-архив в релизе.", "error"); return
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url):
        update_zip_name = "update.zip"
        try:
            self.app.log("Скачивание обновления...", "info"); response = requests.get(url, stream=True, timeout=60); response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
            self.app.log("Распаковка архива...", "info"); update_folder = "update_temp"
            if os.path.isdir(update_folder): import shutil; shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref: zip_ref.extractall(update_folder)
            self.app.log("Обновление скачано. Перезапуск...", "info"); updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]); current_dir = os.path.dirname(current_exe_path); exe_name = os.path.basename(current_exe_path)
            script_content = f'@echo off\nchcp 65001 > NUL\necho Waiting for OiHelper to close...\ntimeout /t 1 /nobreak > NUL\ntaskkill /pid {os.getpid()} /f > NUL\necho Waiting for process to terminate...\ntimeout /t 2 /nobreak > NUL\necho Moving new files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\nrd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\ndel "{current_dir}\\{update_zip_name}"\necho Starting new version...\nstart "" "{exe_name}"\n(goto) 2>nul & del "%~f0"'
            with open(updater_script_path, "w", encoding="cp866") as f: f.write(script_content)
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE); self.app.on_closing()
        except Exception as e:
            self.app.log(f"Ошибка при обновлении: {e}", "error"); logging.error(f"Update error: {e}", exc_info=True)
            if os.path.exists(update_zip_name):
                try: os.remove(update_zip_name)
                except OSError as err: logging.error(f"Не удалось удалить временный файл обновления: {err}")
            self.app.log(AppConfig.MSG_UPDATE_FAIL, "warning")
            self.app.after(100, self.app.handle_update_failure)

# ===================================================================
# 4. ГЛАВНОЕ ОКНО ПРИЛОЖЕНИЯ (TKINTER)
# ===================================================================
class OiHelperApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.current_project = None
        self.last_table_count = 0
        self.recording_start_time = 0
        self.is_sending_logs = False
        self.player_was_open = False
        self.is_record_stopping = False
        self.timers = {}

        self.window_manager = WindowManager(self)
        self.update_manager = UpdateManager(self)
        self.telegram_notifier = TelegramNotifier(AppConfig.TELEGRAM_BOT_TOKEN, AppConfig.TELEGRAM_CHAT_ID)
        self.click_indicator = ClickIndicator(self)

        self.title(AppConfig.APP_TITLE)
        if os.path.exists(AppConfig.ICON_PATH): self.iconbitmap(AppConfig.ICON_PATH)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Стили
        self.style = ttk.Style(self)
        self.style.theme_use('vista')
        BG_COLOR, TEXT_COLOR, PRIMARY_COLOR, PRIMARY_ACTIVE, STATUS_TEXT, LOG_BG, WARN, ERROR, SECONDARY_BG, SECONDARY_ACTIVE = "#F0F0F0", "#000000", "#0078D7", "#005A9E", "#444444", "#FFFFFF", '#E67E22', '#D32F2F', "#E1E1E1", "#CACACA"
        self.configure(bg=BG_COLOR)
        self.style.configure('.', background=BG_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 9))
        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('TLabel', background=BG_COLOR, foreground=TEXT_COLOR)
        self.style.configure('Title.TLabel', font=("Segoe UI", 12, "bold"))
        self.style.configure('Status.TLabel', foreground=STATUS_TEXT)
        
        # Настройка кнопок с фикс. цветом и без обводки фокуса
        self.style.configure('TButton', font=("Segoe UI", 9, "bold"), padding=5, background=SECONDARY_BG, foreground='black', borderwidth=1, focusthickness=0)
        self.style.map('TButton',
                       background=[('active', SECONDARY_ACTIVE)],
                       focuscolor=[('!disabled', self.style.lookup('TButton', 'background'))]
        )
        self.style.configure('Primary.TButton', background=PRIMARY_COLOR, foreground="white", borderwidth=1, focusthickness=0)
        self.style.map('Primary.TButton',
                       background=[('active', PRIMARY_ACTIVE)],
                       focuscolor=[('!disabled', self.style.lookup('Primary.TButton', 'background'))]
        )
        
        self.style.configure('TCheckbutton', background=BG_COLOR)
        self.style.configure('Horizontal.TProgressbar', troughcolor='#E0E0E0', background=PRIMARY_COLOR, thickness=10)
        self.log_widget_colors = {'info': TEXT_COLOR, 'warning': WARN, 'error': ERROR, 'bg': LOG_BG}

        self.is_automation_enabled = tk.BooleanVar(value=True)
        self.is_auto_record_enabled = tk.BooleanVar(value=True)
        self.status_text = tk.StringVar(value=AppConfig.STATUS_MSG_OK)
        self.progress_text = tk.StringVar(value="")

        self.loading_frame = None
        self.main_frame = None
        self.init_startup_checks()

    def on_closing(self):
        logging.info("Приложение закрывается, остановка таймеров...")
        for timer_id in self.timers.values():
            if timer_id:
                try: self.after_cancel(timer_id)
                except Exception: pass
        self.destroy()

    def log(self, message, message_type):
        log_levels = {'info': logging.INFO, 'warning': logging.WARNING, 'error': logging.ERROR}
        logging.log(log_levels.get(message_type, logging.INFO), message)
        if message_type == 'error' and AppConfig.TELEGRAM_REPORT_LEVEL == 'all':
            self.telegram_notifier.send_message(f"OiHelper Критическая ошибка: {message}")
        
        if not hasattr(self, 'log_widget') or not self.log_widget.winfo_exists():
            return
            
        try:
            timestamp = time.strftime('%H:%M:%S')
            full_message = f"[{timestamp}] {message}\n"
            self.log_widget.config(state='normal')
            self.log_widget.insert(tk.END, full_message, message_type)
            self.log_widget.config(state='disabled')
            self.log_widget.see(tk.END)
        except tk.TclError as e:
            print(f"Ошибка обновления UI лога (возможно, окно закрывается): {e}")

    def set_status_message(self, message, is_persistent=False):
        self.status_text.set(message)
        if hasattr(self, 'status_message_timer_id') and self.status_message_timer_id:
            self.after_cancel(self.status_message_timer_id)
        if not is_persistent:
            self.status_message_timer_id = self.after(AppConfig.STATUS_MESSAGE_DURATION, lambda: self.status_text.set(AppConfig.STATUS_MSG_OK))

    def create_loading_ui(self):
        if self.loading_frame: self.loading_frame.destroy()
        self.loading_frame = ttk.Frame(self, padding="20")
        self.loading_frame.pack(fill="both", expand=True)
        ttk.Label(self.loading_frame, text="Проверка обновлений...", font=("Segoe UI", 12)).pack(pady=10)
        self.loading_bar = ttk.Progressbar(self.loading_frame, mode='indeterminate')
        self.loading_bar.pack(fill='x', padx=10, pady=5)
        self.loading_bar.start(10)

    def create_main_widgets(self):
        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.pack(fill="both", expand=True)

        top_panel = ttk.Frame(self.main_frame)
        top_panel.pack(fill='x', expand=False)
        
        self.title_label = ttk.Label(top_panel, text=AppConfig.APP_TITLE, style='Title.TLabel')
        self.title_label.pack(side='left', padx=(0,10))
        
        ttk.Separator(top_panel, orient='horizontal').pack(fill='x', pady=5, side='bottom')
        
        controls_frame = ttk.Frame(self.main_frame)
        controls_frame.pack(fill='x', pady=5)

        toggles_frame = ttk.Frame(controls_frame)
        toggles_frame.pack(side='left', padx=(0, 20))
        ttk.Checkbutton(toggles_frame, text="Автоматика", variable=self.is_automation_enabled, command=self.toggle_automation).pack(anchor='w')
        ttk.Checkbutton(toggles_frame, text="Автозапись", variable=self.is_auto_record_enabled, command=self.toggle_auto_record).pack(anchor='w')

        buttons_frame = ttk.Frame(controls_frame)
        buttons_frame.pack(side='left', fill='x', expand=True)
        self.arrange_tables_button = ttk.Button(buttons_frame, text=AppConfig.MSG_ARRANGE_TABLES, command=self.arrange_tables, style='Primary.TButton')
        self.arrange_tables_button.pack(fill='x', pady=2)
        self.arrange_system_button = ttk.Button(buttons_frame, text=AppConfig.MSG_ARRANGE_SYSTEM, command=self.arrange_other_windows)
        self.arrange_system_button.pack(fill='x', pady=2)
        self.sit_out_button = ttk.Button(buttons_frame, text=AppConfig.MSG_CLICK_COMMAND, command=self.perform_special_clicks)
        
        status_bar_frame = ttk.Frame(self.main_frame)
        status_bar_frame.pack(fill='x', side='bottom', pady=(10, 0))
        self.status_label = ttk.Label(status_bar_frame, textvariable=self.status_text, style='Status.TLabel')
        self.status_label.pack(fill='x')
        self.progress_frame = ttk.Frame(status_bar_frame)
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate', style='Horizontal.TProgressbar')
        self.progress_bar.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.progress_bar_label = ttk.Label(self.progress_frame, textvariable=self.progress_text, style='Status.TLabel')
        self.progress_bar_label.pack(side='right')

        log_frame = ttk.LabelFrame(self.main_frame, text="Лог событий")
        log_frame.pack(fill='both', expand=True, pady=5)
        self.log_widget = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, bg=self.log_widget_colors['bg'], fg=self.log_widget_colors['info'], font=("Consolas", 9), borderwidth=0, relief="solid")
        self.log_widget.pack(fill='both', expand=True, padx=2, pady=2)
        for tag, color in self.log_widget_colors.items():
            if tag != 'bg':
                self.log_widget.tag_config(tag, foreground=color)

    def sync_ui_state(self):
        if not self.main_frame: return
        
        title = AppConfig.APP_TITLE
        if self.current_project:
            title += f" [{self.current_project}]"
        self.title(title)
        self.title_label.config(text=title)

        self.geometry(f"{AppConfig.DEFAULT_WIDTH}x{AppConfig.DEFAULT_HEIGHT}")
        if self.current_project == "QQ":
            self.position_window_top_right()
        else:
            self.position_window_default()
        
        is_project_active = self.current_project is not None
        btn_state = 'normal' if is_project_active else 'disabled'
        self.arrange_tables_button.config(state=btn_state)
        self.arrange_system_button.config(state=btn_state)
        
        if self.current_project == "GG":
            self.sit_out_button.pack(fill='x', pady=2, before=self.arrange_system_button)
            self.sit_out_button.config(state=btn_state)
        else:
            self.sit_out_button.pack_forget()

        if self.recording_start_time > 0:
            self.progress_frame.pack(fill='x', expand=True, pady=(5,0), after=self.status_label)
        else:
            self.progress_frame.pack_forget()

    def init_startup_checks(self):
        self.create_loading_ui()
        threading.Thread(target=self.update_manager.check_for_updates, daemon=True).start()

    def handle_update_failure(self):
        if self.loading_frame: self.loading_frame.destroy()
        self.finish_startup()

    def finish_startup(self):
        if self.loading_frame:
            self.loading_bar.stop()
            self.loading_frame.destroy()
            self.loading_frame = None
        self.create_main_widgets()
        self.sync_ui_state()
        self.start_main_logic()

    def start_main_logic(self):
        self.log("Запуск основных таймеров...", "info")
        self.timers['player_check'] = self.after(AppConfig.PLAYER_CHECK_INTERVAL, self.check_for_player)
        self.timers['recorder_check'] = self.after(AppConfig.RECORDER_CHECK_INTERVAL, self.check_for_recorder)
        self.timers['auto_record'] = self.after(AppConfig.AUTO_RECORD_INTERVAL, self.check_auto_record_logic)
        self.timers['auto_arrange'] = self.after(AppConfig.AUTO_ARRANGE_INTERVAL, self.check_for_new_tables)
        if AppConfig.TELEGRAM_REPORT_LEVEL == 'all': self.telegram_notifier.send_message(f"OiHelper {AppConfig.CURRENT_VERSION} запущен.")
        self.check_system_uptime(); self.check_admin_rights(); self.log("Приложение запущено и готово к работе.", "info")

    def on_project_changed(self, new_project_name):
        if self.current_project == new_project_name: return
        self.current_project = new_project_name
        self.last_table_count = 0
        self.sync_ui_state()
        if self.is_automation_enabled.get() and new_project_name:
            self.arrange_other_windows()
            if new_project_name == "GG": self.minimize_injector_window()
            elif new_project_name == "QQ": self.check_and_launch_opencv_server()
    
    def check_for_player(self):
        if not self.is_sending_logs:
            player_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
            if not player_hwnd:
                if 'player_start' in self.timers and self.timers['player_start']:
                    self.after_cancel(self.timers['player_start'])
                    self.timers['player_start'] = None
                self.on_project_changed(None)
                if self.player_was_open:
                    self.player_was_open = False
                    self.handle_player_close()
                else:
                    self.set_status_message(AppConfig.STATUS_MSG_PLAYER_NOT_FOUND)
                    if self.is_automation_enabled.get(): self.check_and_launch_player()
            else:
                self.player_was_open = True
                try:
                    title = win32gui.GetWindowText(player_hwnd)
                    project_name = next((short for full, short in {"QQPoker": "QQ", "ClubGG": "GG"}.items() if f"[{full}]" in title), None)
                except Exception: project_name = None
                
                if project_name:
                    if 'player_start' in self.timers and self.timers['player_start']:
                        self.after_cancel(self.timers['player_start'])
                        self.timers['player_start'] = None
                        self.log("Авто-старт плеера успешно завершен.", "info")
                    self.on_project_changed(project_name)
                else:
                    self.on_project_changed(None)
                    self.set_status_message(AppConfig.STATUS_MSG_PRESS_START)
                    if self.is_automation_enabled.get() and not ('player_start' in self.timers and self.timers.get('player_start')):
                        self.log("Лаунчер найден. Включаю попытки авто-старта...", "info")
                        self.attempt_player_start_click()
        self.timers['player_check'] = self.after(AppConfig.PLAYER_CHECK_INTERVAL, self.check_for_player)

    def attempt_player_start_click(self):
        if not self.is_automation_enabled.get() or self.current_project:
            if 'player_start' in self.timers and self.timers.get('player_start'):
                self.after_cancel(self.timers['player_start'])
                self.timers['player_start'] = None
            return

        launcher_hwnd = self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE)
        if launcher_hwnd:
            def check_start_success():
                try:
                    current_title = win32gui.GetWindowText(launcher_hwnd)
                    return any(proj in current_title for proj in ["QQPoker", "ClubGG"])
                except Exception:
                    return not win32gui.IsWindow(launcher_hwnd)
            
            self.window_manager.robust_click(launcher_hwnd, 50, 50, check_success_func=check_start_success)
        
        self.timers['player_start'] = self.after(AppConfig.PLAYER_AUTOSTART_INTERVAL, self.attempt_player_start_click)

    def check_admin_rights(self):
        try:
            if ctypes.windll.shell32.IsUserAnAdmin() == 0: self.log(AppConfig.MSG_ADMIN_WARNING, "warning")
        except Exception as e: logging.error(f"Не удалось проверить права администратора: {e}")

    def check_system_uptime(self):
        try:
            if (ctypes.windll.kernel32.GetTickCount64() / (1000 * 60 * 60 * 24)) > 5: self.log(AppConfig.MSG_UPTIME_WARNING, "warning")
        except Exception as e: logging.error(f"Не удалось проверить время работы системы: {e}")

    def handle_player_close(self):
        if not self.is_automation_enabled.get(): return
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
        self.log(f"Перезапуск плеера через {AppConfig.PLAYER_RELAUNCH_DELAY_S} секунд...", "info"); self.after(AppConfig.PLAYER_RELAUNCH_DELAY_S * 1000, lambda: self.wait_for_logs_to_finish(time.monotonic()))

    def wait_for_logs_to_finish(self, start_time):
        if (time.monotonic() - start_time) > AppConfig.LOG_SENDER_TIMEOUT_S: self.log("Тайм-аут ожидания отправки логов. Возобновление работы.", "error"); self.is_sending_logs = False; return
        if any(self.window_manager.find_first_window_by_title(k) for k in AppConfig.LOG_SENDER_KEYWORDS): self.log("Ожидание завершения отправки логов...", "info"); self.after(3000, lambda: self.wait_for_logs_to_finish(start_time))
        else: self.log("Отправка логов завершена.", "info"); self.is_sending_logs = False

    def check_and_launch_player(self):
        if self.is_sending_logs: return
        if self.window_manager.find_first_window_by_title(AppConfig.PLAYER_GAME_LAUNCHER_TITLE) or self.window_manager.find_first_window_by_title("launch"): self.log("Обнаружен процесс запуска/обновления плеера. Ожидание...", "info"); return
        self.log("Плеер не найден, ищу ярлык 'launch'...", "warning"); desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'launch' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: self.log("Найден ярлык плеера. Запускаю...", "info"); os.startfile(shortcut_path)
            else: self.log("Ярлык плеера 'launch' на рабочем столе не найден.", "error")
        except Exception as e: self.log(f"Ошибка при поиске/запуске ярлыка: {e}", "error")

    def check_and_launch_opencv_server(self):
        if not self.is_automation_enabled.get(): return
        config = PROJECT_CONFIGS.get("QQ");
        if not config or self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winfo_id()): return
        self.log("Сервер OpenCV не найден, ищу ярлык...", "warning"); desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if 'opencv' in f.lower() and f.lower().endswith('.lnk')), None)
            if shortcut_path: self.log("Найден ярлык OpenCV. Запускаю...", "info"); os.startfile(shortcut_path); self.after(4000, self.arrange_other_windows)
            else: self.log("Ярлык для OpenCV на рабочем столе не найден.", "error")
        except Exception as e: self.log(f"Ошибка при поиске/запуске ярлыка OpenCV: {e}", "error")

    def check_for_recorder(self):
        if not self.window_manager.is_process_running(AppConfig.CAMTASIA_PROCESS_NAME):
            if self.is_automation_enabled.get():
                desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
                try:
                    shortcut_path = next((os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if AppConfig.CAMTASIA_PROCESS_NAME in f.lower() and f.lower().endswith('.lnk')), None)
                    if shortcut_path: self.log("Camtasia не найдена. Запускаю...", "warning"); os.startfile(shortcut_path)
                    else: self.log("Ярлык для Camtasia на рабочем столе не найден.", "error")
                except Exception as e: self.log(f"Ошибка при поиске/запуске Camtasia: {e}", "error")
        self.timers['recorder_check'] = self.after(AppConfig.RECORDER_CHECK_INTERVAL, self.check_for_recorder)

    def toggle_auto_record(self): self.log(f"Автозапись {'включена' if self.is_auto_record_enabled.get() else 'выключена'}.", "info")
    def toggle_automation(self):
        self.log(f"Автоматика {'включена' if self.is_automation_enabled.get() else 'выключена'}.", "info")
        if self.is_automation_enabled.get(): self.check_for_player(); self.check_for_recorder()

    def check_auto_record_logic(self):
        if self.is_auto_record_enabled.get() and self.current_project and not self.is_record_stopping:
            config = PROJECT_CONFIGS.get(self.current_project)
            if config:
                try:
                    should_be_recording = False
                    if self.current_project == "GG":
                        should_be_recording = self.window_manager.is_process_running("ClubGG.exe") or self.window_manager.is_process_running("chrome.exe")
                    else:
                        should_be_recording = bool(self.window_manager.find_windows_by_config(config, "LOBBY", self.winfo_id()) or self.window_manager.find_windows_by_config(config, "TABLE", self.winfo_id()))

                    is_recording = self.window_manager.find_first_window_by_title("Recording...") is not None
                    is_paused = self.window_manager.find_first_window_by_title("Paused...") is not None

                    if should_be_recording and not (is_recording or is_paused):
                        self.log("Начинаю автозапись...", "info"); self.start_recording_session()
                    elif should_be_recording and is_paused:
                        self.log("Запись на паузе. Возобновляю...", "warning"); self.perform_camtasia_action(win32con.VK_F9, "возобновление записи")
                    elif not should_be_recording and (is_recording or is_paused):
                        self.log("Активность завершена. Останавливаю запись...", "info"); self.stop_recording_session()
                except Exception as e:
                    self.log(f"Ошибка в логике автозаписи: {e}", "error")
        self.timers['auto_record'] = self.after(AppConfig.AUTO_RECORD_INTERVAL, self.check_auto_record_logic)

    def start_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F9, "старт записи")
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        self.recording_start_time = time.monotonic()
        self.timers['session'] = self.after(1000, self.update_session_progress)
        self.progress_bar.config(maximum=config["SESSION_MAX_DURATION_S"], value=0)
        self.sync_ui_state()

    def stop_recording_session(self):
        self.perform_camtasia_action(win32con.VK_F10, "остановку записи")
        if 'session' in self.timers and self.timers['session']:
            self.after_cancel(self.timers['session'])
            self.timers['session'] = None
        self.recording_start_time = 0
        self.is_record_stopping = True
        self.after(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000, lambda: setattr(self, 'is_record_stopping', False))
        self.sync_ui_state()

    def perform_camtasia_action(self, key_code, action_name):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME)
        if not recorder_hwnd:
            self.log(f"Не удалось выполнить '{action_name}': окно Camtasia не найдено.", "error")
            return
        self.log(f"Выполняю '{action_name}' для Camtasia...", "info")
        try:
            if action_name == "старт записи" and not self.window_manager.find_first_window_by_title("Recording..."):
                self.window_manager.robust_click(recorder_hwnd, 55, 80)
                time.sleep(0.5)
            if self.window_manager.focus_window(recorder_hwnd):
                self.window_manager.press_key(key_code)
                self.after(500, self.position_recorder_window)
            else:
                 self.log(f"Не удалось сфокусироваться на Camtasia для '{action_name}'.", "error")
        except Exception as e:
            self.log(f"Ошибка при взаимодействии с Camtasia: {e}", "error")

    def update_session_progress(self):
        if self.recording_start_time > 0:
            config = PROJECT_CONFIGS.get(self.current_project)
            if config and "SESSION_MAX_DURATION_S" in config:
                elapsed = time.monotonic() - self.recording_start_time
                self.progress_bar.config(value=elapsed)
                max_duration = config["SESSION_MAX_DURATION_S"]
                if elapsed >= max_duration:
                    self.progress_text.set(AppConfig.MSG_LIMIT_REACHED)
                    self.after(100, self.handle_session_limit_reached)
                else:
                    remaining_s = max(0, max_duration - elapsed)
                    self.progress_text.set(AppConfig.MSG_PROGRESS_LABEL.format(time.strftime('%H:%M:%S', time.gmtime(remaining_s))))
        self.timers['session'] = self.after(1000, self.update_session_progress)

    def handle_session_limit_reached(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        self.log(f"{config['SESSION_MAX_DURATION_S']/3600:.0f} часа записи истекли. Перезапуск...", "info")
        self.stop_recording_session()
        self.after(AppConfig.RECORD_RESTART_COOLDOWN_S * 1000 + 1000, self.check_auto_record_logic)

    def check_for_new_tables(self):
        if self.is_automation_enabled.get() and self.current_project:
            config = PROJECT_CONFIGS.get(self.current_project)
            if config:
                if self.current_project == "GG" and not self.window_manager.is_process_running("ClubGG.exe"):
                    self.last_table_count = 0
                else:
                    current_tables = self.window_manager.find_windows_by_config(config, "TABLE", self.winfo_id())
                    current_count = len(current_tables)
                    if current_count != self.last_table_count:
                        self.log(f"Изменилось количество столов: {self.last_table_count} -> {current_count}. Перерасстановка...", "info")
                        self.after(500, self.arrange_tables)
                    self.last_table_count = current_count
        self.timers['auto_arrange'] = self.after(AppConfig.AUTO_ARRANGE_INTERVAL, self.check_for_new_tables)

    def arrange_tables(self):
        if not self.current_project: self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error"); return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "TABLE" not in config: self.log(f"Нет конфига столов для {self.current_project}.", "warning"); return
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winfo_id())
        if not found_windows: self.log(AppConfig.MSG_ARRANGE_TABLES_NOT_FOUND, "warning"); return

        if self.current_project == "QQ" and len(found_windows) > 4:
            self.arrange_dynamic_qq_tables(found_windows, config)
            return

        slots_key = "TABLE_SLOTS_5" if self.current_project == "QQ" and len(found_windows) >= 5 else "TABLE_SLOTS"
        SLOTS = config[slots_key]; arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS) or not win32gui.IsWindow(hwnd): continue
            x, y = SLOTS[i]
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try:
                if self.current_project == "GG": win32gui.MoveWindow(hwnd, x, y, config["TABLE"]["W"], config["TABLE"]["H"], True)
                else: rect = win32gui.GetWindowRect(hwnd); win32gui.MoveWindow(hwnd, x, y, rect[2] - rect[0], rect[3] - rect[1], True)
                arranged_count += 1
            except Exception as e:
                self.log(f"Не удалось разместить стол {i+1}: {e}", "error")
        if arranged_count > 0: self.log(f"Расставлено столов: {arranged_count}", "info")

    def arrange_dynamic_qq_tables(self, found_windows, config):
        max_tables = len(found_windows); screen_width, screen_height = self.winfo_screenwidth(), self.winfo_screenheight()
        base_width, base_height = config["TABLE"]["W"], config["TABLE"]["H"]; tables_per_row = min(max_tables, 5)
        rows = (max_tables + tables_per_row - 1) // tables_per_row; new_width = max(int(base_width * 0.6), screen_width // tables_per_row)
        new_height = max(int(base_height * 0.6), screen_height // rows); overlap_x, overlap_y = (0.93 if tables_per_row > 1 else 1.0), (0.93 if rows > 1 else 1.0)
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            row, col = i // tables_per_row, i % tables_per_row; x, y = int(col * new_width * overlap_x), int(row * new_height * overlap_y)
            if not win32gui.IsWindow(hwnd): continue
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try: win32gui.MoveWindow(hwnd, x, y, new_width, new_height, True); arranged_count += 1
            except Exception as e: self.log(f"Не удалось динамически разместить стол {i+1}: {e}", "error")
        if arranged_count > 0: self.log(f"Динамически расставлено столов: {arranged_count}", "info")

    def arrange_other_windows(self):
        if not self.current_project: self.log(AppConfig.MSG_PROJECT_UNDEFINED, "error"); return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: self.log(f"Нет конфига для {self.current_project}.", "warning"); return
        self.position_player_window(config); self.position_recorder_window()
        if self.current_project == "GG": self.position_lobby_window(config)
        elif self.current_project == "QQ": self.position_cv_server_window(config)
        self.log("Системные окна расставлены.", "info")

    def position_window(self, hwnd, x, y, w, h, log_fail):
        if hwnd and win32gui.IsWindow(hwnd): win32gui.MoveWindow(hwnd, x, y, w, h, True)
        else: self.log(log_fail, "warning")

    def position_player_window(self, config):
        player_config = config.get("PLAYER", {})
        if player_config: self.position_window(self.window_manager.find_first_window_by_title(AppConfig.PLAYER_LAUNCHER_TITLE), player_config["X"], player_config["Y"], player_config["W"], player_config["H"], "Плеер не найден.")

    def position_lobby_window(self, config):
        lobbies = self.window_manager.find_windows_by_config(config, "LOBBY", self.winfo_id())
        cfg = config.get("LOBBY", {})
        if cfg: self.position_window(lobbies[0] if lobbies else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "Лобби не найдено.")

    def position_cv_server_window(self, config):
        cv_windows = self.window_manager.find_windows_by_config(config, "CV_SERVER", self.winfo_id())
        cfg = config.get("CV_SERVER", {})
        if cfg: self.position_window(cv_windows[0] if cv_windows else None, cfg["X"], cfg["Y"], cfg["W"], cfg["H"], "CV Сервер не найден.")

    def position_recorder_window(self):
        recorder_hwnd = self.window_manager.find_first_window_by_process_name(AppConfig.CAMTASIA_PROCESS_NAME)
        if recorder_hwnd and win32gui.IsWindow(recorder_hwnd):
            try:
                if win32gui.IsIconic(recorder_hwnd): win32gui.ShowWindow(recorder_hwnd, win32con.SW_RESTORE); time.sleep(0.2)
                screen_width, screen_height = self.winfo_screenwidth(), self.winfo_screenheight()
                rect = win32gui.GetWindowRect(recorder_hwnd); w, h = rect[2] - rect[0], rect[3] - rect[1]
                x, y = (screen_width - w) // 2, screen_height - h - 40 # Position at bottom center
                win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
            except Exception as e: self.log(f"Ошибка позиционирования Camtasia: {e}", "error")

    def minimize_injector_window(self):
        injector_hwnd = self.window_manager.find_first_window_by_title(AppConfig.INJECTOR_WINDOW_TITLE)
        if injector_hwnd and win32gui.IsWindow(injector_hwnd):
            try: win32gui.ShowWindow(injector_hwnd, win32con.SW_MINIMIZE); self.log("Окно 'injector' свернуто.", "info")
            except Exception as e: self.log(f"Не удалось свернуть окно 'injector': {e}", "error")

    def perform_special_clicks(self):
        if self.current_project != "GG": return
        config = PROJECT_CONFIGS.get("GG")
        if not config: return
        self.log("Выполняю команду SIT-OUT для столов...", "info")
        found_windows = self.window_manager.find_windows_by_config(config, "TABLE", self.winfo_id())
        if not found_windows: self.log("Столы для выполнения команды не найдены.", "warning"); return
        click_count = 0
        for hwnd in found_windows:
            if win32gui.IsWindow(hwnd):
                if self.window_manager.robust_click(hwnd, 25, 410):
                    click_count += 1
                time.sleep(0.1)
        if click_count > 0: self.log(f"Команда SIT-OUT выполнена для {click_count} столов.", "info")

    def position_window_top_right(self):
        self.update_idletasks(); width = self.winfo_width(); margin = AppConfig.WINDOW_MARGIN
        x = self.winfo_screenwidth() - width - margin; self.geometry(f"+{x}+{margin}")

    def position_window_default(self):
        self.update_idletasks(); height = self.winfo_height(); margin = AppConfig.WINDOW_MARGIN
        x = margin; y = self.winfo_screenheight() - height - margin - 40 # 40px for taskbar
        self.geometry(f"+{x}+{y}")


# ===================================================================
# 5. ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ===================================================================
if __name__ == '__main__':
    app = OiHelperApp()
    app.mainloop()
