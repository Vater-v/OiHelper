import sys
import win32gui
import win32api
import win32con
import win32process # ДОБАВЛЕНО: для работы с процессами
import re
import requests
import threading
import os
import subprocess
import zipfile
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox
from PyQt6.QtCore import QObject, QPropertyAnimation, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QFont, QPalette, QColor

# --- ВАЖНО: Настройки для автообновления ---
CURRENT_VERSION = "v1.5" 
GITHUB_REPO = "Vater-v/OiHelper" 
ASSET_NAME = "OiHelper.zip" 

# ===================================================================
# Начало кода для уведомлений
# ===================================================================

# --- Цвета для разных типов уведомлений ---
COLORS = {
    "info": "#2E86C1",
    "warning": "#F39C12",
    "error": "#E74C3C"
}

# --- Класс самого уведомления (окна) ---
class Notification(QWidget):
    # Сигнал, который отправляется при закрытии уведомления
    closed = pyqtSignal(QWidget)

    def __init__(self, message, message_type):
        super().__init__()
        self.is_closing = False

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose) 

        layout = QVBoxLayout(self)
        label = QLabel(message)
        label.setFont(QFont("Arial", 22))
        label.setWordWrap(True)

        background_color = COLORS.get(message_type, COLORS["info"])
        label.setStyleSheet(f"""
            background-color: {background_color};
            color: white;
            padding: 28px;
            border-radius: 10px;
        """)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(label)
        
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)

        QTimer.singleShot(7000, self.hide_animation)

    def show_animation(self):
        self.setWindowOpacity(0.0)
        self.show()
        self.animation.setStartValue(0.0)
        self.animation.setEndValue(1.0)
        self.animation.start()

    def hide_animation(self):
        if self.is_closing:
            return
        self.is_closing = True
        
        self.animation.setStartValue(self.windowOpacity())
        self.animation.setEndValue(0.0)
        self.animation.finished.connect(self.close)
        self.animation.start()

    def closeEvent(self, event):
        self.closed.emit(self)
        super().closeEvent(event)

# --- Класс-менеджер для управления уведомлениями ---
class NotificationManager(QObject):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.notifications = []

    def show(self, message, message_type):
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
            width = n.width()
            height = n.height()
            
            x = screen_geo.right() - width - margin
            y = screen_geo.bottom() - height - margin - total_height
            
            n.move(x, y)
            total_height += height + 10

# ===================================================================
# Конец кода для уведомлений
# ===================================================================


# --- Класс главного окна ---
class MainWindow(QMainWindow):
    log_request = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()

        self.setWindowTitle(f"OiHelper {CURRENT_VERSION}")
        window_width = 470
        window_height = 330
        self.setFixedSize(window_width, window_height)
        
        self.project_label = None
        self.player_check_timer = None
        self.recorder_check_timer = None
        self.update_info = {}
        
        self.is_recording = False
        self.recording_status_timer = QTimer(self)
        self.recording_status_timer.timeout.connect(self.check_recording_status)
        self.toggle_record_button = None

        self.notification_manager = NotificationManager(self)
        self.log_request.connect(self.log)

        self.init_ui()
        self.init_project_checker()
        self.init_recorder_checker()
        self.position_window_bottom_left()
        
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def log(self, message, message_type):
        if self.notification_manager:
            self.notification_manager.show(message, message_type)
        else:
            print("Критическая ошибка: Менеджер уведомлений не существует.")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 10, 20, 10)
        layout.setSpacing(10)

        self.project_label = QLabel("Поиск плеера...")
        self.project_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.project_label.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        layout.addWidget(self.project_label)
        
        layout.addStretch()

        self.toggle_record_button = QPushButton("Начать запись")
        self.toggle_record_button.clicked.connect(self.toggle_recording)
        
        arrange_tables_button = QPushButton("Расставить столы")
        arrange_tables_button.clicked.connect(self.arrange_tables)

        arrange_other_button = QPushButton("Расставить остальное")
        arrange_other_button.clicked.connect(self.arrange_other_windows)
        
        buttons = [self.toggle_record_button, arrange_tables_button, arrange_other_button]
        
        self.button_style_sheet = """
            QPushButton {{
                font-size: 15px;
                font-weight: bold;
                color: white;
                background-color: {color};
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
            QPushButton:disabled {{
                background-color: #5D6D7E;
                color: #BDC3C7;
            }}
        """
        self.update_record_button_style()
        arrange_tables_button.setStyleSheet(self.button_style_sheet.format(color="#3498DB", hover_color="#5DADE2"))
        arrange_other_button.setStyleSheet(self.button_style_sheet.format(color="#9B59B6", hover_color="#AF7AC5"))


        for button in buttons:
            button.setMinimumHeight(45) 
            layout.addWidget(button)

        layout.addStretch()

    def update_record_button_style(self):
        if self.is_recording:
            self.toggle_record_button.setText("Закончить запись")
            self.toggle_record_button.setStyleSheet(self.button_style_sheet.format(color="#E74C3C", hover_color="#EC7063")) 
        else:
            self.toggle_record_button.setText("Начать запись")
            self.toggle_record_button.setStyleSheet(self.button_style_sheet.format(color="#27AE60", hover_color="#2ECC71")) 

    def press_key(self, key_code):
        try:
            win32api.keybd_event(key_code, 0, 0, 0)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e:
            self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

    def toggle_recording(self):
        self.is_recording = not self.is_recording
        self.update_record_button_style()

        if self.is_recording:
            self.log_request.emit("Начинаю запись...", "info")
            self.press_key(win32con.VK_F9)
            self.recording_status_timer.start(5000) 
            QTimer.singleShot(2000, self.check_recording_status)
        else:
            self.log_request.emit("Останавливаю запись...", "info")
            self.press_key(win32con.VK_F10)
            self.recording_status_timer.stop()
        
        QTimer.singleShot(500, self.position_recorder_window)


    def find_window_by_title(self, text_in_title, exact_match=False):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if (exact_match and text_in_title == title) or \
                       (not exact_match and text_in_title.lower() in title.lower()):
                        hwnds.append(hwnd)
                except Exception:
                    pass 
        win32gui.EnumWindows(callback, None)
        return hwnds[0] if hwnds else None

    def check_recording_status(self):
        if not self.is_recording:
            if self.recording_status_timer.isActive():
                self.recording_status_timer.stop()
            return

        is_recording_active = self.find_window_by_title("Recording...")
        is_paused = self.find_window_by_title("Paused...")

        if is_recording_active:
            return
        
        if is_paused:
            self.log_request.emit("Запись на паузе. Возобновляю...", "warning")
            self.press_key(win32con.VK_F9)
            return

        if self.is_recording: 
            self.log_request.emit("Запись прервана или не началась!", "error")
            self.is_recording = False
            self.update_record_button_style()
            if self.recording_status_timer.isActive():
                self.recording_status_timer.stop()
    
    def arrange_tables(self):
        # ИЗМЕНЕНО: Упрощенное логирование
        self.log_request.emit("Ищу столы...", "info")

        TARGET_W, TARGET_H = 557, 424
        TARGET_ASPECT_RATIO = TARGET_W / TARGET_H
        TOLERANCE = 0.05
        MIN_RATIO = TARGET_ASPECT_RATIO * (1 - TOLERANCE)
        MAX_RATIO = TARGET_ASPECT_RATIO * (1 + TOLERANCE)
        # ИЗМЕНЕНО: Список исключений
        EXCLUDED_TITLES = ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"]

        SLOTS = [(-5, 0), (276, 420), (826, 0), (1101, 425)]
        found_windows = []

        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd) or hwnd == self.winId():
                return True
            try:
                title = win32gui.GetWindowText(hwnd)
                # ИЗМЕНЕНО: Проверка на исключения
                if any(excluded.lower() in title.lower() for excluded in EXCLUDED_TITLES):
                    return True

                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w > 0 and h > 0:
                    current_aspect_ratio = w / h
                    if MIN_RATIO <= current_aspect_ratio <= MAX_RATIO:
                        found_windows.append(hwnd)
            except Exception: pass
            return True

        win32gui.EnumWindows(enum_windows_callback, None)

        if not found_windows:
            self.log_request.emit("Столы не найдены.", "warning")
            return

        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS): break
            x, y = SLOTS[i]
            win32gui.MoveWindow(hwnd, x, y, TARGET_W, TARGET_H, True)
            arranged_count += 1
        
        self.log_request.emit(f"Расставил столов: {arranged_count}", "info")

    def arrange_other_windows(self):
        self.log_request.emit("Расставляю остальные окна...", "info")
        self.position_player_window()
        self.position_ratio_window()
        self.position_recorder_window()
        self.log_request.emit("Расстановка завершена.", "info")

    def position_player_window(self):
        self.log_request.emit("Ставлю плеер на место...", "info")
        player_hwnd = self.find_window_by_title("holdem")
        if player_hwnd:
            win32gui.MoveWindow(player_hwnd, 1385, 0, 700, 365, True)
            self.log_request.emit("Плеер на месте.", "info")
        else:
            self.log_request.emit("Не нашел плеер.", "error")

    def position_ratio_window(self):
        self.log_request.emit("Ищу дополнительное окно...", "info")
        TARGET_W, TARGET_H = 333, 623
        TARGET_ASPECT_RATIO = TARGET_W / TARGET_H
        TOLERANCE = 0.05
        MIN_RATIO = TARGET_ASPECT_RATIO * (1 - TOLERANCE)
        MAX_RATIO = TARGET_ASPECT_RATIO * (1 + TOLERANCE)
        EXCLUDED_TITLES = ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"]
        
        found_hwnd = None
        
        def enum_windows_callback(hwnd, _):
            nonlocal found_hwnd
            if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd) or hwnd == self.winId(): return True
            try:
                title = win32gui.GetWindowText(hwnd)
                if any(excluded.lower() in title.lower() for excluded in EXCLUDED_TITLES):
                    return True

                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                if w > 0 and h > 0:
                    ratio = w / h
                    if MIN_RATIO <= ratio <= MAX_RATIO:
                        found_hwnd = hwnd
                        return False 
            except Exception: pass
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)

        if found_hwnd:
            win32gui.MoveWindow(found_hwnd, 1640, 140, TARGET_W, TARGET_H, True)
            self.log_request.emit("Дополнительное окно на месте.", "info")
        else:
            self.log_request.emit("Не нашел дополнительное окно.", "error")
            
    def position_recorder_window(self):
        self.log_request.emit("Ставлю рекордер на место...", "info")
        recorder_hwnd = self.find_window_by_title("recorder")
        if not recorder_hwnd:
            self.log_request.emit("Не нашел рекордер.", "error")
            return
            
        try:
            screen_rect = QApplication.primaryScreen().availableGeometry()
            rect = win32gui.GetWindowRect(recorder_hwnd)
            w, h = rect[2] - rect[0], rect[3] - rect[1]
            
            x = screen_rect.left() + (screen_rect.width() - w) // 2
            y = screen_rect.bottom() - h - 20 
            
            win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
            self.log_request.emit("Рекордер на месте.", "info")
        except Exception as e:
            self.log_request.emit(f"Ошибка позиционирования рекордера: {e}", "error")


    def init_project_checker(self):
        self.player_check_timer = QTimer(self)
        self.player_check_timer.timeout.connect(self.check_for_player)
        self.check_for_player()

    def check_for_player(self):
        player_found = False
        project_name = None
        project_map = {"QQPoker": "QQ", "ClubGG": "GG"}
        
        def find_window_callback(hwnd, _):
            nonlocal player_found, project_name
            if player_found: return
            try:
                title = win32gui.GetWindowText(hwnd)
                if "holdem" in title.lower():
                    player_found = True
                    for full_name, short_name in project_map.items():
                        if f"[{full_name}]" in title:
                            project_name = short_name
                            break
            except Exception: pass

        try:
            win32gui.EnumWindows(find_window_callback, None)
        except Exception as e:
            print(f"Ошибка при перечислении окон: {e}")

        if player_found:
            if self.player_check_timer.isActive(): self.player_check_timer.stop() 
            if project_name: self.project_label.setText(f"{project_name} - Панель управления")
            else:
                self.project_label.setText("Проект не определен")
                self.log("Нажмите 'Start' на плеере!", "warning")
        else:
            self.project_label.setText("Плеер не найден")
            if not self.player_check_timer.isActive():
                self.log("Плеер не запущен! Повторная проверка через 10 сек.", "warning")
                self.player_check_timer.start(10000)

    def is_recorder_process_running(self):
        try:
            output = subprocess.check_output(['tasklist'], universal_newlines=True, creationflags=0x08000000)
            return 'recorder' in output.lower()
        except (subprocess.CalledProcessError, FileNotFoundError):
            self.log_request.emit("Не удалось проверить процессы.", "error")
            return False

    def init_recorder_checker(self):
        self.recorder_check_timer = QTimer(self)
        self.recorder_check_timer.timeout.connect(self.check_for_recorder)
        self.check_for_recorder() 

    def check_for_recorder(self):
        if self.is_recorder_process_running():
            self.toggle_record_button.setEnabled(True)
            if self.recorder_check_timer.isActive(): self.recorder_check_timer.stop()
            return
        
        self.toggle_record_button.setEnabled(False)

        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        shortcut_found = False
        try:
            for filename in os.listdir(desktop_path):
                if 'recorder' in filename.lower() and filename.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, filename)
                    self.log_request.emit(f"Рекордер не запущен. Запускаю {filename}...", "warning")
                    try:
                        os.startfile(shortcut_path)
                        shortcut_found = True
                        if self.recorder_check_timer.isActive(): self.recorder_check_timer.stop()
                        QTimer.singleShot(3000, self.check_for_recorder)
                        return
                    except Exception as e:
                        self.log_request.emit(f"Не удалось запустить ярлык: {e}", "error")
                        break
            
            if not shortcut_found:
                self.log_request.emit("Ярлык рекордера на рабочем столе не найден.", "error")

        except Exception as e:
            self.log_request.emit(f"Ошибка при поиске на рабочем столе: {e}", "error")

        if not self.recorder_check_timer.isActive():
            self.recorder_check_timer.start(10000)


    def check_for_updates(self):
        self.log_request.emit("Проверка обновлений...", "info")
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release.get("tag_name")

            if latest_version and latest_version > CURRENT_VERSION:
                self.log_request.emit(f"Доступна новая версия: {latest_version}. Автоматическое обновление...", "info")
                self.update_info = latest_release
                threading.Thread(target=self.apply_update, daemon=True).start()
            else:
                self.log_request.emit("Вы используете последнюю версию.", "info")

        except requests.RequestException as e:
            self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error")
        except Exception as e:
            print(f"Неожиданная ошибка при проверке обновлений: {e}")

    def apply_update(self):
        assets = self.update_info.get("assets", [])
        download_url = None
        for asset in assets:
            if asset["name"] == ASSET_NAME:
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
            
            with open(updater_script_path, "w", encoding="cp866") as f:
                f.write(f"""
@echo off
chcp 65001 > NUL
echo Waiting for OiHelper to close...
taskkill /pid {os.getpid()} /f > NUL
timeout /t 2 /nobreak > NUL

echo Removing old files...
robocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL
rd /s /q "{current_dir}\\{update_folder}"

echo Cleaning up...
del "{current_dir}\\{update_zip_name}"

echo Starting new version...
start "" "{exe_name}"

(goto) 2>nul & del "%~f0"
                """)
            
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()

        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error")
            print(f"Update error: {e}")

    def position_window_bottom_left(self):
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry()
                margin = 65
                x = available_geometry.left() + margin
                y = available_geometry.bottom() - self.frameGeometry().height() - margin
                self.move(x, y)
        except Exception as e:
            print(f"Could not position window: {e}")


# --- Основной код для запуска ---
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
