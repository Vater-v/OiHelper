import sys
import win32gui
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
# ИЗМЕНЕНО: Версия изменена на 1.1
CURRENT_VERSION = "v1.1" 
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
        # Настройки окна: без рамки, поверх всех окон, не мешает другим окнам
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose) # Автоматическое удаление при закрытии

        # --- Внешний вид ---
        layout = QVBoxLayout(self)
        label = QLabel(message)
        label.setFont(QFont("Arial", 20))
        label.setWordWrap(True)

        background_color = COLORS.get(message_type, COLORS["info"])
        label.setStyleSheet(f"""
            background-color: {background_color};
            color: white;
            padding: 24px;
            border-radius: 10px;
        """)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(label)
        
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)

        # Таймер на автоматическое закрытие через 7 секунд
        QTimer.singleShot(7000, self.hide_animation)

    def show_animation(self):
        """Анимация плавного появления."""
        self.setWindowOpacity(0.0)
        self.show()
        self.animation.setStartValue(0.0)
        self.animation.setEndValue(1.0)
        self.animation.start()

    def hide_animation(self):
        """Анимация плавного исчезновения."""
        self.animation.setStartValue(self.windowOpacity())
        self.animation.setEndValue(0.0)
        self.animation.finished.connect(self.close)
        self.animation.start()

    def closeEvent(self, event):
        """Отправляет сигнал при закрытии окна."""
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
        """Вызывается, когда уведомление закрывается."""
        if notification in self.notifications:
            self.notifications.remove(notification)
        self.reposition_all()

    def reposition_all(self):
        """
        Пересчитывает и устанавливает позицию для всех активных уведомлений.
        """
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

        self.setWindowTitle("OiHelper")
        window_width = 520
        window_height = 330
        self.setFixedSize(window_width, window_height)
        
        self.project_label = None
        self.player_check_timer = None
        self.recorder_check_timer = None
        self.update_info = {}

        self.notification_manager = NotificationManager(self)
        self.log_request.connect(self.log)

        self.init_ui()
        self.init_project_checker()
        self.init_recorder_checker()
        self.position_window_bottom_left()
        
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def log(self, message, message_type):
        """Безопасный метод для отображения уведомлений."""
        if self.notification_manager:
            self.notification_manager.show(message, message_type)
        else:
            print("Критическая ошибка: Менеджер уведомлений не существует.")

    def init_ui(self):
        """Инициализирует и настраивает элементы интерфейса."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 10, 20, 10)
        layout.setSpacing(10)

        # ИЗМЕНЕНО: Метка проекта с увеличенным шрифтом
        self.project_label = QLabel("Поиск плеера...")
        self.project_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.project_label.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        layout.addWidget(self.project_label)
        
        layout.addStretch()

        # Создаем и стилизуем кнопки
        button1 = QPushButton("Инфо-уведомление")
        button2 = QPushButton("Предупреждение")
        button3 = QPushButton("Ошибка")
        
        buttons = [button1, button2, button3]
        
        button_style = """
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
        """
        button1.setStyleSheet(button_style.format(color="#3498DB", hover_color="#5DADE2"))
        button2.setStyleSheet(button_style.format(color="#F1C40F", hover_color="#F4D03F"))
        button3.setStyleSheet(button_style.format(color="#E74C3C", hover_color="#EC7063"))


        for button in buttons:
            button.setMinimumHeight(45) 
            layout.addWidget(button)

        layout.addStretch()

        button1.clicked.connect(lambda: self.log("Это простое информационное уведомление.", "info"))
        button2.clicked.connect(lambda: self.log("Это уведомление-предупреждение.", "warning"))
        button3.clicked.connect(lambda: self.log("Это уведомление об ошибке.", "error"))

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
            except Exception:
                pass

        try:
            win32gui.EnumWindows(find_window_callback, None)
        except Exception as e:
            print(f"Ошибка при перечислении окон: {e}")

        if player_found:
            if self.player_check_timer.isActive():
                self.player_check_timer.stop() 
            if project_name:
                self.project_label.setText(f"{project_name} - Панель управления")
                self.log(f"Проект {project_name} определен.", "info")
            else:
                self.project_label.setText("Проект не определен")
                self.log("Нажмите 'Start' на плеере!", "warning")
        else:
            self.project_label.setText("Плеер не найден")
            if not self.player_check_timer.isActive():
                self.log("Плеер не запущен! Повторная проверка через 10 сек.", "warning")
                self.player_check_timer.start(10000)

    def init_recorder_checker(self):
        """Инициализирует таймер для проверки процесса рекордера."""
        self.recorder_check_timer = QTimer(self)
        self.recorder_check_timer.timeout.connect(self.check_for_recorder)
        self.check_for_recorder() # Первая проверка сразу при запуске

    def check_for_recorder(self):
        """Проверяет, запущен ли рекордер, и пытается запустить его, если нет."""
        try:
            output = subprocess.check_output(['tasklist'], universal_newlines=True, creationflags=0x08000000) # 0x08000000 = CREATE_NO_WINDOW
            if 'recorder' in output.lower():
                if self.recorder_check_timer.isActive():
                    self.recorder_check_timer.stop()
                return 
        except (subprocess.CalledProcessError, FileNotFoundError):
            self.log_request.emit("Не удалось проверить процессы.", "error")
            return

        self.log_request.emit("Рекордер не запущен. Поиск ярлыка...", "warning")

        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        shortcut_found = False
        try:
            for filename in os.listdir(desktop_path):
                if 'recorder' in filename.lower() and filename.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, filename)
                    self.log_request.emit(f"Найден ярлык: {filename}. Запускаю...", "info")
                    try:
                        os.startfile(shortcut_path)
                        self.log_request.emit("Команда на запуск отправлена.", "info")
                        shortcut_found = True
                        if self.recorder_check_timer.isActive():
                            self.recorder_check_timer.stop()
                        break 
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
        """Проверяет наличие новой версии на GitHub."""
        self.log_request.emit("Проверка обновлений...", "info")
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release.get("tag_name")

            if latest_version and latest_version > CURRENT_VERSION:
                self.log_request.emit(f"Доступна новая версия: {latest_version}", "info")
                self.update_info = latest_release
                QTimer.singleShot(0, self.prompt_for_update)
            else:
                self.log_request.emit("Вы используете последнюю версию.", "info")

        except requests.RequestException as e:
            self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error")
        except Exception as e:
            print(f"Неожиданная ошибка при проверке обновлений: {e}")

    def prompt_for_update(self):
        """Спрашивает пользователя, хочет ли он обновиться."""
        if not self.update_info: return
        
        latest_version = self.update_info.get("tag_name")
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setText(f"Доступна новая версия {latest_version}!\nХотите скачать и установить ее?")
        msg_box.setWindowTitle("Доступно обновление")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if msg_box.exec() == QMessageBox.StandardButton.Yes:
            threading.Thread(target=self.apply_update, daemon=True).start()

    def apply_update(self):
        """Скачивает и применяет обновление."""
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
        """Скачивает архив, распаковывает и запускает скрипт-установщик."""
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
start "" "{current_exe_path}"

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
    
    # ИЗМЕНЕНО: Принудительная установка темной темы
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
