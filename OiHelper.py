import sys
import win32gui
import re
import requests
import threading
import os
import subprocess
import zipfile # <-- Добавляем модуль для работы с ZIP-архивами
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox
from PyQt6.QtGui import QScreen
from PyQt6.QtCore import QTimer, Qt

# --- ВАЖНО: Настройки для автообновления ---
CURRENT_VERSION = "v1.0" 
# ИЗМЕНЕНО: Указываю ваш репозиторий
GITHUB_REPO = "Vater-v/OiHelper" 
# ИЗМЕНЕНО: Теперь мы ищем ZIP-архив в релизах
ASSET_NAME = "OiHelper.zip" 

# Предполагается, что файл notification.py находится в той же директории
from notification import NotificationManager 

# --- Глобальные переменные ---
manager = None

def Log(text: str, message_type: str = "info"):
    """Вызывает появление уведомления."""
    if manager:
        manager.show(text, message_type)
        QApplication.processEvents()
    else:
        print("Ошибка: Менеджер уведомлений не был создан.")

# --- Класс главного окна ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("OiHelper")
        window_width = 520
        window_height = 330
        self.setFixedSize(window_width, window_height)
        self.position_window_bottom_left()

        self.project_label = None
        self.check_timer = None
        self.update_info = {}

        self.init_ui()
        self.init_project_checker()
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def init_ui(self):
        """Инициализирует и настраивает элементы интерфейса."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.project_label = QLabel("Поиск плеера...")
        self.project_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.project_label.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(self.project_label)

        button1 = QPushButton("Инфо-уведомление")
        button2 = QPushButton("Предупреждение")
        button3 = QPushButton("Ошибка")
        button1.clicked.connect(lambda: Log("Это простое информационное уведомление.", "info"))
        button2.clicked.connect(lambda: Log("Это уведомление-предупреждение.", "warning"))
        button3.clicked.connect(lambda: Log("Это уведомление об ошибке.", "error"))
        layout.addWidget(button1)
        layout.addWidget(button2)
        layout.addWidget(button3)

    def init_project_checker(self):
        self.check_timer = QTimer(self)
        self.check_timer.timeout.connect(self.check_for_player)
        self.check_for_player()

    def check_for_player(self):
        player_found = False
        project_name = None
        project_map = {"QQPoker": "QQ", "ClubGG": "GG"}
        def find_window_callback(hwnd, _):
            nonlocal player_found, project_name
            if player_found: return
            title = win32gui.GetWindowText(hwnd)
            if "holdem" in title.lower():
                player_found = True
                match = re.search(r'\[(.*?)\]', title)
                if match:
                    project_name = project_map.get(match.group(1))
        try:
            win32gui.EnumWindows(find_window_callback, None)
        except Exception as e:
            print(f"Ошибка при перечислении окон: {e}")

        if player_found:
            self.check_timer.stop() 
            if project_name:
                self.project_label.setText(f"{project_name} - Панель управления")
            else:
                self.project_label.setText("Проект не определен")
                # ИЗМЕНЕНО: Возвращаем уведомление
                Log("Нажмите 'Start' на плеере!", "warning")
        else:
            self.project_label.setText("Плеер не найден")
            # ИЗМЕНЕНО: Возвращаем уведомление
            Log("Плеер не запущен!", "warning")
            if not self.check_timer.isActive():
                self.check_timer.start(10000)

    # --- ОБНОВЛЕННАЯ ЛОГИКА: Автообновление ---
    def check_for_updates(self):
        """Проверяет наличие новой версии на GitHub."""
        # ИЗМЕНЕНО: Уведомление о начале проверки
        Log("Проверка обновлений...", "info")
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=5)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release.get("tag_name")

            if latest_version and latest_version > CURRENT_VERSION:
                Log(f"Доступна новая версия: {latest_version}", "info")
                self.update_info = latest_release
                QTimer.singleShot(0, self.prompt_for_update)
            else:
                # ИЗМЕНЕНО: Уведомление, если версия актуальна
                Log("Вы используете последнюю версию.", "info")

        except requests.RequestException as e:
            Log(f"Ошибка проверки обновлений: {e}", "error")
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
            self.apply_update()

    def apply_update(self):
        """Скачивает и применяет обновление."""
        assets = self.update_info.get("assets", [])
        download_url = None
        for asset in assets:
            if asset["name"] == ASSET_NAME:
                download_url = asset["browser_download_url"]
                break

        if not download_url:
            Log("Не удалось найти ZIP-архив в релизе.", "error")
            return

        threading.Thread(target=self.download_and_run_updater, args=(download_url,), daemon=True).start()

    def download_and_run_updater(self, url):
        """Скачивает архив, распаковывает и запускает скрипт-установщик."""
        try:
            Log("Скачивание обновления...", "info")
            update_zip_name = "update.zip"
            response = requests.get(url, stream=True)
            response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            Log("Распаковка архива...", "info")
            update_folder = "update_temp"
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref:
                zip_ref.extractall(update_folder)

            Log("Обновление скачано. Перезапуск...", "info")

            updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            
            with open(updater_script_path, "w", encoding="cp866") as f:
                f.write(f"""
@echo off
chcp 65001 > NUL
echo Ожидание закрытия OiHelper...
timeout /t 3 /nobreak > NUL

echo Удаление старых файлов...
for /d %%i in ("{current_dir}\\*") do (
    if /i not "%%~nxi"=="{update_folder}" (
        rd /s /q "%%i"
    )
)
for %%i in ("{current_dir}\\*.*") do (
    if /i not "%%~nxi"=="{os.path.basename(updater_script_path)}" (
        if /i not "%%~nxi"=="{update_zip_name}" (
            del "%%i"
        )
    )
)

echo Копирование новых файлов...
xcopy /e /y "{current_dir}\\{update_folder}" "{current_dir}"

echo Очистка...
rd /s /q "{current_dir}\\{update_folder}"
del "{current_dir}\\{update_zip_name}"

echo Запуск новой версии...
start "" "{current_exe_path}"

del "{updater_script_path}"
                """)
            
            subprocess.Popen(updater_script_path, shell=True)
            QApplication.instance().quit()

        except Exception as e:
            Log(f"Ошибка при обновлении: {e}", "error")

    def position_window_bottom_left(self):
        screen = QApplication.primaryScreen()
        if screen:
            available_geometry = screen.availableGeometry()
            margin = 65
            frame_h = self.frameGeometry().height()
            x = available_geometry.left() + margin
            y = available_geometry.bottom() - frame_h - margin
            self.move(x, y)

# --- Основной код для запуска ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    manager = NotificationManager()
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
