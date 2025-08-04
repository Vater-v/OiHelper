import requests, os, sys, logging, threading, zipfile, subprocess
from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtWidgets import QApplication

from config import AppConfig

class UpdateManager(QObject):
    log_request = pyqtSignal(str, str)
    check_finished = pyqtSignal()
    status_update = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.update_info = {}

    def is_new_version_available(self, current_v_str: str, latest_v_str: str) -> bool:
        try:
            current = [int(p) for p in current_v_str.lstrip('v').split('.')]
            latest = [int(p) for p in latest_v_str.lstrip('v').split('.')]
            max_len = max(len(current), len(latest))
            current += [0] * (max_len - len(current))
            latest += [0] * (max_len - len(current))
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
            latest_version = latest_release.get("tag_name")
            if latest_version and self.is_new_version_available(AppConfig.CURRENT_VERSION, latest_version):
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
        download_url = next(
            (asset["browser_download_url"] for asset in self.update_info.get("assets", [])
             if asset["name"] == AppConfig.ASSET_NAME), None)
        if not download_url:
            self.log_request.emit("Не найден архив обновления.", "error")
            return
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url: str):
        update_zip_name = "update.zip"
        try:
            self.status_update.emit("Скачивание обновления...")
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            self.status_update.emit("Распаковка архива...")
            update_folder = "update_temp"
            if os.path.isdir(update_folder):
                import shutil
                shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref:
                zip_ref.extractall(update_folder)

            self.status_update.emit("Перезапуск...")
            updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(
                sys.executable if getattr(sys, 'frozen', False) else sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            exe_name = os.path.basename(current_exe_path)
            script_content = (
                f'@echo off\nchcp 65001 > NUL\necho Waiting for OiAuto to close...\n'
                f'timeout /t 2 /nobreak > NUL\ntaskkill /pid {os.getpid()} /f > NUL\n'
                f'echo Waiting for process to terminate...\ntimeout /t 3 /nobreak > NUL\n'
                f'echo Moving new files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\n'
                f'rd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\n'
                f'del "{current_dir}\\{update_zip_name}"\n'
                f'echo Starting new version...\nstart "" "{exe_name}"\n'
                f'(goto) 2>nul & del "%~f0"'
            )
            with open(updater_script_path, "w", encoding="cp866") as f:
                f.write(script_content)
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()
        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error")
            logging.error("Ошибка при обновлении", exc_info=True)
            if os.path.exists(update_zip_name):
                try:
                    os.remove(update_zip_name)
                except OSError as err:
                    logging.error(f"Не удалось удалить временный файл обновления: {err}")
            self.log_request.emit(AppConfig.MSG_UPDATE_FAIL, "warning")
            self.check_finished.emit()
