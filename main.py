import sys
import subprocess
import time
import threading
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QLabel, QListWidget, QCheckBox, QMessageBox
)
import psutil
import win32gui
from update_manager import UpdateManager
from config import AppConfig
import gspread
from google.oauth2.service_account import Credentials
from PyQt6.QtCore import QTimer

# --- Константы из AppConfig
SPREADSHEET_ID = AppConfig.SPREADSHEET_ID
SHEET_NAME = AppConfig.SHEET_NAME
SERVICE_ACCOUNT_FILE = AppConfig.get_service_account_file()
LDPLAYER_PATH = AppConfig.LDPLAYER_PATH


def get_all_ldplayer_emulators():
    """
    Возвращает список всех эмуляторов и их статусы.
    [{ 'index': int, 'name': str, 'status': str }]
    """
    try:
        result = subprocess.run(
            [LDPLAYER_PATH, "list2"],
            capture_output=True, text=True, timeout=10
        )
        emulators = []
        for line in result.stdout.splitlines()[1:]:
            parts = line.split(",")
            if len(parts) >= 5:
                emulators.append({
                    'index': int(parts[0].strip()),
                    'name': parts[1].strip(),
                    'status': parts[4].strip()
                })
        return emulators
    except Exception as e:
        print(f"Ошибка при получении списка эмуляторов: {e}")
        return []

def get_ldplayer_windows():
    windows = []
    def callback(hwnd, _):
        try:
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            pid = None
            try:
                _, pid = win32gui.GetWindowThreadProcessId(hwnd)
            except Exception:
                return
            if not pid:
                return
            try:
                proc = psutil.Process(pid)
                if proc.name().lower() == "dnplayer.exe":
                    windows.append({'pid': pid, 'title': title})
            except Exception:
                pass
        except Exception:
            pass
    win32gui.EnumWindows(callback, None)
    return windows


def is_emulator_window_running(account_key):
    key_l = account_key.strip().lower()
    for win in get_ldplayer_windows():
        if key_l in win['title'].lower():
            return True
    return False

def find_emulator_name_by_id(emulator_id):
    """
    Находит имя эмулятора по уникальному ID из Google Sheets.
    """
    all_names = get_all_ldplayer_emulator_names()
    for name in all_names:
        # name: 'S26-1|31346277|TPuke-4'
        parts = name.split('|')
        if len(parts) >= 2 and parts[1] == str(emulator_id):
            return name
    return None

# --- Окно основного приложения
class OiAutoMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"OiAuto v{AppConfig.CURRENT_VERSION}")
        self.setMinimumSize(480, 360)
        self.accounts = []
        self.automation_active = False
        self._setup_ui()

        # Обновления (оставь как есть)
        self.updater = UpdateManager()
        self.updater.log_request.connect(self.log_update)
        self.updater.status_update.connect(self.status_update)
        self.updater.check_finished.connect(self.update_check_finished)

        self._center_window()
        self.updater.check_for_updates()

        # --- QTimer для автообновления
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_accounts_background)
        self.timer.start(60 * 1000)  # раз в минуту

    def log_update(self, msg, level):
        self.label_status.setText(f"[{level}] {msg}")

    def status_update(self, msg):
        self.label_status.setText(msg)

    def update_check_finished(self):
        # Можно сразу обновить список аккаунтов и т.д.
        self.refresh_accounts()



    def _setup_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        self.label_status = QLabel("Готов к работе")
        self.btn_refresh = QPushButton("Обновить список аккаунтов")
        self.btn_refresh.clicked.connect(self.refresh_accounts)

        self.checkbox_automation = QCheckBox("Автоматика: OFF")
        self.checkbox_automation.setChecked(False)
        self.checkbox_automation.stateChanged.connect(self.toggle_automation)

        self.list_accounts = QListWidget()
        self.list_accounts.addItem("Нет данных")

        layout.addWidget(self.label_status)
        layout.addWidget(self.btn_refresh)
        layout.addWidget(self.list_accounts)
        layout.addWidget(self.checkbox_automation)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        QApplication.setStyle("Fusion")
        self.refresh_accounts()

    def toggle_automation(self, state):
        self.automation_active = bool(state)
        if state:
            self.checkbox_automation.setText("Автоматика: ON")
            self.label_status.setText("Автоматика включена, следим за эмуляторами...")
            self.run_automation_cycle()
        else:
            self.checkbox_automation.setText("Автоматика: OFF")
            self.label_status.setText("Автоматика выключена")
            self.automation_active = False

    def run_automation_cycle(self):
        def cycle():
            while self.automation_active:
                all_emulators = get_all_ldplayer_emulators()
                running_names = [emu['name'] for emu in all_emulators if emu['status'] == 'running']

                remove_running_accounts_from_sheet()  # удаляет уже запущенные из таблицы

                needed = []
                name_map = {}
                real_emulators = get_all_ldplayer_emulator_names()
                needed = []
                for emulator_id in self.accounts:  # Теперь self.accounts — список id из таблицы!
                    real_name = find_emulator_name_by_id(emulator_id)
                    if not real_name:
                        print(f"Нет такого эмулятора с ID: {emulator_id}")
                        self.label_status.setText(f"Нет такого эмулятора с ID: {emulator_id}")
                        QApplication.processEvents()
                        continue
                    if not is_emulator_window_running(real_name):
                        needed.append(real_name)
                if needed:
                    for real_name in needed:
                        self.label_status.setText(f"Запуск: {real_name}")
                        QApplication.processEvents()
                        try:
                            subprocess.Popen([
                                LDPLAYER_PATH, "launch", "--name", real_name
                            ])
                            time.sleep(1)
                        except Exception as e:
                            self.label_status.setText(f"Ошибка запуска: {real_name}: {e}")
                            QApplication.processEvents()
                    self.label_status.setText(
                        f"Запущено: {len(self.accounts) - len(needed)}/{len(self.accounts)} эмуляторов"
                    )
                else:
                    self.label_status.setText("Все нужные эмуляторы уже запущены")
                QApplication.processEvents()
                for _ in range(2):
                    if not self.automation_active:
                        break
                    time.sleep(1)
        threading.Thread(target=cycle, daemon=True).start()

    def _center_window(self):
        frameGm = self.frameGeometry()
        screen = QApplication.primaryScreen().availableGeometry().center()
        frameGm.moveCenter(screen)
        self.move(frameGm.topLeft())

    def update_accounts_background(self):
        threading.Thread(target=self._update_accounts_thread, daemon=True).start()

    def refresh_accounts(self):
        self.label_status.setText("Обновление списка...")
        QApplication.processEvents()
        threading.Thread(target=self._update_accounts_thread, daemon=True).start()

    def _update_accounts_thread(self):
        try:
            short_names = get_accounts_from_sheet_fast()
            name_to_id = get_name_to_id_map()
            self.accounts = []
            self.list_accounts.clear()
            for name in short_names:
                acc_id = name_to_id.get(name)
                if acc_id:
                    self.accounts.append(acc_id)
                    self.list_accounts.addItem(name)
                else:
                    self.list_accounts.addItem(f"{name} [ID не найден!]")
            self.label_status.setText(f"Список аккаунтов обновлён ({len(self.accounts)})")
        except Exception as e:
            self.list_accounts.clear()
            self.list_accounts.addItem("Ошибка чтения таблицы")
            self.label_status.setText(f"Ошибка: {e}")



def get_all_ldplayer_emulator_names():
    try:
        result = subprocess.run(
            [LDPLAYER_PATH, "list2"],
            capture_output=True, text=True, timeout=10
        )
        names = []
        for line in result.stdout.splitlines()[1:]:
            parts = line.split(",")
            if len(parts) >= 2:
                names.append(parts[1].strip())
        return names
    except Exception as e:
        print(f"Ошибка при получении списка эмуляторов: {e}")
        return []

def get_name_to_id_map():
    """
    Возвращает dict: { 'S40-3': '31346277', ... } из листа Accs
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet('Accs')
    data = worksheet.get_all_records()[:130]
    # Для надёжности str()
    return {str(row['Account']).strip(): str(row['ID']).strip() for row in data if row.get('Account') and row.get('ID')}


def remove_running_accounts_from_sheet():
    # 1. Получаем данные из таблицы
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(SHEET_NAME)
    data = worksheet.get_all_records()[:130]
    
    # 3. Проверяем и собираем индексы строк для удаления
    rows_to_delete = []
    for idx, row in enumerate(data, start=2):
        acc = str(row.get('Start', '')).strip()
        if not acc:
            continue
        if is_emulator_window_running(acc):
            rows_to_delete.append(idx)

    # 4. Удаляем строки (с конца вверх, чтобы не сместились индексы)
    for row in reversed(rows_to_delete):
        worksheet.delete_rows(row)

    print(f"Удалено {len(rows_to_delete)} строк. Оставшиеся: {worksheet.row_count}")
    
def get_accounts_from_sheet_fast():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet('OiAuto')
    # Получаем только 130 строк одного столбца 'Start' (A2:A131)
    values = worksheet.get_values('A2:A131')
    accounts = [row[0].strip() for row in values if row and row[0].strip()]
    return accounts




# точка входа
def main():
    app = QApplication(sys.argv)
    window = OiAutoMainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()