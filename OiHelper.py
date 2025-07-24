import sys
import win32gui
import win32api
import win32con
import win32process 
import re
import requests
import threading
import os
import subprocess
import zipfile
import time
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox, QProgressBar
from PyQt6.QtCore import QObject, QPropertyAnimation, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QFont, QPalette, QColor, QIcon

# --- ВАЖНО: Настройки для автообновления ---
CURRENT_VERSION = "v1.8" 
GITHUB_REPO = "Vater-v/OiHelper" 
ASSET_NAME = "OiHelper.zip" 

# ===================================================================
# Конфигурация для разных проектов
# ===================================================================
PROJECT_CONFIGS = {
    "GG": {
        "TABLE": { "W": 557, "H": 424, "TOLERANCE": 0.05 },
        "LOBBY": { "W": 333, "H": 623, "TOLERANCE": 0.05, "X": 1600, "Y": 140 },
        "PLAYER": { "W": 700, "H": 365, "X": 1385, "Y": 0 },
        "TABLE_SLOTS": [(-5, 0), (271, 420), (816, 0), (1086, 425)],
        "EXCLUDED_TITLES": ["OiHelper", "NekoRay", "NekoBox", "Chrome", "Sandbo", "Notepad", "Explorer"],
        "SESSION_MAX_DURATION_S": 4 * 3600, # 4 часа в секундах
        "SESSION_WARN_TIME_S": 3.5 * 3600   # 3.5 часа в секундах
    },
    "QQ": {
        # Конфигурация для QQ будет добавлена позже
    }
}


# ===================================================================
# Начало кода для уведомлений
# ===================================================================

# ИЗМЕНЕНО: Добавлены цвета для градиента в уведомлениях
COLORS = { 
    "info": ("#3498DB", "#2E86C1"), 
    "warning": ("#F39C12", "#D35400"), 
    "error": ("#E74C3C", "#C0392B") 
}

# --- Класс самого уведомления (окна) ---
class Notification(QWidget):
    closed = pyqtSignal(QWidget)
    def __init__(self, message, message_type):
        super().__init__()
        self.is_closing = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose) 
        layout = QVBoxLayout(self)
        label = QLabel(message)
        # ИЗМЕНЕНО: Шрифт уведомлений стал крупнее
        label.setFont(QFont("Arial", 21))
        label.setWordWrap(True)
        
        # ИЗМЕНЕНО: Используется градиент для улучшения качества графики
        stop1, stop2 = COLORS.get(message_type, COLORS["info"])
        label.setStyleSheet(f"""
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {stop1}, stop:1 {stop2}); 
            color: white; 
            padding: 28px; 
            border-radius: 12px;
            border: 1px solid #555;
        """)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(label)
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)
        QTimer.singleShot(7000, self.hide_animation)
    def show_animation(self):
        self.setWindowOpacity(0.0)
        self.show()
        self.animation.setStartValue(0.0); self.animation.setEndValue(1.0); self.animation.start()
    def hide_animation(self):
        if self.is_closing: return
        self.is_closing = True
        self.animation.setStartValue(self.windowOpacity()); self.animation.setEndValue(0.0)
        self.animation.finished.connect(self.close); self.animation.start()
    def closeEvent(self, event):
        self.closed.emit(self); super().closeEvent(event)

# --- Класс-менеджер для управления уведомлениями ---
class NotificationManager(QObject):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window; self.notifications = []
    def show(self, message, message_type):
        if len(self.notifications) >= 5: self.notifications[0].hide_animation()
        notification = Notification(message, message_type)
        notification.closed.connect(self.on_notification_closed)
        self.notifications.append(notification); self.reposition_all(); notification.show_animation()
    def on_notification_closed(self, notification):
        if notification in self.notifications: self.notifications.remove(notification)
        self.reposition_all()
    def reposition_all(self):
        try: screen_geo = QApplication.primaryScreen().availableGeometry()
        except AttributeError: screen_geo = QApplication.screens()[0].availableGeometry()
        margin = 20; total_height = 0
        for n in reversed(self.notifications):
            n.adjustSize(); width = n.width(); height = n.height()
            x = screen_geo.right() - width - margin
            y = screen_geo.bottom() - height - margin - total_height
            n.move(x, y); total_height += height + 10

# ===================================================================
# Конец кода для уведомлений
# ===================================================================


# --- Класс главного окна ---
class MainWindow(QMainWindow):
    log_request = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"OiHelper {CURRENT_VERSION}")
        self.setFixedSize(470, 380) # Высота немного увеличена
        
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        
        if os.path.exists('icon.ico'):
            self.setWindowIcon(QIcon('icon.ico'))
        
        self.project_label = None; self.player_check_timer = None; self.recorder_check_timer = None
        self.update_info = {}; self.current_project = None

        self.is_auto_record_enabled = True
        self.auto_record_timer = QTimer(self)
        self.auto_record_timer.timeout.connect(self.check_auto_record_logic)
        self.auto_record_toggle_button = None
        
        self.auto_arrange_timer = QTimer(self)
        self.auto_arrange_timer.timeout.connect(self.check_for_new_tables)
        self.last_table_count = 0

        self.session_timer = QTimer(self)
        self.session_timer.timeout.connect(self.update_session_progress)
        self.recording_start_time = 0
        
        self.flash_timer = QTimer(self)
        self.flash_timer.timeout.connect(self.toggle_window_flash)
        self.is_flashing = False
        self.flash_state = False

        self.notification_manager = NotificationManager(self)
        self.log_request.connect(self.log)

        self.init_ui()
        self.init_project_checker()
        self.init_recorder_checker()
        self.position_window_bottom_left()
        
        self.auto_record_timer.start(3000) 
        QTimer.singleShot(1000, self.initial_startup_check)
        
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def log(self, message, message_type):
        if self.notification_manager: self.notification_manager.show(message, message_type)
        else: print(f"Критическая ошибка: {message}")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 15, 20, 15); layout.setSpacing(12)

        self.project_label = QLabel("Поиск плеера...")
        self.project_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # ИЗМЕНЕНО: Заголовок стал еще крупнее
        self.project_label.setStyleSheet("font-size: 28px; font-weight: bold; color: white;")
        layout.addWidget(self.project_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar { border: 2px solid #555; border-radius: 5px; text-align: center; height: 12px; }
            QProgressBar::chunk { background-color: #05B8CC; border-radius: 5px;}
        """)
        layout.addWidget(self.progress_bar)

        self.progress_bar_label = QLabel("Следующий перезапуск через: 04:00:00")
        self.progress_bar_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar_label.setStyleSheet("font-size: 14px; color: #aaa;")
        self.progress_bar_label.setVisible(False)
        layout.addWidget(self.progress_bar_label)

        self.stop_flash_button = QPushButton("✋ Остановить мигание")
        self.stop_flash_button.setVisible(False)
        self.stop_flash_button.clicked.connect(self.stop_flashing)
        layout.addWidget(self.stop_flash_button)
        
        layout.addSpacing(10)
        layout.addStretch()

        self.auto_record_toggle_button = QPushButton("▶️ Автозапись: ВКЛ")
        self.auto_record_toggle_button.clicked.connect(self.toggle_auto_record)
        
        arrange_tables_button = QPushButton("🃏 Расставить столы")
        arrange_tables_button.clicked.connect(self.arrange_tables)

        arrange_other_button = QPushButton("✨ Расставить остальное")
        arrange_other_button.clicked.connect(self.arrange_other_windows)
        
        buttons = [self.auto_record_toggle_button, arrange_tables_button, arrange_other_button]
        
        # ИЗМЕНЕНО: Увеличен шрифт кнопок и улучшен дизайн
        self.button_style_sheet = """
            QPushButton {{ 
                font-size: 17px; 
                font-weight: bold; 
                color: white; 
                background-color: {color}; 
                border-radius: 8px; 
                border: 1px solid #222;
            }}
            QPushButton:hover {{ background-color: {hover_color}; }}
            QPushButton:disabled {{ background-color: #5D6D7E; color: #BDC3C7; border: 1px solid #444;}}
        """
        self.update_auto_record_button_style()
        arrange_tables_button.setStyleSheet(self.button_style_sheet.format(color="#3498DB", hover_color="#5DADE2"))
        arrange_other_button.setStyleSheet(self.button_style_sheet.format(color="#9B59B6", hover_color="#AF7AC5"))
        self.stop_flash_button.setStyleSheet(self.button_style_sheet.format(color="#F39C12", hover_color="#F5B041"))

        for button in buttons:
            button.setMinimumHeight(50); layout.addWidget(button)
        layout.addStretch()

    def update_auto_record_button_style(self):
        if self.is_auto_record_enabled:
            self.auto_record_toggle_button.setText("▶️ Автозапись: ВКЛ")
            self.auto_record_toggle_button.setStyleSheet(self.button_style_sheet.format(color="#27AE60", hover_color="#2ECC71")) 
        else:
            self.auto_record_toggle_button.setText("⏹️ Автозапись: ВЫКЛ")
            self.auto_record_toggle_button.setStyleSheet(self.button_style_sheet.format(color="#E74C3C", hover_color="#EC7063")) 

    def press_key(self, key_code):
        try:
            win32api.keybd_event(key_code, 0, 0, 0)
            win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            QTimer.singleShot(500, self.position_recorder_window)
        except Exception as e:
            self.log_request.emit(f"Ошибка эмуляции нажатия: {e}", "error")

    def toggle_auto_record(self):
        self.is_auto_record_enabled = not self.is_auto_record_enabled
        self.update_auto_record_button_style()
        if self.is_auto_record_enabled:
            self.log_request.emit("Автозапись включена.", "info")
            self.auto_record_timer.start(3000)
            self.check_auto_record_logic()
        else:
            self.log_request.emit("Автозапись выключена. Ручной режим.", "warning")
            self.auto_record_timer.stop()
            self.progress_bar.setVisible(False)
            self.progress_bar_label.setVisible(False)
            if self.is_flashing: self.stop_flashing()
    
    def check_auto_record_logic(self):
        if not self.is_auto_record_enabled or not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project);
        if not config: return

        try:
            tables_exist = len(self._find_windows_by_ratio(config, "TABLE")) > 0
            lobby_exists = self._find_windows_by_ratio(config, "LOBBY")
            is_recording = self.find_window_by_title("Recording...") is not None
            is_paused = self.find_window_by_title("Paused...") is not None

            if tables_exist or lobby_exists:
                if not is_recording and not is_paused:
                    self.log_request.emit("Обнаружены столы/лобби. Начинаю автозапись...", "info")
                    self.start_recording_session()
                elif is_paused:
                    self.log_request.emit("Запись на паузе. Возобновляю...", "warning")
                    self.press_key(win32con.VK_F9)
            else:
                if is_recording or is_paused:
                    self.log_request.emit("Столы и лобби закрыты. Останавливаю запись...", "info")
                    self.stop_recording_session()
        except Exception as e:
            self.log_request.emit(f"Ошибка в логике автозаписи: {e}", "error")

    def start_recording_session(self):
        self.press_key(win32con.VK_F9)
        if self.current_project == "GG":
            self.recording_start_time = time.time()
            config = PROJECT_CONFIGS["GG"]
            self.progress_bar.setMaximum(config["SESSION_MAX_DURATION_S"])
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_bar_label.setVisible(True)
            self.session_timer.start(1000)

    def stop_recording_session(self):
        self.press_key(win32con.VK_F10)
        if self.session_timer.isActive():
            self.session_timer.stop()
            self.recording_start_time = 0
            self.progress_bar.setVisible(False)
            self.progress_bar_label.setVisible(False)
            self.stop_flashing()

    def update_session_progress(self):
        if self.recording_start_time == 0: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "SESSION_MAX_DURATION_S" not in config: return

        elapsed = time.time() - self.recording_start_time
        self.progress_bar.setValue(int(elapsed))
        
        remaining_s = config["SESSION_MAX_DURATION_S"] - elapsed
        if remaining_s < 0: remaining_s = 0
        formatted_time = time.strftime('%H:%M:%S', time.gmtime(remaining_s))
        self.progress_bar_label.setText(f"Следующий перезапуск через: {formatted_time}")
        
        progress_percent = elapsed / config["SESSION_MAX_DURATION_S"]
        if progress_percent < 0.5: color = "#2ECC71"
        elif progress_percent < 0.85: color = "#F1C40F"
        else: color = "#E74C3C"
        self.progress_bar.setStyleSheet(f"QProgressBar::chunk {{ background-color: {color}; border-radius: 5px;}} QProgressBar {{ border: 2px solid #555; border-radius: 5px; text-align: center; height: 12px;}}")

        if elapsed >= config["SESSION_WARN_TIME_S"] and not self.is_flashing:
            self.start_flashing()
        
        if elapsed >= config["SESSION_MAX_DURATION_S"]:
            self.log_request.emit("4 часа записи истекли. Перезапускаю сессию...", "info")
            self.stop_recording_session()
            QTimer.singleShot(2000, self.check_auto_record_logic)
    
    def start_flashing(self):
        if self.is_flashing: return
        self.is_flashing = True
        self.stop_flash_button.setVisible(True)
        self.flash_timer.start(500)

    def stop_flashing(self):
        if not self.is_flashing: return
        self.is_flashing = False
        self.stop_flash_button.setVisible(False)
        self.flash_timer.stop()
        self.setStyleSheet("")

    def toggle_window_flash(self):
        if self.flash_state:
            self.setStyleSheet("QMainWindow { border: 3px solid red; }")
        else:
            self.setStyleSheet("")
        self.flash_state = not self.flash_state

    def initial_startup_check(self):
        if self.find_window_by_title("Recording..."):
            self.log_request.emit("Обнаружена активная запись. Перезапускаю для синхронизации...", "warning")
            self.stop_recording_session()
            QTimer.singleShot(2000, self.check_auto_record_logic)

    def find_window_by_process_name(self, process_name_to_find):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    process_name = win32process.GetModuleFileNameEx(h_process, 0)
                    win32api.CloseHandle(h_process)
                    if process_name_to_find.lower() in os.path.basename(process_name).lower():
                        hwnds.append(hwnd)
                except Exception: pass
        win32gui.EnumWindows(callback, None)
        return hwnds[0] if hwnds else None

    def find_window_by_title(self, text_in_title, exact_match=False):
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if (exact_match and text_in_title == title) or (not exact_match and text_in_title.lower() in title.lower()):
                        hwnds.append(hwnd)
                except Exception: pass 
        win32gui.EnumWindows(callback, None)
        return hwnds[0] if hwnds else None

    def _find_windows_by_ratio(self, config, config_key):
        window_config = config.get(config_key, {});
        if not window_config: return []
        TARGET_ASPECT_RATIO = window_config["W"] / window_config["H"]
        TOLERANCE = window_config["TOLERANCE"]
        MIN_RATIO = TARGET_ASPECT_RATIO * (1 - TOLERANCE); MAX_RATIO = TARGET_ASPECT_RATIO * (1 + TOLERANCE)
        EXCLUDED_TITLES = config.get("EXCLUDED_TITLES", [])
        found_windows = []
        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd) or hwnd == self.winId(): return True
            try:
                title = win32gui.GetWindowText(hwnd)
                if any(excluded.lower() in title.lower() for excluded in EXCLUDED_TITLES): return True
                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                if w > 0 and h > 0 and MIN_RATIO <= (w / h) <= MAX_RATIO:
                    found_windows.append(hwnd)
            except Exception: pass
            return True
        win32gui.EnumWindows(enum_windows_callback, None)
        return found_windows

    def check_for_new_tables(self):
        if not self.current_project: return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: return
        
        current_tables = self._find_windows_by_ratio(config, "TABLE")
        current_count = len(current_tables)

        if current_count > self.last_table_count:
            self.log_request.emit("Обнаружен новый стол. Расставляю...", "info")
            self.arrange_tables()
        
        self.last_table_count = current_count

    def arrange_tables(self):
        if not self.current_project: self.log_request.emit("Проект не определен. Не могу расставить столы.", "error"); return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config or "TABLE" not in config: self.log_request.emit(f"Нет конфигурации столов для проекта {self.current_project}.", "warning"); return
        self.log_request.emit("Ищу столы...", "info")
        found_windows = self._find_windows_by_ratio(config, "TABLE")
        if not found_windows: self.log_request.emit("Столы не найдены.", "warning"); return
        SLOTS = config["TABLE_SLOTS"]; TARGET_W, TARGET_H = config["TABLE"]["W"], config["TABLE"]["H"]
        arranged_count = 0
        for i, hwnd in enumerate(found_windows):
            if i >= len(SLOTS): break
            x, y = SLOTS[i]
            win32gui.MoveWindow(hwnd, x, y, TARGET_W, TARGET_H, True); arranged_count += 1
        self.log_request.emit(f"Расставил {arranged_count} столов.", "info")

    def arrange_other_windows(self):
        if not self.current_project: self.log_request.emit("Проект не определен. Не могу расставить окна.", "error"); return
        config = PROJECT_CONFIGS.get(self.current_project)
        if not config: self.log_request.emit(f"Нет конфигурации для проекта {self.current_project}.", "warning"); return
        self.log_request.emit("Расставляю остальные окна...", "info")
        self.position_player_window(config); self.position_lobby_window(config); self.position_recorder_window()
        self.log_request.emit("Расстановка завершена.", "info")

    def position_player_window(self, config):
        player_config = config.get("PLAYER", {});
        if not player_config: return
        player_hwnd = self.find_window_by_title("holdem")
        if player_hwnd:
            win32gui.MoveWindow(player_hwnd, player_config["X"], player_config["Y"], player_config["W"], player_config["H"], True)
            self.log_request.emit("Плеер на месте.", "info")
        else: self.log_request.emit("Не нашел плеер.", "error")

    def position_lobby_window(self, config):
        lobby_config = config.get("LOBBY", {});
        if not lobby_config: return
        lobbies = self._find_windows_by_ratio(config, "LOBBY"); lobby_hwnd = lobbies[0] if lobbies else None
        if lobby_hwnd:
            win32gui.MoveWindow(lobby_hwnd, lobby_config["X"], lobby_config["Y"], lobby_config["W"], lobby_config["H"], True)
            self.log_request.emit("Лобби на месте.", "info")
        else: self.log_request.emit("Не нашел Лобби.", "error")
            
    def position_recorder_window(self):
        recorder_hwnd = self.find_window_by_process_name("recorder");
        if not recorder_hwnd: return
        try:
            screen_rect = QApplication.primaryScreen().availableGeometry()
            rect = win32gui.GetWindowRect(recorder_hwnd); w, h = rect[2] - rect[0], rect[3] - rect[1]
            x = screen_rect.left() + (screen_rect.width() - w) // 2
            y = screen_rect.bottom() - h - 20 
            win32gui.MoveWindow(recorder_hwnd, x, y, w, h, True)
        except Exception as e:
            self.log_request.emit(f"Ошибка позиционирования Camtasia: {e}", "error")

    def init_project_checker(self):
        self.player_check_timer = QTimer(self); self.player_check_timer.timeout.connect(self.check_for_player)
        self.check_for_player()

    def check_for_player(self):
        player_found = False; project_name = None
        project_map = {"QQPoker": "QQ", "ClubGG": "GG"}
        def find_window_callback(hwnd, _):
            nonlocal player_found, project_name
            if player_found: return
            try:
                title = win32gui.GetWindowText(hwnd)
                if "holdem" in title.lower():
                    player_found = True
                    for full_name, short_name in project_map.items():
                        if f"[{full_name}]" in title: project_name = short_name; break
            except Exception: pass
        try: win32gui.EnumWindows(find_window_callback, None)
        except Exception as e: print(f"Ошибка при перечислении окон: {e}")

        if player_found:
            if self.player_check_timer.isActive(): self.player_check_timer.stop() 
            if project_name != self.current_project:
                self.current_project = project_name 
                self.last_table_count = 0 
                if project_name: 
                    self.project_label.setText(f"{project_name} - Панель управления")
                    self.auto_arrange_timer.start(2500) 
                else:
                    self.project_label.setText("Проект не определен")
                    self.log("Нажмите 'Start' на плеере!", "warning")
        else:
            if self.current_project is not None:
                self.current_project = None 
                self.last_table_count = 0
                self.auto_arrange_timer.stop() 
                self.project_label.setText("Плеер не найден")
                if not self.player_check_timer.isActive():
                    self.log("Плеер не запущен! Повторная проверка через 10 сек.", "warning")
                    self.player_check_timer.start(10000)

    def is_recorder_process_running(self):
        return self.find_window_by_process_name("recorder") is not None

    def init_recorder_checker(self):
        self.recorder_check_timer = QTimer(self); self.recorder_check_timer.timeout.connect(self.check_for_recorder)
        self.check_for_recorder() 

    def check_for_recorder(self):
        if self.is_recorder_process_running():
            self.auto_record_toggle_button.setEnabled(True)
            if self.recorder_check_timer.isActive(): self.recorder_check_timer.stop()
            return
        
        self.auto_record_toggle_button.setEnabled(False)
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop'); shortcut_found = False
        try:
            for filename in os.listdir(desktop_path):
                if 'recorder' in filename.lower() and filename.lower().endswith('.lnk'):
                    shortcut_path = os.path.join(desktop_path, filename)
                    self.log_request.emit("Camtasia не найдена. Запускаю...", "warning")
                    try:
                        os.startfile(shortcut_path); shortcut_found = True
                        if self.recorder_check_timer.isActive(): self.recorder_check_timer.stop()
                        QTimer.singleShot(3000, self.check_for_recorder)
                        return
                    except Exception as e:
                        self.log_request.emit(f"Не удалось запустить Camtasia: {e}", "error"); break
            if not shortcut_found: self.log_request.emit("Ярлык для Camtasia на рабочем столе не найден.", "error")
        except Exception as e: self.log_request.emit(f"Ошибка при поиске на рабочем столе: {e}", "error")
        if not self.recorder_check_timer.isActive(): self.recorder_check_timer.start(10000)

    def check_for_updates(self):
        self.log_request.emit("Проверка обновлений...", "info")
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10); response.raise_for_status()
            latest_release = response.json(); latest_version = latest_release.get("tag_name")
            if latest_version and latest_version > CURRENT_VERSION:
                self.log_request.emit(f"Доступна новая версия: {latest_version}. Автоматическое обновление...", "info")
                self.update_info = latest_release
                threading.Thread(target=self.apply_update, daemon=True).start()
            else: self.log_request.emit("Вы используете последнюю версию.", "info")
        except requests.RequestException as e: self.log_request.emit(f"Ошибка проверки обновлений: {e}", "error")
        except Exception as e: print(f"Неожиданная ошибка при проверке обновлений: {e}")

    def apply_update(self):
        assets = self.update_info.get("assets", []); download_url = None
        for asset in assets:
            if asset["name"] == ASSET_NAME: download_url = asset["browser_download_url"]; break
        if not download_url: self.log_request.emit("Не удалось найти ZIP-архив в релизе.", "error"); return
        self.download_and_run_updater(download_url)

    def download_and_run_updater(self, url):
        try:
            self.log_request.emit("Скачивание обновления...", "info")
            update_zip_name = "update.zip"
            response = requests.get(url, stream=True, timeout=60); response.raise_for_status()
            with open(update_zip_name, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
            self.log_request.emit("Распаковка архива...", "info")
            update_folder = "update_temp"
            if os.path.isdir(update_folder): import shutil; shutil.rmtree(update_folder)
            with zipfile.ZipFile(update_zip_name, 'r') as zip_ref: zip_ref.extractall(update_folder)
            self.log_request.emit("Обновление скачано. Перезапуск...", "info")
            updater_script_path = "updater.bat"
            current_exe_path = os.path.realpath(sys.executable if getattr(sys, 'frozen', False) else sys.argv[0])
            current_dir = os.path.dirname(current_exe_path)
            exe_name = os.path.basename(current_exe_path)
            with open(updater_script_path, "w", encoding="cp866") as f:
                f.write(f'@echo off\nchcp 65001 > NUL\necho Waiting for OiHelper to close...\ntaskkill /pid {os.getpid()} /f > NUL\ntimeout /t 2 /nobreak > NUL\necho Removing old files...\nrobocopy "{current_dir}\\{update_folder}" "{current_dir}" /e /move /is > NUL\nrd /s /q "{current_dir}\\{update_folder}"\necho Cleaning up...\ndel "{current_dir}\\{update_zip_name}"\necho Starting new version...\nstart "" "{exe_name}"\n(goto) 2>nul & del "%~f0"')
            subprocess.Popen([updater_script_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            QApplication.instance().quit()
        except Exception as e:
            self.log_request.emit(f"Ошибка при обновлении: {e}", "error"); print(f"Update error: {e}")

    def position_window_bottom_left(self):
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry(); margin = 65
                x = available_geometry.left() + margin
                y = available_geometry.bottom() - self.frameGeometry().height() - margin
                self.move(x, y)
        except Exception as e: print(f"Could not position window: {e}")

# --- Основной код для запуска ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    dark_palette = QPalette()
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45)); dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25)); dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white); dark_palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
    dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white); dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white); dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218)); dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
    app.setPalette(dark_palette)
    app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
