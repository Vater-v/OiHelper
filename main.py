from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel
from update_manager import UpdateManager
from config import AppConfig
import sys

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QCheckBox
)


class OiAutoMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.updater = UpdateManager()
        self.setWindowTitle(f"OiAuto v{AppConfig.CURRENT_VERSION}")
        self.setMinimumSize(480, 320)
        self._setup_ui()
        self._center_window()
        self.updater.log_request.connect(self.log_update)
        self.updater.status_update.connect(self.status_update)
        self.updater.check_finished.connect(self.update_check_finished)

    def _setup_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        self.label_status = QLabel("Готов к работе")
        self.btn_update = QPushButton("Проверить обновления")
        self.btn_update.clicked.connect(self.updater.check_for_updates)

        # 1. Тумблер автоматики
        self.checkbox_automation = QCheckBox("Автоматический режим")
        self.checkbox_automation.stateChanged.connect(self.toggle_automation)

        # 2. Кнопка "Расставить окна"
        self.btn_arrange_windows = QPushButton("Расставить окна")
        self.btn_arrange_windows.clicked.connect(self.arrange_windows)

        # 3. Кнопка "Вернуться" (скрытая)
        self.btn_return = QPushButton("Вернуться")
        self.btn_return.hide()
        self.btn_return.clicked.connect(self.return_to_main)

        layout.addWidget(self.label_status)
        layout.addWidget(self.checkbox_automation)
        layout.addWidget(self.btn_arrange_windows)
        layout.addWidget(self.btn_update)
        layout.addWidget(self.btn_return)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        QApplication.setStyle("Fusion")

    def toggle_automation(self, state):
        if state:
            self.label_status.setText("Автоматический режим: ВКЛ")
            # здесь можно включить автоматику
        else:
            self.label_status.setText("Автоматический режим: ВЫКЛ")
            # здесь можно выключить автоматику

    def arrange_windows(self):
        # логика расстановки окон
        self.label_status.setText("Окна расставлены!")

    def log_update(self, msg, level):
        self.label_status.setText(f"[{level}] {msg}")
        # Показать кнопку только при fail или warning
        if "fail" in msg.lower() or level in ("warning", "error"):
            self.btn_return.show()

    def status_update(self, msg):
        self.label_status.setText(msg)

    def update_check_finished(self):
        pass

    def return_to_main(self):
        self.btn_return.hide()
        self.label_status.setText("Готов к работе")

    def _center_window(self):
        frameGm = self.frameGeometry()
        screen = QApplication.primaryScreen().availableGeometry().center()
        frameGm.moveCenter(screen)
        self.move(frameGm.topLeft())

def main():
    app = QApplication(sys.argv)
    window = OiAutoMainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
