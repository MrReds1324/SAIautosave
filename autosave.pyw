import ctypes
import os
import sys

from PyQt5.QtCore import QCoreApplication, QTimer, Qt, QRect
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSystemTrayIcon, QApplication, QMenu, QWidget, QLabel, QVBoxLayout, QDialog, QSlider, QHBoxLayout, QGroupBox
from pyautogui import hotkey
from win32gui import GetWindowText, GetForegroundWindow

# Hacky way to set task bar icon
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('sai.autosave')

SAI_NAME_STRING = "PaintTool SAI Ver.2"
if getattr(sys, 'frozen', False):
    ICON_IMAGE = QIcon(os.path.join(sys._MEIPASS, "resources/save.svg"))
else:
    ICON_IMAGE = QIcon("resources/save.svg")

import sys


def convert_min_to_milli(minute_time: int) -> int:
    return minute_time * 60 * 1000


class SliderDialog(QDialog):

    def __init__(self):
        super().__init__()
        self.timer_value: int = 1

        self.setWindowTitle("Auto Save Period")
        self.slider_group = QGroupBox(self)
        self.slider_group.setGeometry(QRect(10, 10, 221, 81))

        self.vertical_layout_widget = QWidget(self.slider_group)
        self.vertical_layout_widget.setGeometry(QRect(10, 20, 201, 61))
        self.timer_layout = QVBoxLayout(self.vertical_layout_widget)
        self.horizontal_layout = QHBoxLayout()

        self.timer_label = QLabel(self.vertical_layout_widget)
        self.timer_label.setText("Auto Save Period (Minutes)")
        self.timer_label.setAlignment(Qt.AlignHCenter)
        self.timer_layout.addWidget(self.timer_label)

        self.timer_slider = QSlider(Qt.Horizontal, self.vertical_layout_widget)
        self.timer_slider.setMaximum(30)
        self.timer_slider.setMinimum(1)
        self.horizontal_layout.addWidget(self.timer_slider)

        self.timer_value_label = QLabel(self.vertical_layout_widget)
        self.timer_value_label.setNum(1)
        self.horizontal_layout.addWidget(self.timer_value_label)

        self.timer_layout.addLayout(self.horizontal_layout)

        self.timer_slider.valueChanged['int'].connect(self.timer_value_label.setNum)
        self.timer_slider.valueChanged['int'].connect(self.update_timer_value)

        self.setGeometry(700, 200, 241, 101)
        self.setFixedSize(241, 101)

    def closeEvent(self, event):
        event.ignore()
        self.setHidden(True)

    def update_timer_value(self, value: int) -> None:
        self.timer_value = value


class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, icon, parent=None):
        QSystemTrayIcon.__init__(self, icon, parent)
        self.paused: bool = False
        self.timer_window = SliderDialog()
        self.menu = QMenu(parent)

        # Menu Creation
        self.change_action = self.menu.addAction("Change Time")
        self.pause_action = self.menu.addAction("Pause")
        self.exit_action = self.menu.addAction("Exit")

        self.setContextMenu(self.menu)

        # Event Handlers
        self.exit_action.triggered.connect(self.exit_tray)
        self.change_action.triggered.connect(self.change_timer)
        self.pause_action.triggered.connect(self.pause_timer)

        # Timer for handling auto saving
        self.timer = QTimer()
        self.timer.setInterval(convert_min_to_milli(1))
        self.timer.timeout.connect(self.recurring_timer)
        self.timer.start()

    def exit_tray(self):
        QCoreApplication.exit(0)

    def recurring_timer(self):
        if GetWindowText(GetForegroundWindow()).startswith(SAI_NAME_STRING):
            hotkey('ctrl', 's')

    def change_timer(self):
        self.timer_window.exec_()
        self.timer.setInterval(convert_min_to_milli(self.timer_window.timer_value))

    def pause_timer(self):
        self.paused = not self.paused
        if self.paused:
            self.pause_action.setText("Resume")
            self.timer.stop()
        else:
            self.pause_action.setText("Pause")
            self.timer.start()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(ICON_IMAGE)
    trayIcon = SystemTrayIcon(ICON_IMAGE)

    trayIcon.show()
    sys.exit(app.exec_())
