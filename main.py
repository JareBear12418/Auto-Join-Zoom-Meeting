from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QGridLayout, QMessageBox, QCheckBox
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QThread
from PyQt5.QtGui import QIcon
import sys
import openpyxl
from functools import partial
import pyautogui
import time

pyautogui.FAILSAFE = True

wb = openpyxl.load_workbook('List.xlsx')
sheet = wb['Sheet1']

meetings = [i for i in sheet.iter_rows(values_only=True) if i[0] != None]
meetings.pop(0)
meetings.sort()

directory = 'images/light/'

class Button(QPushButton):
    entered = pyqtSignal()
    leaved = pyqtSignal()

    def enterEvent(self, event):
        super().enterEvent(event)
        self.entered.emit()

    def leaveEvent(self, event):
        super().leaveEvent(event)
        self.leaved.emit()
        
class Worker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal()

    def __init__(self, meeting_id, meeting_passcode) -> None:
        super().__init__()
        self.meeting_id = meeting_id
        self.meeting_passcode = meeting_passcode
        
    def locate_image_coords(self, img: str) -> int:
        global directory
        start_time: time.time() = time.time()
        while True:
            coords = pyautogui.locateOnScreen(directory + img, confidence=0.9)
            if coords != None:
                break
            if 'light' in directory and (time.time() - start_time) > 4:
                directory = directory.replace('light', 'dark')
            elif 'dark' in directory and (time.time() - start_time) > 4:
                directory = directory.replace('dark', 'light')
            if (time.time() - start_time) > 10:
                return None
            time.sleep(0.1)
        return coords

    def press(self, location) -> None:
        pyautogui.click(location)

    def type_string(self, img: str, text: str) -> None:
        self.locate_image_coords(img)
        pyautogui.typewrite(str(text))

    def run(self, meeting_id, meeting_passcode):
        coords = self.locate_image_coords('zoom_icon.png')
        if coords is None:
            self.error.emit()
            return
        self.press(coords)
        # Locate JOIN button (large one)
        coords = self.locate_image_coords('BIG_join.png')
        if coords is None:
            self.error.emit()
            return
        # Press JOIN button (large one)
        self.press(coords)
        # Enter meeting ID
        self.type_string('meeting_id.png', meeting_id)
        # Press JOIN button (large one)
        coords = self.locate_image_coords('SMALL_join.png')
        if coords is None:
            self.error.emit()
            return
        
        self.press(coords)
        
        if meeting_passcode is not None:
            # Enter meeting passcode
            self.type_string('meeting_passcode.png', meeting_passcode)
            coords = self.locate_image_coords('join_meeting.png')
        if coords is None:
            self.error.emit()
            return
        self.press(coords)
        self.finished.emit()

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto Join Zoom Meeting")
        self.setFixedWidth(1300)
        self.setWindowIcon(QIcon('images/icon.ico'))
        self.dark_mode: bool = False
        layout: QVBoxLayout() = QVBoxLayout(self)
        grid_layout = QGridLayout(self)
        grid_layout.setSpacing(10)

        self.lbl: QLabel() = QLabel("Press the meeting you'd like to join:", self)
        self.lbl.setFixedHeight(120)
        layout.addWidget(self.lbl)

        colSize = 2
        for i, meeting in enumerate(meetings):
            button = Button(meeting[0], self)
            button.clicked.connect(partial(self.join_zoom_meeting, meeting[1], meeting[2]))
            button.entered.connect(self.handle_entered)
            button.setStyleSheet('QPushButton{border-radius: 25px; border-style: none; background-color: #0E72ED; color: white; font-size: 25px; font-weight: bold;} QPushButton:hover{background-color: #0D68D8;} QPushButton:pressed{background-color: #235DB4;}')
            button.leaved.connect(self.handle_leaved)
            if meeting[2] is None:
                button.setToolTip(f'ID: {meeting[1]}')
            else:
                button.setToolTip(f'ID: {meeting[1]}\nPasscode: {meeting[2]}')
            button.setFixedHeight(100)
            grid_layout.addWidget(button, int(i/colSize), int(i%colSize))
        layout.addLayout(grid_layout)

        self.checkBoxDarkTheme = QCheckBox(self)
        self.load_theme()
        self.checkBoxDarkTheme.setText("Dark mode")
        self.checkBoxDarkTheme.setStatusTip("Check this if zoom has darkmode enabled.")
        self.checkBoxDarkTheme.toggled.connect(self.toggle_theme)
        layout.addWidget(self.checkBoxDarkTheme)
        
        self.setLayout(layout)
        self.show()
    
    def load_theme(self):
        with open('zoom_theme', 'r') as f:
            text = f.read()
            self.dark_mode = text == 'true'
        self.checkBoxDarkTheme.setChecked(self.dark_mode)
        if self.dark_mode:
            self.checkBoxDarkTheme.setStyleSheet("QCheckBox::indicator { width:  20px; height: 20px;} QCheckBox {font-size: 25px; color: #F5F5F5; }")
            self.setStyleSheet("background-color: #363636")
            self.lbl.setStyleSheet('QLabel{font-size: 40px; color: #F5F5F5; font-weight: bold;}')
        else:
            self.checkBoxDarkTheme.setStyleSheet("QCheckBox::indicator { width:  20px; height: 20px;} QCheckBox {font-size: 25px; color: #39394D; }")
            self.setStyleSheet("background-color: #FFFFFF")
            self.lbl.setStyleSheet('QLabel{font-size: 40px; color: #39394D; font-weight: bold;}')
    
    def toggle_theme(self):
        with open('zoom_theme', 'w') as f:
            if self.dark_mode:
                f.write('false')
            else:
                f.write("true")
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self.checkBoxDarkTheme.setStyleSheet("QCheckBox::indicator { width:  20px; height: 20px; } QCheckBox {font-size: 25px; color: #F5F5F5;}")
            self.setStyleSheet("background-color: #363636")
            self.lbl.setStyleSheet('QLabel{font-size: 40px; color: #F5F5F5; font-weight: bold;}')
        else:
            self.checkBoxDarkTheme.setStyleSheet("QCheckBox::indicator { width:  20px; height: 20px; } QCheckBox {font-size: 25px; color: #39394D;}")
            self.setStyleSheet("background-color: #FFFFFF")
            self.lbl.setStyleSheet('QLabel{font-size: 40px; color: #39394D; font-weight: bold;}')
            
    def show_error_dialog(self):
        QMessageBox.critical(self, '10 second timeout has exceeded', "Can't proceed with the operation.\nMake sure Zoom is open.\nAborted.", QMessageBox.Ok)
    
    def handle_entered(self):
        QApplication.setOverrideCursor(Qt.PointingHandCursor)

    def handle_leaved(self):
        QApplication.restoreOverrideCursor()
    
    def join_zoom_meeting(self, meeting_id: str, meeting_passcode: str):
        self.thread = QThread()
        self.worker = Worker(meeting_id=meeting_id, meeting_passcode=meeting_passcode)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(partial(self.worker.run, meeting_id, meeting_passcode))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.error.connect(self.thread.quit)
        self.worker.error.connect(self.worker.deleteLater)
        self.worker.error.connect(self.show_error_dialog)
        self.thread.start()
        
app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec())