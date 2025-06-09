import os
import sys
import json
import ctypes
import keyboard
import win32api
import win32con
import pyautogui
import threading
import pyperclip
import pydirectinput as pdi
from PyQt6.QtGui import *
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *


SendInput = ctypes.windll.user32.SendInput
MapVirtualKey = ctypes.windll.user32.MapVirtualKeyW
MAPVK_VK_TO_VSC = 0

VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'clear':0x0C,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'pause':0x13,
           'capslock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'pageup':0x21,
           'pagedown':0x22,
           'end':0x23,
           'home':0x24,
           'left':0x25,
           'up':0x26,
           'right':0x27,
           'down':0x28,
           'select':0x29,
           'print':0x2A,
           'execute':0x2B,
           'printscreen':0x2C,
           'ins':0x2D,
           'del':0x2E,
           'help':0x2F,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'numpad0':0x60,
           'numpad1':0x61,
           'numpad2':0x62,
           'numpad3':0x63,
           'numpad4':0x64,
           'numpad5':0x65,
           'numpad6':0x66,
           'numpad7':0x67,
           'numpad8':0x68,
           'numpad9':0x69,
           'multiply':0x6A,
           'add':0x6B,
           'separator':0x6C,
           'subtract':0x6D,
           'decimal':0x6E,
           'divide':0x6F,
           'F1':0x70,
           'F2':0x71,
           'F3':0x72,
           'F4':0x73,
           'F5':0x74,
           'F6':0x75,
           'F7':0x76,
           'F8':0x77,
           'F9':0x78,
           'F10':0x79,
           'F11':0x7A,
           'F12':0x7B,
           'F13':0x7C,
           'F14':0x7D,
           'F15':0x7E,
           'F16':0x7F,
           'F17':0x80,
           'F18':0x81,
           'F19':0x82,
           'F20':0x83,
           'F21':0x84,
           'F22':0x85,
           'F23':0x86,
           'F24':0x87,
           'numlock':0x90,
           'scrolllock':0x91,
           'shiftleft':0xA0,
           'shiftright ':0xA1,
           'ctrlleft':0xA2,
           'ctrlright':0xA3,
           'altleft':0xA4,
           'altright':0xA5,
           'browser_back':0xA6,
           'browser_forward':0xA7,
           'browser_refresh':0xA8,
           'browser_stop':0xA9,
           'browser_search':0xAA,
           'browser_favorites':0xAB,
           'browser_start_and_home':0xAC,
           'volume_mute':0xAD,
           'volume_Down':0xAE,
           'volume_up':0xAF,
           'next_track':0xB0,
           'previous_track':0xB1,
           'stop_media':0xB2,
           'play/pause_media':0xB3,
           'start_mail':0xB4,
           'select_media':0xB5,
           'start_application_1':0xB6,
           'start_application_2':0xB7,
           'attn_key':0xF6,
           'crsel_key':0xF7,
           'exsel_key':0xF8,
           'play_key':0xFA,
           'zoom_key':0xFB,
           'clear_key':0xFE,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,
           '`':0xC0}

KEYBOARD_MAPPING = {
    'escape': 0x01,
    'esc': 0x01,
    'f1': 0x3B,
    'f2': 0x3C,
    'f3': 0x3D,
    'f4': 0x3E,
    'f5': 0x3F,
    'f6': 0x40,
    'f7': 0x41,
    'f8': 0x42,
    'f9': 0x43,
    'f10': 0x44,
    'f11': 0x57,
    'f12': 0x58,
    'printscreen': 0xB7,
    'prntscrn': 0xB7,
    'prtsc': 0xB7,
    'prtscr': 0xB7,
    'scrolllock': 0x46,
    'pause': 0xC5,
    '`': 0x29,
    '1': 0x02,
    '2': 0x03,
    '3': 0x04,
    '4': 0x05,
    '5': 0x06,
    '6': 0x07,
    '7': 0x08,
    '8': 0x09,
    '9': 0x0A,
    '0': 0x0B,
    '-': 0x0C,
    '=': 0x0D,
    'backspace': 0x0E,
    'insert': 0xD2 + 1024,
    'home': 0xC7 + 1024,
    'pageup': 0xC9 + 1024,
    'pagedown': 0xD1 + 1024,
    # numpad
    'numlock': 0x45,
    'divide': 0xB5 + 1024,
    'multiply': 0x37,
    'subtract': 0x4A,
    'add': 0x4E,
    'decimal': 0x53,
    'numpadenter': 0x9C + 1024,
    'numpad1': 0x4F,
    'numpad2': 0x50,
    'numpad3': 0x51,
    'numpad4': 0x4B,
    'numpad5': 0x4C,
    'numpad6': 0x4D,
    'numpad7': 0x47,
    'numpad8': 0x48,
    'numpad9': 0x49,
    'numpad0': 0x52,
    # end numpad
    'tab': 0x0F,
    'q': 0x10,
    'w': 0x11,
    'e': 0x12,
    'r': 0x13,
    't': 0x14,
    'y': 0x15,
    'u': 0x16,
    'i': 0x17,
    'o': 0x18,
    'p': 0x19,
    '[': 0x1A,
    ']': 0x1B,
    '\\': 0x2B,
    'del': 0xD3 + 1024,
    'delete': 0xD3 + 1024,
    'end': 0xCF + 1024,
    'capslock': 0x3A,
    'a': 0x1E,
    's': 0x1F,
    'd': 0x20,
    'f': 0x21,
    'g': 0x22,
    'h': 0x23,
    'j': 0x24,
    'k': 0x25,
    'l': 0x26,
    ';': 0x27,
    "'": 0x28,
    'enter': 0x1C,
    'return': 0x1C,
    'shift': 0x2A,
    'shiftleft': 0x2A,
    'z': 0x2C,
    'x': 0x2D,
    'c': 0x2E,
    'v': 0x2F,
    'b': 0x30,
    'n': 0x31,
    'm': 0x32,
    ',': 0x33,
    '.': 0x34,
    '/': 0x35,
    'shiftright': 0x36,
    'ctrl': 0x1D,
    'ctrlleft': 0x1D,
    'win': 0xDB + 1024,
    'winleft': 0xDB + 1024,
    'alt': 0x38,
    'altleft': 0x38,
    ' ': 0x39,
    'space': 0x39,
    'altright': 0xB8 + 1024,
    'winright': 0xDC + 1024,
    'apps': 0xDD + 1024,
    'ctrlright': 0x9D + 1024,
    # arrow key scancodes can be different depending on the hardware,
    # so I think the best solution is to look it up based on the virtual key
    # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mapvirtualkeya?redirectedfrom=MSDN
    'up': MapVirtualKey(0x26, MAPVK_VK_TO_VSC),
    'left': MapVirtualKey(0x25, MAPVK_VK_TO_VSC),
    'down': MapVirtualKey(0x28, MAPVK_VK_TO_VSC),
    'right': MapVirtualKey(0x27, MAPVK_VK_TO_VSC),
}

# C struct redefinitions 
PUL = ctypes.POINTER(ctypes.c_ulong)
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]

def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

class QCustomListWidget(QListWidget):
    def __init__(self):
        super().__init__()

        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragEnabled(True)
        self.viewport().setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
    
    def dropEvent(self, event):
        super().dropEvent(event)
        items = [self.item(i) for i in range(self.count())]
        ex.signalClass.handleDNDSignal.emit(items)

class QCustomListWidgetItem(QWidget):
    def __init__(self, parent = None):
        super(QCustomListWidgetItem, self).__init__(parent)

        layout = QHBoxLayout()
        self.actionLbl = QLabel(self)
        print(self.actionLbl.styleSheet())
        self.deleteBtn = QPushButton(self)
        self.deleteBtn.clicked.connect(self.emitDeleteSignal)
        self.deleteBtn.setText("X")
        layout.addWidget(self.actionLbl)
        layout.addWidget(self.deleteBtn)
        self.setLayout(layout)
    
    def setLblText(self, text: str):
        self.actionLbl.setText(text)
    
    def emitDeleteSignal(self):
        ex.signalClass.deleteActionSignal.emit(self)

class SignalClass(QObject):
    addKeySignal = pyqtSignal(list)
    addMouseSignal = pyqtSignal(list)
    addDelaySignal = pyqtSignal(list)
    addWriteSignal = pyqtSignal(list)
    deleteActionSignal = pyqtSignal(QCustomListWidgetItem)
    showMessageSignal = pyqtSignal(int, str)
    handleDNDSignal = pyqtSignal(list)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        os.makedirs("./preset", exist_ok=True)

        QFontDatabase.addApplicationFont(resourcePath("dependencies/NanumGothic-Regular.ttf"))
        self.setFont(QFont("NanumGothic"))
        self.setWindowIcon(QIcon(resourcePath("dependencies/pymacro.ico")))

        self.stopEvent = threading.Event()
        self.pressedKey = ""

        self.signalClass = SignalClass()
        self.signalClass.addKeySignal.connect(self.addKeyToQueue)
        self.signalClass.addMouseSignal.connect(self.addMouseToQueue)
        self.signalClass.addDelaySignal.connect(self.addDelayToQueue)
        self.signalClass.addWriteSignal.connect(self.addWriteToQueue)
        self.signalClass.deleteActionSignal.connect(self.deleteActionFromQueue)
        self.signalClass.showMessageSignal.connect(self.showMessage)
        self.signalClass.handleDNDSignal.connect(self.handleDND)
        
        queueLayout = QVBoxLayout()
        self.queue = QCustomListWidget()
        self.clwContainer = []
        self.queueList = []
        self.itemsList = []
        queueLayout.addWidget(self.queue)

        buttonsLayout = QVBoxLayout()
        buttonsLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        addMouseACtionBtn = QPushButton("마우스 입력 추가", self)
        addMouseACtionBtn.clicked.connect(self.addMouseAction)
        addKeyboardActionBtn = QPushButton("키보드 입력 추가", self)
        addKeyboardActionBtn.clicked.connect(self.addKeyboardAction)
        addWaitingActionBtn = QPushButton("대기시간 추가", self)
        addWaitingActionBtn.clicked.connect(self.addDelayAction)
        addWriteActionBtn = QPushButton("문자 입력 추가", self)
        addWriteActionBtn.clicked.connect(self.addWriteAction)
        buttonsLayout.addWidget(addMouseACtionBtn, 0)
        buttonsLayout.addWidget(addKeyboardActionBtn, 0)
        buttonsLayout.addWidget(addWaitingActionBtn, 0)
        buttonsLayout.addWidget(addWriteActionBtn, 0)
        buttonsLayout.addStretch(5)

        presetLayout = QVBoxLayout()
        presetLayout.setAlignment(Qt.AlignmentFlag.AlignBottom)
        savePresetBtn = QPushButton("프리셋 저장하기", self)
        savePresetBtn.clicked.connect(self.savePresetAsFile)
        loadPresetBtn = QPushButton("프리셋 불러오기", self)
        loadPresetBtn.clicked.connect(self.loadPresetFromFile)
        presetLayout.addWidget(savePresetBtn)
        presetLayout.addWidget(loadPresetBtn)
        buttonsLayout.addLayout(presetLayout)

        controlGroupBox = QGroupBox("컨트롤", self)
        controlLayout = QVBoxLayout()
        self.startMacroBtn = QPushButton("시작", self)
        self.startMacroBtn.clicked.connect(self.startMacroThread)
        self.stopMacroBtn = QPushButton("종료", self)
        self.stopMacroBtn.clicked.connect(self.stopMacro)
        self.shortcut = QShortcut(QKeySequence("Ctrl+Q"), self)
        self.shortcut.activated.connect(self.stopMacro)
        repeatLbl = QLabel("반복할 횟수 (0은 무한반복) : ", self)
        self.repeatQle = QLineEdit(self)
        self.repeatQle.setText("0")
        onlyInt = QIntValidator(self)
        onlyInt.setBottom(0)
        self.repeatQle.setValidator(onlyInt)
        controlLayout.addWidget(repeatLbl)
        controlLayout.addWidget(self.repeatQle)
        controlLayout.addWidget(self.startMacroBtn)
        controlLayout.addWidget(self.stopMacroBtn)
        controlGroupBox.setLayout(controlLayout)
        buttonsLayout.addWidget(controlGroupBox, 2)

        mainLayout = QHBoxLayout()
        mainLayout.addLayout(queueLayout, stretch=1)
        mainLayout.addLayout(buttonsLayout, stretch=1)
        
        self.setWindowTitle("Py Macro")
        self.setLayout(mainLayout)
        self.resize(720, 600)
        self.center()
        self.show()
    
    def center(self):
        qr = self.frameGeometry()
        cp = QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    
    @pyqtSlot(list)
    def handleDND(self, items):
        indexList = []
        for i in range(len(items)):
            if(self.itemsList[i] != items[i]):
                indexList.append(i)
        if indexList:
            self.itemsList[indexList[0]], self.itemsList[indexList[1]] = self.itemsList[indexList[1]], self.itemsList[indexList[0]]
            self.queueList[indexList[0]], self.queueList[indexList[1]] = self.queueList[indexList[1]], self.queueList[indexList[0]]
            self.clwContainer[indexList[0]], self.clwContainer[indexList[1]] = self.clwContainer[indexList[1]], self.clwContainer[indexList[0]]
        
        print(self.queueList)

    @pyqtSlot(int, str)
    def showMessage(self, msgType, content):
        if msgType == 0:
            msg = QMessageBox.information(self, "알림", content)
    
    def savePresetAsFile(self):
        try:
            if(len(self.queueList)):
                jsonList = json.dumps(self.queueList)
                print(jsonList)
                name, _ = QFileDialog.getSaveFileName(self, "파일 저장하기", "./preset", filter="*.json")
                if(name != ""):
                    with open(name, "w") as f:
                        f.write(jsonList)
            else:
                self.signalClass.showMessageSignal.emit(0, "최소 1개의 액션을 추가해야 저장할 수 있습니다.")
        except:
            pass
    
    def loadPresetFromFile(self):
        name, _ = QFileDialog.getOpenFileName(self, "파일 불러오기", "./preset", filter="*.json")
        try:
            with open(name, "r") as f:
                data = f.read()
            jsonList = json.loads(data)

            self.__init__()

            for i in jsonList:
                if(i[0] == "mouse"):
                    self.addMouseToQueue(i)
                elif(i[0] == "key"):
                    self.addKeyToQueue(i)
                elif(i[0] == "delay"):
                    self.addDelayToQueue(i)
                elif(i[0] == "write"):
                    self.addWriteToQueue(i)
        except:
            pass

    def addKeyboardAction(self):
        self.kw = KeyboardInputWindow()
        self.kw.show()
    
    def addMouseAction(self):
        self.mw = MouseInputWindow()
        self.mw.show()
    
    def addDelayAction(self):
        self.dw = DelayInputWindow()
        self.dw.show()
    
    def addWriteAction(self):
        self.ww = WriteInputWindow()
        self.ww.show()
    
    def addClwToQueue(self, text):
        myClw = QCustomListWidgetItem(self)
        myClw.setLblText(text)
        myQListWidgetItem = QListWidgetItem(self.queue)
        myQListWidgetItem.setSizeHint(myClw.sizeHint())
        self.itemsList = [self.queue.item(i) for i in range(self.queue.count())]
        self.queue.addItem(myQListWidgetItem)
        self.clwContainer.append(myClw)
        self.queue.setItemWidget(myQListWidgetItem, myClw)
    
    @pyqtSlot(list)
    def addMouseToQueue(self, mouse):
        self.queueList.append(mouse)
        print(self.queueList)
        leftRightMode, clickMode = mouse[1]
        if(leftRightMode == 0):
            clickDir = "왼쪽"
        elif(leftRightMode == 1):
            clickDir = "오른쪽"
        x, y = mouse[2]
        if(clickMode == 0):
            self.addClwToQueue("%d, %d %s 눌렀다 떼기" % (x, y, clickDir))
        elif(clickMode == 1):
            self.addClwToQueue("%d, %d %s 누르기" % (x, y, clickDir))
        elif(clickMode == 2):
            self.addClwToQueue("%d, %d %s 떼기" % (x, y, clickDir))
    
    @pyqtSlot(list)
    def addKeyToQueue(self, key): 
        self.queueList.append(key)
        print(self.queueList)
        mode, keyStr = key[1:3]
        if(mode == 0):
            self.addClwToQueue("%s 눌렀다 떼기" % keyStr)
        elif(mode == 1):
            self.addClwToQueue("%s 누르기" % keyStr)
        elif(mode == 2):
            self.addClwToQueue("%s 떼기" % keyStr)

    @pyqtSlot(list)
    def addDelayToQueue(self, delay):
        self.queueList.append(delay)
        print(self.queueList)
        self.addClwToQueue("%0.2f초 대기하기" % delay[1])
    
    @pyqtSlot(list)
    def addWriteToQueue(self, write):
        self.queueList.append(write)
        print(self.queueList)
        self.addClwToQueue("%s 입력하기" % write[1])
    
    @pyqtSlot(QCustomListWidgetItem)
    def deleteActionFromQueue(self, element):
        index = self.clwContainer.index(element)
        self.queue.takeItem(index)
        del self.queueList[index]
        del self.clwContainer[index]
        print(self.queueList)
        print(self.clwContainer)
    
    def startMacroThread(self):
        self.macroThr = threading.Thread(target=self.startMacro)
        self.macroThr.start()
        self.startMacroBtn.setDisabled(True)
        self.startMacroBtn.setText("실행 중")
    
    def startMacro(self):
        repeatNum = int(self.repeatQle.text())
        if(repeatNum == 0):
            while not self.stopEvent.is_set():
                self.processQueueList()
        elif(repeatNum > 0):
            for i in range(repeatNum):
                self.processQueueList()
        
    def processQueueList(self):
        for i in self.queueList:
            if(i[0] == "mouse"):
                leftRightMode, clickMode = i[1]
                if(leftRightMode == 0):
                    df = win32con.MOUSEEVENTF_LEFTDOWN
                    uf = win32con.MOUSEEVENTF_LEFTUP
                elif(leftRightMode == 1):
                    df = win32con.MOUSEEVENTF_RIGHTDOWN
                    uf = win32con.MOUSEEVENTF_RIGHTUP
                x, y = i[2]
                win32api.SetCursorPos((x, y))
                if(clickMode == 0):
                    win32api.mouse_event(df, x, y, 0, 0)
                    win32api.mouse_event(uf, x, y, 0, 0)
                elif(clickMode == 1):
                    win32api.mouse_event(df, x, y, 0, 0)
                elif(clickMode == 2):
                    win32api.mouse_event(uf, x, y, 0, 0)
            elif(i[0] == "key"):
                mode, key = i[1:3]
                if(mode == 0):
                    pdi.keyDown(key)
                    pdi.keyUp(key)
                elif(mode == 1):
                    pdi.keyDown(key)
                elif(mode == 2):
                    pdi.keyUp(key)
            elif(i[0] == "delay"):
                self.stopEvent.wait(i[1])
            elif(i[0] == "write"):
                pyperclip.copy(i[1])
                pdi.keyDown('ctrl')
                pdi.keyDown('v')
                pdi.keyUp('ctrl')
                pdi.keyUp('v')
    
    def stopMacro(self):
        self.startMacroBtn.setEnabled(True)
        self.startMacroBtn.setText("시작")
        self.stopEvent.set()

class MouseInputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFont(QFont("NanumGothic"))
        self.setWindowIcon(QIcon(resourcePath("dependencies/pymacro.ico")))
        layout = QVBoxLayout()
        positionGbx = QGroupBox("좌표", self)
        self.positionLayout = QHBoxLayout()
        self.xPosLbl = QLabel("X좌표 : ", self)
        self.yPosLbl = QLabel("Y좌표 : ", self)
        self.manualCbx = QCheckBox("직접 입력", self)
        self.manualCbx.stateChanged.connect(self.toggleManualAutomatic)
        self.leftRightLayout = QHBoxLayout()
        self.leftRightMode = 0
        self.leftRightGrp = QButtonGroup(self)
        self.leftRbt = QRadioButton("왼쪽 클릭", self)
        self.leftRbt.setChecked(True)
        self.leftRbt.clicked.connect(self.setLeftRightType)
        self.rightRbt = QRadioButton("오른쪽 클릭", self)
        self.rightRbt.clicked.connect(self.setLeftRightType)
        self.leftRightGrp.addButton(self.leftRbt)
        self.leftRightGrp.addButton(self.rightRbt)
        self.leftRightLayout.addWidget(self.leftRbt)
        self.leftRightLayout.addWidget(self.rightRbt)
        self.pressAndReleaseRbt = QRadioButton("눌렀다 떼기", self)
        self.pressAndReleaseRbt.clicked.connect(self.setClickType)
        self.pressAndReleaseRbt.setChecked(True)
        self.clickMode = 0
        self.pressRbt = QRadioButton("누르기", self)
        self.pressRbt.clicked.connect(self.setClickType)
        self.releaseRbt = QRadioButton("떼기", self)
        self.releaseRbt.clicked.connect(self.setClickType)
        self.cbxLayout = QHBoxLayout()
        self.cbxLayout.addWidget(self.pressAndReleaseRbt)
        self.cbxLayout.addWidget(self.pressRbt)
        self.cbxLayout.addWidget(self.releaseRbt)
        self.addKeyBtn = QPushButton("추가 (Enter)", self)
        self.addKeyBtn.setShortcut("Return")
        self.addKeyBtn.clicked.connect(self.emitAddMouseignal)
        self.cancelBtn = QPushButton("취소 (Esc)", self)
        self.cancelBtn.setShortcut("Esc")
        self.cancelBtn.clicked.connect(self.destroyWindow)
        self.positionLayout.addWidget(self.xPosLbl)
        self.positionLayout.addWidget(self.yPosLbl)
        positionGbx.setLayout(self.positionLayout)
        
        layout.addWidget(positionGbx)
        layout.addWidget(self.manualCbx)
        layout.addLayout(self.leftRightLayout)
        layout.addLayout(self.cbxLayout)
        layout.addWidget(self.addKeyBtn)
        layout.addWidget(self.cancelBtn)
        self.setLayout(layout)
        self.setFixedSize(350, 250)
        self.setWindowTitle("마우스 입력 추가")
        self.setMouseTracking(True)

        timer = QTimer(self)
        timer.timeout.connect(self.getCurrentPosition)
        timer.start(100)
    
    def toggleManualAutomatic(self, state):
        while(self.positionLayout.count()):
            child = self.positionLayout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        if(state == Qt.CheckState.Checked.value):
            self.xPosLbl = QLabel("X좌표 : ", self)
            self.xPosLbl.setFont(QFont("NanumGothic"))
            self.yPosLbl = QLabel("Y좌표 : ", self)
            self.yPosLbl.setFont(QFont("NanumGothic"))
            self.xPosQle = QLineEdit(self)
            self.xPosQle.setAlignment(Qt.AlignmentFlag.AlignVCenter)
            self.xPosQle.setFont(QFont("NanumGothic"))
            self.yPosQle = QLineEdit(self)
            self.yPosQle.setFont(QFont("NanumGothic"))
            self.positionLayout.addWidget(self.xPosLbl)
            self.positionLayout.addWidget(self.xPosQle)
            self.positionLayout.addWidget(self.yPosLbl)
            self.positionLayout.addWidget(self.yPosQle)
        else:
            self.xPosLbl = QLabel("X좌표 : ", self)
            self.xPosLbl.setFont(QFont("NanumGothic"))
            self.yPosLbl = QLabel("Y좌표 : ", self)
            self.yPosLbl.setFont(QFont("NanumGothic"))
            self.positionLayout.addWidget(self.xPosLbl)
            self.positionLayout.addWidget(self.yPosLbl)

    def setLeftRightType(self):
        if(self.leftRbt.isChecked()):
            self.leftRightMode = 0
        elif(self.rightRbt.isChecked()):
            self.leftRightMode = 1
    
    def setClickType(self):
        if(self.pressAndReleaseRbt.isChecked()):
            self.clickMode = 0
        elif(self.pressRbt.isChecked()):
            self.clickMode = 1
        elif(self.releaseRbt.isChecked()):
            self.clickMode = 2
    
    def getCurrentPosition(self):
        if not self.manualCbx.checkState() == Qt.CheckState.Checked:
            self.xPos = pyautogui.position().x
            self.yPos = pyautogui.position().y
            self.xPosLbl.setText("X좌표 : %s" % self.xPos)
            self.yPosLbl.setText("Y좌표 : %s" % pyautogui.position().y)
    
    def emitAddMouseignal(self):
        if not self.manualCbx.checkState() == Qt.CheckState.Checked:
            mouseStruct = ["mouse", [self.leftRightMode, self.clickMode], [self.xPos, self.yPos]]
        else:
            mouseStruct = ["mouse", [self.leftRightMode, self.clickMode], [int(self.xPosQle.text()), int(self.yPosQle.text())]]
        ex.signalClass.addMouseSignal.emit(mouseStruct)
        self.destroyWindow()
    
    def destroyWindow(self):
        self.destroy()

class KeyboardInputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFont(QFont("NanumGothic"))
        self.setWindowIcon(QIcon(resourcePath("dependencies/pymacro.ico")))
        layout = QVBoxLayout()
        self.inputLabel = QLabel("키를 입력하세요 : ", self)
        self.keyGbx = QGroupBox("키", self)
        self.keyLayout = QHBoxLayout()
        self.keyLayout.addWidget(self.inputLabel)
        self.keyGbx.setLayout(self.keyLayout)
        # self.inputLabel.setFixedWidth(200)
        # self.inputLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.keyMode = 0
        self.pressedKey = ""
        self.pressAndReleaseRbt = QRadioButton("눌렀다 떼기", self)
        self.pressAndReleaseRbt.clicked.connect(self.setKeyMode)
        self.pressAndReleaseRbt.setChecked(True)
        self.pressRbt = QRadioButton("누르기", self)
        self.pressRbt.clicked.connect(self.setKeyMode)
        self.releaseRbt = QRadioButton("떼기", self)
        self.releaseRbt.clicked.connect(self.setKeyMode)
        self.cbxLayout = QHBoxLayout()
        self.cbxLayout.addWidget(self.pressAndReleaseRbt)
        self.cbxLayout.addWidget(self.pressRbt)
        self.cbxLayout.addWidget(self.releaseRbt)
        self.addKeyBtn = QPushButton("추가", self)
        # self.addKeyBtn.setShortcut("Return")
        self.addKeyBtn.clicked.connect(self.emitAddKeySignal)
        self.cancelBtn = QPushButton("취소", self)
        # self.cancelBtn.setShortcut("Esc")
        self.cancelBtn.clicked.connect(self.destroyWindow)
        layout.addWidget(self.keyGbx)
        layout.addLayout(self.cbxLayout)
        layout.addWidget(self.addKeyBtn)
        layout.addWidget(self.cancelBtn)
        self.setLayout(layout)
        self.setFixedSize(350, 250)
        self.setWindowTitle("키보드 입력 추가")
    
    def setKeyMode(self):
        if(self.pressAndReleaseRbt.isChecked()):
            self.keyMode = 0
        elif(self.pressRbt.isChecked()):
            self.keyMode = 1
        elif(self.releaseRbt.isChecked()):
            self.keyMode = 2
    
    def keyPressEvent(self, event):
        self.grabKeyboard()
        super().keyPressEvent(event)
        for i in VK_CODE:
            if win32api.GetAsyncKeyState(VK_CODE[i]) != 0:
                self.pressedKey = i
                self.inputLabel.setText("키를 입력하세요 : " + i)
        print(keyboard.read_key())
        self.releaseKeyboard()
    
    def emitAddKeySignal(self):
        keyStruct = ["key", self.keyMode, self.pressedKey]
        ex.signalClass.addKeySignal.emit(keyStruct)
        self.destroyWindow()
    
    def destroyWindow(self):
        self.destroy()

class DelayInputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFont(QFont("NanumGothic"))
        self.setWindowIcon(QIcon(resourcePath("dependencies/pymacro.ico")))
        layout = QGridLayout()
        self.delayQbx = QGroupBox("대기시간", self)
        delayLayout = QHBoxLayout()
        self.delayLbl = QLabel("대기시간(sec) : ", self)
        self.delayQle = QLineEdit(self)
        self.doubleValidator = QDoubleValidator()
        self.doubleValidator.setRange(0, 21600, 2)
        self.delayQle.setValidator(self.doubleValidator)
        delayLayout.addWidget(self.delayLbl)
        delayLayout.addWidget(self.delayQle)
        self.delayQbx.setLayout(delayLayout)
        self.addKeyBtn = QPushButton("추가 (Enter)", self)
        self.addKeyBtn.setShortcut("Return")
        self.addKeyBtn.clicked.connect(self.emitAddDelaySignal)
        self.cancelBtn = QPushButton("취소 (Esc)", self)
        self.cancelBtn.setShortcut("Esc")
        self.cancelBtn.clicked.connect(self.destroyWindow)
        layout.addWidget(self.delayQbx, 0, 0)
        layout.addWidget(self.addKeyBtn, 1, 0, 1, -1)
        layout.addWidget(self.cancelBtn, 2, 0, 1, -1)
        self.setLayout(layout)
        self.setFixedSize(350, 250)
        self.setWindowTitle("대기시간 추가")
    
    def emitAddDelaySignal(self):
        if(isFloat(self.delayQle.text())):
            delay = float(self.delayQle.text())
            delayStruct = ["delay", delay]
            ex.signalClass.addDelaySignal.emit(delayStruct)
        self.destroyWindow()
    
    def destroyWindow(self):
        self.destroy()

class WriteInputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFont(QFont("NanumGothic"))
        self.setWindowIcon(QIcon(resourcePath("dependencies/pymacro.ico")))
        layout = QGridLayout()
        self.writeQbx = QGroupBox("문자열", self)
        writeLayout = QHBoxLayout()
        self.writeLbl = QLabel("문자열 : ", self)
        self.writeQle = QLineEdit(self)
        writeLayout.addWidget(self.writeLbl)
        writeLayout.addWidget(self.writeQle)
        self.writeQbx.setLayout(writeLayout)
        self.addWriteBtn = QPushButton("추가 (Enter)", self)
        self.addWriteBtn.setShortcut("Return")
        self.addWriteBtn.clicked.connect(self.emitAddWriteSignal)
        self.cancelBtn = QPushButton("취소 (Esc)", self)
        self.cancelBtn.setShortcut("Esc")
        self.cancelBtn.clicked.connect(self.destroyWindow)
        layout.addWidget(self.writeQbx, 0, 0)
        layout.addWidget(self.addWriteBtn, 1, 0, 1, -1)
        layout.addWidget(self.cancelBtn, 2, 0, 1, -1)
        self.setLayout(layout)
        self.setFixedSize(350, 250)
        self.setWindowTitle("문자열 추가")
    
    def emitAddWriteSignal(self):
        write = self.writeQle.text()
        writeStruct = ["write", write]
        ex.signalClass.addWriteSignal.emit(writeStruct)
        self.destroyWindow()
    
    def destroyWindow(self):
        self.destroy()

def resourcePath(relativePath):
	try:
		basePath = sys._MEIPASS
	except Exception:
		basePath = os.path.abspath(".")
	return os.path.join(basePath, relativePath)

def isFloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def isUserAdmin():
	try:
		if os.name == 'nt':
			return ctypes.windll.shell32.IsUserAnAdmin()
		else:
			return True
	except:
		return False

STYLESHEET = open(resourcePath("dependencies/custom.qss"), "r").read()

if isUserAdmin():
    if __name__ == '__main__':
        app = QApplication(sys.argv)
        app.setStyleSheet(STYLESHEET)
        ex = MainWindow()
        sys.exit(app.exec())
else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)