import threading
import time

class ThreadTest():
    def __init__(self):
        super().__init__()
        self.exitEvent = threading.Event()
        self.sleepThr = threading.Thread(target=self.startSleep)

    def startSleep(self):
        self.exitEvent.wait(10)

    def startSleepThread(self):
        self.sleepThr.start()
    
    def inputCommand(self):
        self.command = input()
        if self.command == "stop":
            self.exitEvent.set()

test = ThreadTest()
test.startSleepThread()
test.inputCommand()