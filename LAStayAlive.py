from time import sleep
import pyautogui
import random
import keyboard
import time
from os import system
import win32gui, win32com.client

screenWidth, screenHeight = pyautogui.size()
random.seed()
startedflag = False
clickCounter = 0
startTime = 0
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)
        
    def setwindowhandle(self,hwnd):
        self._handle = hwnd[0]

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)
def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst


def FindLostArkWindowAndFocus():
    w = WindowMgr()
    appwindows = get_app_list()
    for i in appwindows:
        result = i[1].find('LOST ARK')
        if result != -1:
            w.setwindowhandle(i)
            w.set_foreground()
            return
def FocusConsoleOnClose():
    w = WindowMgr()
    appwindows = get_app_list()
    for i in appwindows:
        result = i[1].find('LAStay')
        if result != -1:
            w.setwindowhandle(i)
            w.set_foreground()
            return

def GenerateRandomInts():
    moveX = random.randint((int)(screenWidth/2 - 150),(int)(screenWidth/2 + 150))
    moveY = random.randint((int)(screenHeight/2 - 100),(int)(screenHeight/2 + 100))
    return moveX,moveY
def GetRandomTimeIntervall():
    return random.randint(1,5)
    
def CloseScript():
    print("test")
    
def EntryPoint():
    input("Start Program? \nPress any key to continue." )
    system('cls')
def GetTime():
    return time.time()
def PrintInfo():
    system('cls')
    print("Press 'Esc' to terminate the Program")
    elapsedTime = GetTime() - startTime
    print(f'Clicks: {clickCounter} || Time since beginning: {"%02d" %int((elapsedTime/3600)%24)}:{"%02d" %int((elapsedTime/60)%60)}:{"%02d" %(int)(elapsedTime%60)}')

while 1:
    if not startedflag:
        EntryPoint()    
        startedflag = True
        startTime = time.time()
        FindLostArkWindowAndFocus()
    x,y = GenerateRandomInts()
    pyautogui.moveTo(x,y)
    for i in range(GetRandomTimeIntervall()*100):
        sleep(1/100)
        if i % 60 == 0:
            PrintInfo()
        if keyboard.is_pressed('esc'):
            FocusConsoleOnClose()
            print("\nProgram Exited!")
            sleep(2)
            exit()
    FindLostArkWindowAndFocus() #Some of my friends really pulled of the impossible and Lost Ark had no focus or lost it during the AFK-time so calling the Focus before clicking...
    sleep(0.1)
    pyautogui.click(button='right')
    pyautogui.click() #clicking left and right mousebutton because there is movement option for both
    clickCounter += 1
    