import time
import pyautogui as pgui

class RpaDelisis():
    def dl_itemsheet(self):
        pgui.hotkey('ctrl', 's')
        time.sleep(1)
        for i in range(6):
            time.sleep(1)
            pgui.press('tab')
        pgui.press('enter')
        pgui.write(r'C:\Users\LH21475\Desktop')
        time.sleep(1)
        for i in range(3):
            pgui.hotkey('shift', 'tab')
            time.sleep(1)
        time.sleep(1)
        pgui.press('enter')
        time.sleep(10)
