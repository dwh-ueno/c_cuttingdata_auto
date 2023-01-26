import pyautogui as ag
import pyperclip
from time import sleep
import subprocess
import os

path_ueno = r"\\dwhnas1\DWH1\u_ueno\c_cuttingdata_auto"

def main(path):
    labelsoft_path = r"C:/Program Files (x86)/JustSystems/LMIGHTYB/LMIGHTYB.EXE"
    subprocess.Popen(labelsoft_path)

    sleep(3)
    
    ag.press("tab", presses=3)
    ag.press("Enter")
    sleep(0.5)

    ag.hotkey("alt", "n")
    pyperclip.copy(os.path.join(path_ueno, "s切断資料_雛形_保管場所\s切断資料 自動化テスト用.jlb"))
    ag.hotkey("ctrl", "v")
    ag.press("Enter")
    
    p = ag.locateOnScreen(os.path.join(path_ueno, "data_setting_ueno.png"))
    x, y = ag.center(p)
    ag.click(x, y)
    
    sleep(1)
    q = ag.locateOnScreen(os.path.join(path_ueno, "data_setting.png"))
    x, y = ag.center(q)
    print(x,y)
    ag.click(x, y)
    
    ag.hotkey("alt", "n")
    pyperclip.copy(path)
    ag.hotkey("ctrl", "v")
    ag.press("Enter")
    
    # ag.moveTo(0.3946*x, 0.0885*y, duration=1)
    # ag.click()

    # ag.moveTo(0.9158*x, 0.2565*y, duration=1)
    # ag.click()

    # ag.moveTo(0.06222*x, 0.1302*y, duration=1)
    # ag.doubleClick()

    # ag.press("r")
    # ag.press("t")
    # sleep(0.5)
    # ag.press("Enter")
    # sleep(0.5)
    # ag.press("Enter")

    # ag.moveTo(0.5542*x, 0.08854*y, duration=1)
    # ag.click()