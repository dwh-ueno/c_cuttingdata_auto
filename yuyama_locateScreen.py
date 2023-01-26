# -*- coding: utf-8 -*-
"""
Created on Thu Jan 26 13:57:11 2023

@author: dwg-yuyama
"""

import streamlit as st
import openpyxl
import subprocess
import win32api
import os
import glob
import pyautogui as ag
import time
from PIL import Image



path = r"C:\label_picture\data_setting_ueno.png"

# img = Image.open(path)
# img.show()


time.sleep(2)
p = ag.locateOnScreen((path),
                      confidence=0.8
                      )
print(p)
x, y = ag.center(p)
ag.click(x, y)
