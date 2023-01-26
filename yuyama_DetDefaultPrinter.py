# -*- coding: utf-8 -*-
"""
Created on Thu Jan 26 14:57:24 2023

@author: dwg-yuyama
"""

import win32api
import win32print


print(win32print.GetDefaultPrinter())

win32print.SetDefaultPrinter("FUJIFILM Apeos C2570")

print(win32print.GetDefaultPrinter())