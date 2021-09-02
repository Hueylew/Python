#!/usr/bin/env python3

import pyautogui
import webbrowser
import os

os.system("open /Applications/Safari.app http://192.168.0.85:8181/tos/")

pyautogui.doubleClick(x = 36, y = 197 )
pyautogui.PAUSE = 1
pyautogui.write(['Test','enter'])