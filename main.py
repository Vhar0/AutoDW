import os
import time

from mss import mss
import mss.tools
import ctypes
from ctypes import wintypes
from WinPos import get_window_hwnd, get_window_pos_from_hwnd
import cv2
import numpy as np
import win32api, win32con, win32gui
import shutil
import pyautogui

"""
pyautogui.click(100, 100)
pyautogui.moveTo(100, 150)
pyautogui.moveRel(0, 10)  # move mouse 10 pixels down
pyautogui.dragTo(100, 150)
pyautogui.dragRel(0, 10)  # drag mouse 10 pixels down
"""

# todo: recognize sidebar for scroll.
#       find binding box of the app.
#       improve error handling (when the download button is not found).

user32 = ctypes.WinDLL('user32', use_last_error=True)


# Define RECT structure
class RECT(ctypes.Structure):
    _fields_ = [('left', ctypes.c_long),
                ('top', ctypes.c_long),
                ('right', ctypes.c_long),
                ('bottom', ctypes.c_long)]

# delimito il rettangolo della finestra
def get_window_rect(hwnd):
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def click(x, y):

    # pos is not usable because it refers to the image
    x += 10
    y += 10
    pos = (x, y)
    # win32api.SetCursorPos(pos)
    print("mouse grabbed to: " + str(pos))
    cx, cy = win32gui.GetCursorPos()
    win32api.SetCursorPos(pos)
    # time.sleep(1)
    pyautogui.click(x, y, 1, 2, "left", 1.0)
    # time.sleep(1)
    win32api.SetCursorPos((cx, cy))


# function that tries to move the window if the button is covered
def skip(hwnd):
    # save the mouse coordinates to leter restore them
    cx, cy = win32gui.GetCursorPos()
    print(hwnd)
    x = hwnd[0]
    y = hwnd[1]
    win32api.SetCursorPos((x, y))
    # time.sleep(10)
    pyautogui.click((x, y))
    # time.sleep(10)
    win32api.SetCursorPos((cx, cy))


def scroll():
    img = cv2.imread("img.png")
    template = cv2.imread("cattura2.png", 0)
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Applica il template matching
    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    threshold = 0.8
    loc = np.where(res >= threshold)

    # Disegno dei rettangoli dove il template è stato trovato
    for pt in zip(*loc[::-1]):  # (*loc[::-1]) inverte loc
        cv2.rectangle(img, pt, (pt[0] + template.shape[1], pt[1] + template.shape[0]), (0, 255, 0), 2)

    # Salva l'immagine invece di visualizzarla
    cv2.imwrite("output2.png", img)


# get the window location
hwnd = get_window_hwnd('Browser Window')


# BoundingRectangle	[l=599,t=420,r=1447,b=1031] width=848px height=611px

# a[start:stop]   items start through stop-1
# a[start:]       items start through the rest of the array
# a[:stop]        items from the beginning through stop-1
# a[:]            a copy of the whole array

pos = get_window_rect(hwnd)

print(pos)

# print(cv2.__version__)

with mss.mss() as sct:
    # The screen part to capture
    monitor = {"top": pos[1], "left": pos[0], "width": 848, "height": 848}
    output = "img.png".format(**monitor)

    # Grab the data
    sct_img = sct.grab(monitor)

    # Save to the picture file
    mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
    print(output)

# analisi per la barra laterale
scroll()


# move the output img to the output sub-dir
# shutil.move("img.png", "output")

# --------------- analisi per il pulsante download --------------- #

# Carica l'immagine principale e il template
img = cv2.imread("img.png")
template = cv2.imread("cattura.png", 0)
img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# Applica il template matching
res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
threshold = 0.8
loc = np.where(res >= threshold)


# Disegno dei rettangoli dove il template è stato trovato
for pt in zip(*loc[::-1]):  # (*loc[::-1]) inverte loc
    cv2.rectangle(img, pt, (pt[0] + template.shape[1], pt[1] + template.shape[0]), (0, 255, 0), 2)

# Salva l'immagine invece di visualizzarla
cv2.imwrite("output.png", img)


point = str(loc[0]) + str(loc[1])

print(point)

try:
    print(pt)

except NameError:
    print("Variable 'x' is not defined")
    # function to find the button if is covered
    skip(pos)

else:
    x = pos[0] + pt[0]
    y = pos[1] + pt[1]

    # print(type(point))

    click(x, y)
    os.remove("img.png")
