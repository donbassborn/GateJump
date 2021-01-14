from __future__ import division 
from PIL import Image

import win32gui
import win32com
import win32ui
import win32con
import time
import random
import mouse


import numpy as np                # Numerical Python 
import matplotlib.pyplot as plt   # Python plotting


def EnumWindows_Callback(hwnd, hwnds):
    if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
        if("EVE " in win32gui.GetWindowText(hwnd)):
            hwnds.append(hwnd)

def Click_Callback():
    global click_cords
    click_cords = mouse.get_position()
    print("Jump button at", click_cords)


client_hwnds = []
win32gui.EnumWindows(EnumWindows_Callback, client_hwnds)

print("Found " + str(len(client_hwnds)) + " EVE clients.")
pos = 1
for client_hwnd in client_hwnds:
    print(str(pos) + ". " + win32gui.GetWindowText(client_hwnd))
    pos += 1

active_hwnd = 0
print("Enter the client number:")
while 1:

    
    #

    count = input()
    try:
        count = int(count)

        if(count > 0 and count < len(client_hwnds)+1):
            active_hwnd = client_hwnds[count-1]
            print("Selected client: " + win32gui.GetWindowText(active_hwnd))
        else:
            print("Enter values between 1 and " + str(len(client_hwnds)))
            continue
        break
    except ValueError:
        print("Enter the client number:")
left, top, right, bottom = win32gui.GetWindowRect(active_hwnd)
w = right-left
h = bottom-top
print("Resolution:", str(w) + 'x' + str(h))

print("Switch to watched window in 3...")
time.sleep(1)
print("Switch to watched window in 2...")
time.sleep(1)
print("Switch to watched window in 1...")
time.sleep(1)

print("Click the center of Jump button:")

click_cords = (0,0)
mouse.on_click(Click_Callback)
mouse.wait(target_types=mouse.UP)
mouse.unhook_all()


hwndDC = win32gui.GetWindowDC(active_hwnd)
mfcDC = win32ui.CreateDCFromHandle(hwndDC)
saveDC = mfcDC.CreateCompatibleDC()
skip_first = True
prev_line = object()


while 1:
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, right-left, bottom-top)
    saveDC.SelectObject(saveBitMap)

    saveDC.BitBlt((0, 0), (right-left, bottom-top), mfcDC, (0, 0), win32con.SRCCOPY)
    pre_bitmap = saveBitMap.GetBitmapBits(True)
    bmp = Image.frombytes('RGB', (w, h), pre_bitmap, 'raw', 'BGRX')

    

    line = bmp.crop((75,85,250,86)).load()

    if(skip_first):
        prev_line = line
        skip_first = False
        continue
    
    diff_count = 0
    for i in range(250-75):
        diff = 0
        diff += abs(line[i,0][0] - prev_line[i,0][0])
        diff += abs(line[i,0][1] - prev_line[i,0][1])
        diff += abs(line[i,0][2] - prev_line[i,0][2])
        if(diff > 20):
            diff_count += 1

    if(diff_count > 20):
        print("Jump...")

        prev_hwnd = win32gui.GetForegroundWindow()
        prev_mode = win32con.SW_RESTORE

        if win32gui.GetWindowPlacement(prev_hwnd)[1] == win32con.SW_SHOWMAXIMIZED:
            prev_mode = win32con.SW_MAXIMIZE

        time.sleep(5+abs(random.normalvariate(2,1)))

        if active_hwnd != prev_hwnd:
            win32gui.ShowWindow(active_hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(active_hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0,0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )
            win32gui.SetWindowPos(active_hwnd,win32con.HWND_TOPMOST, 0, 0, 0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )
            win32gui.SetWindowPos(active_hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0,0,win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
            #~ win32gui.ShowWindow(hWnd,win32con.SW_SHOW); 
            #~ win32gui.BringWindowToTop(hWnd);
            #~ win32gui.SetForegroundWindow(hWnd);
            #http://stackoverflow.com/questions/6312627/windows-7-how-to-bring-a-window-to-the-front-no-matter-what-other-window-has-fo
            #http://timgolden.me.uk/pywin32-docs/win32gui__SetWindowPos_meth.html
            #~ //-- on Windows 7, this workaround brings window to top
            #~ win32gui.SetWindowPos(hWnd,HWND_NOTOPMOST,0,0,0,0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE);
            #~ win32gui.SetWindowPos(hWnd,HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE);
            #~ win32gui.SetWindowPos(hWnd,HWND_NOTOPMOST,0,0,0,0,win32con.SWP_SHOWWINDOW | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE);

        time.sleep(0.5 + abs(random.normalvariate(0.2,0.1)))
        mouse.move(click_cords[0]+random.normalvariate(0,4), click_cords[1]+random.normalvariate(0,4), True, duration=abs(random.normalvariate(0.5,0.1)))
        time.sleep(0.5 + abs(random.normalvariate(0,2.1)))
        mouse.click()

        if prev_hwnd != active_hwnd:
            win32gui.ShowWindow(prev_hwnd, prev_mode)
            win32gui.SetActiveWindow(prev_hwnd)
            win32gui.SetWindowPos(prev_hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )
            win32gui.SetWindowPos(prev_hwnd,win32con.HWND_TOPMOST,   0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )
            win32gui.SetWindowPos(prev_hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE) 
            

        time.sleep(8)
        skip_first = True

    prev_line = line


    
    time.sleep(1)


win32gui.DeleteObject(saveBitMap.GetHandle())
saveDC.DeleteDC()
mfcDC.DeleteDC()
win32gui.ReleaseDC(active_hwnd, hwndDC)