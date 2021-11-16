#          Copyright Joe Coder 2004 - 2006.
# Distributed under the Boost Software License, Version 1.0.
#    (See accompanying file LICENSE_1_0.txt or copy at
#          https://www.boost.org/LICENSE_1_0.txt)

import win32gui, win32api, win32com.client, win32con
from PIL import ImageGrab, ImageTk
import PIL.Image # used this as a workaround for not finding images in thumbnails folder
from tkinter import *
from tkinter import ttk
from pywinauto import Desktop
import os
from functools import partial

def find_window_names():
    # thumbnails dir
    script = os.path.realpath(__file__)
    dir = script.split('format_windows_thirds.py')[0] + 'thumbnails\\'

    # create folder for thumbanil images if not already there
    if os.path.isdir(dir) == False:
        os.mkdir(dir)

    # clear previous thumbnails
    for f in os.listdir(dir):
        os.remove(os.path.join(dir, f))

    # thumbnail size
    size = 100, 100

    # list of handles
    handles = []

    # find windows
    windows = Desktop(backend="uia").windows()

    # for each window
    for w in windows:
        # making sure it is a usable window
        if w.window_text() != '' and w.window_text() != 'Taskbar' and w.window_text() != 'Program Manager':

            # gets window name
            window_name = w.window_text()

            # gets handle from window name
            handle = win32gui.FindWindow(0, window_name)

            # tries until success with max tries of 25 to set window to foreground using handle
            # workaround for pywintypes.error: (0, 'SetForegroundWindow', 'No error message is available')
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            result = 0
            attempts = 0
            while result == 0 and attempts < 25:
                try:
                    win32gui.SetForegroundWindow(handle)
                    result = 1
                except:
                    attempts += 1
                    print('Window could not be set to foreground')

            # get thumbnail image
            #win32gui.ShowWindow(handle, win32con.SW_NORMAL)
            bbox = win32gui.GetWindowRect(handle)
            img = ImageGrab.grab(bbox)
            img = img.resize(size)
            img.save(dir + str(handle) + '.jpeg')
            # keep orginal handles from windows
            handles.append(handle)

    # creates 3 guis in total but one at a time
    num = 0
    while(num<3):
        create_gui(handles, num)
        num += 1

def snap_window(ph, handle, num, w, h, root):
    # set window to foregorund for moving
    win32gui.SetForegroundWindow(handle)

    # num is for first window, second window, third window
    if num == 0:
        # move to left third of monitor
        win32gui.ShowWindow(handle, win32con.SW_NORMAL)
        win32gui.MoveWindow(handle, 0, 0, int(w/3), h-40, True)
    if num == 1:
        # move to middle third of monitor
        win32gui.ShowWindow(handle, win32con.SW_NORMAL)
        win32gui.MoveWindow(handle, int(w/3), 0, int(w/3), h-40, True)
    if num == 2:
        # move to right third of monitor
        win32gui.ShowWindow(handle, win32con.SW_NORMAL)
        win32gui.MoveWindow(handle, int((w/3)*2), 0, int(w/3), h-40, True)

    # closes out of tkinter gui
    root.destroy()

def create_gui(handles, num):
    # find resolution of monitor 0=w, 1=h
    w = win32api.GetSystemMetrics(0)
    h = win32api.GetSystemMetrics(1)

    root = Tk()
    # make the gui top window
    root.attributes('-topmost', True)
    frm = ttk.Frame(root)
    frm.grid()
    root.configure(background='gray')
    root.overrideredirect(True)

    # move gui based on where the window will be snapped
    if num == 0:
        gui_w = int(w/6 - w/12)
        gui_h = int(h/3)
        root.geometry('+'+str(gui_w)+'+'+str(gui_h))
    if num == 1:
        gui_w = int(w/6 + w/3 - w/12)
        gui_h = int(h/3)
        root.geometry('+'+str(gui_w)+'+'+str(gui_h))
    if num == 2:
        gui_w = int(w/6 + w/3 + w/3 - w/12)
        gui_h = int(h/3)
        root.geometry('+'+str(gui_w)+'+'+str(gui_h))


    # for each window create an image button with thumbnail and a label with window name
    count = 0
    row = 1
    script = os.path.realpath(__file__)
    dir = script.split('format_windows_thirds.py')[0] + 'thumbnails\\'
    for photo in os.listdir(dir):
        if count > 2:
            row += 2
            count = 0
        im = PIL.Image.open(dir + photo)
        ph = ImageTk.PhotoImage(im)
        photo_handle = photo.split('.jpeg')[0]
        for i in range(0, len(handles)):
            if(photo_handle in str(handles[i])):
                handle = handles[i]
        window_name = win32gui.GetWindowText(handle)
        if len(window_name) > 20:
            window_name = window_name[0:20] + '...'
        Button(root, image = ph, command=partial(snap_window, ph, handle, num, w, h, root)).grid(column=count, row=row, padx=30, pady=5)
        Label(root, text=window_name, bg='gray', fg='white', font='Helvetica 10 bold').grid(column=count, row=row+1)
        count += 1
    root.mainloop()

find_window_names()