#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Jeff Nuss
#
# Created:     03/11/2011
# Copyright:   (c) Jeff Nuss 2011
# Licence:     <your licence>
#-------------------------------------------------------------------------------
#!/usr/bin/env python

import os, sys, time, fnmatch, subprocess, win32com.client
from stat import *

currentTime = 0
os.chdir("C:\\Users\\User\\AppData\\Local\\Temp\\")
doesFileExist = False

for currentFile in os.listdir("."):
    if fnmatch.fnmatch(currentFile, '*.syslog'):
        doesFileExist = True
        if os.path.getmtime(currentFile) > currentTime:
            mostRecentFile = currentFile
            currentTime = os.path.getctime(currentFile)

if not doesFileExist:
    from _tkinter import *
    import tkinter
    root = tkinter.Tk()
    w = tkinter.Label(root, text="File not found")
    w.pack()
    root.mainloop()
else:
    os.startfile("C:\\Users\\User\\AppData\\Local\\Temp\\" + mostRecentFile)
    shell = win32com.client.Dispatch("WScript.Shell")
    while not shell.AppActivate("Notepad"):
        pass
    shell.SendKeys("^{END}")
exit