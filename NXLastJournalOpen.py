#-------------------------------------------------------------------------------
# Name:        module2
# Purpose:
#
# Author:      Jeff Nuss
#
# Created:     18/11/2011
# Copyright:   (c) Jeff Nuss 2011
# Licence:     <your licence>
#-------------------------------------------------------------------------------
#!/usr/bin/env python

import os, sys, time, fnmatch, subprocess, win32com.client
from stat import *

currentTime = 0
os.chdir("C:\\Users\\User\\Desktop\\NXConnect_Journals")
doesFileExist = False

for currentFile in os.listdir("."):
    if fnmatch.fnmatch(currentFile, '*.cs'):
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
    os.startfile("C:\\Users\\User\\Desktop\\NXConnect_Journals\\" + mostRecentFile)