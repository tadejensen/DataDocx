import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

import os
from ipydex import activate_ips_on_exception, IPS

import pandas as pd

from lengthy_imports import *

import gui_functions as gui_f
import gui_elements as gui_e
import ctypes
import warnings

projects_path = load_config().projects_path

activate_ips_on_exception()
# ignore performance warning to stop console cluttering
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

# correct resolution
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # if your windows version >= 8.1
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware() # win 8.0 or less 
    except:
        pass

myappid = '8p2.DataDocx' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


root = gui_e.Mainwindow()
projpath = filedialog.askdirectory(
                        parent=root,
                        initialdir=projects_path if projects_path else None)

try: os.chdir(projpath)
except OSError: pass # if projpath is '' (no directory selected)


project = gui_f.load_project()
if project is not None:
    root.set_project(project)
print(project)

confbtn = ttk.Button(root, text='Projekt konfigurieren', 
                     command=lambda: gui_e.Project_Setup_Window(root, root.project))
confbtn.pack(pady=5, padx=5, side='right')

IPS_btn = ttk.Button(root, text='Konsole betreten',
                     command=lambda: root.project.enter_console()\
                        if root.project is not None else IPS)
IPS_btn.pack(pady=5, padx=5, side='right')

getdata_btn = gui_e.GitButton(root, root.project, mode='pull')
getdata_btn.pack(side='left', padx=1, pady=5)

pushdata_btn = gui_e.GitButton(root, root.project, mode='push')
pushdata_btn.pack(side='left', padx=1, pady=5)

root.mainloop()




