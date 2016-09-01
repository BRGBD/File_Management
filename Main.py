import os
import sys
import re
from time import strftime
import time
import datetime
import numpy
import pandas as pd
import msvcrt as m
import Menu_Interface as mi
from tkinter import *
import webbrowser
from win32com.client import Dispatch

#--------------------------------------------------------------------------------------------------
#--MAIN
#--------------------------------------------------------------------------------------------------
# Query User for which directories to exclude
root = Tk()
my_gui = mi.MENU_GUI(root)
root.mainloop()

