import os
import sys
import re
from time import strftime
import time
import datetime
import numpy
import pandas as pd
import msvcrt as m
import File_Management as fm
from tkinter import *
import webbrowser
from win32com.client import Dispatch

# print("Hello World")

#--------------------------------------------------------------------------------------------------
#--INPUT (WHAT DIRECTORIES WILL BE TRACKED?)
#--------------------------------------------------------------------------------------------------
#SET USER
usr = 'James'
#Young G

if usr == 'James':
	dirlistAll = ['//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST']
	highestDir = '//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST'
	dirlist = ['//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST']
elif usr == 'Young G':
	dirlistAll = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/GS Work', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/PBC', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/zArchive']
	highestDir = "//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY"
	dirlist = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY']
	#dirlist = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/GS Work', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/zArchive']
else:
	print('FAIL - DEFINE USER')
	GoblyGookgit 


#List directories for path where executable is located (FOR TESTING POINT TO PERSONAL)

i=0
for dirName, subdirList, fileList in os.walk(highestDir):
	i = i + 1
	#Retrieve only subdirs
	if i == 1:
		mdirs=subdirList
	print (dirName)
	print (subdirList)

for x in mdirs:
	print(x)

# dta = fm.operating(dirlist)
# fm.output(dta)

#--------------------------------------------------------------------------------------------------
#--DEFINE Classes
#--------------------------------------------------------------------------------------------------
class MENU_GUI(Frame):
	def __init__(self, master):
		Frame.__init__(self, master)
		self.master = master
		master.title("File Inventory Tool Menu")
		master.geometry("300x200")


		#Label interface
		self.label = Label(master, text="What process would you like to run?")
		self.label.pack()
		#User Options
		self.greet_button = Button(master, text="Motivate Me", command=self.greet)
		self.refresh_button = Button(master, text="Refresh Inventory", command=self.runinventory)
		self.new_button = Button(master, text="Settings", command=self.newinventory)
		self.greet_button.pack()
		self.refresh_button.pack()
		self.new_button.pack()
		
		#Menu 
		menubar = Menu(self.master)
		self.master.config(menu=menubar)
		fileMenu = Menu(menubar)
		menubar.add_command(label="View Inventory", command=self.viewfile)
		menubar.add_command(label="Exit", command=self.onExit)

	def greet(self):
		webbrowser.open_new_tab('hello.html')

	def onExit(self):
		self.quit()

	def newinventory(self):
		root = Tk()
		my_gui = MyFirstGUI(root)
		root.mainloop()

	def runinventory(self):
		dta = fm.operating(fm.dirlist)
		fm.output(dta)

	def viewfile(self):
		xl = Dispatch('Excel.Application')
		wb = xl.Workbooks.Open('//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST/BRG_FILE_INVENTORY.xlsx')
		xl.Visible = True    # optional: if you want to see the spreadsheet


class MyFirstGUI(Frame):
	def __init__(self, master):
		Frame.__init__(self, master)
		self.master = master
		master.title("A simple GUI")
		master.geometry("500x500")


		#Label interface
		self.label = Label(master, text="Please select files to include/exclude")
		# self.label.pack()
		#Button interface
		self.greet_button = Button(master, text="Greet", command=self.greet)
		# self.greet_button.pack()


		menubar = Menu(self.master)
		self.master.config(menu=menubar)
		fileMenu = Menu(menubar)
		menubar.add_command(label="Save Settings", command=self.onExit)
		menubar.add_command(label="Exit", command=self.onExit)

		r = 0
		for item in mdirs:
			r = r + 1
			self.dir_button=  Button(master, text=item, command=self.onExit).grid(row = r, column=2, sticky=W)
			self.dir_button = Checkbutton(master, variable=item, onvalue = 1, offvalue = 0, height=2, width = 20, anchor=W)
			self.dir_button.grid(row = r, column=1, sticky=W)
			self.dir_button.select()
			
		#Select all/unselect all buttom
		
	def greet(self):
		print("Greetings!")

	def onExit(self):
		self.quit()

