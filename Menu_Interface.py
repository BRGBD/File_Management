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

#FOR GREAT TUTORIAL OF THIS CONSTRUCT see https://www.youtube.com/watch?v=A0gaXfM1UN0&list=PLQVvvaa0QuDclKx-QpC9wntnURXVJqLyk&index=2


#--------------------------------------------------------------------------------------------------
#--DEFINE Classes
#--------------------------------------------------------------------------------------------------
class MENU_GUI(Frame):
	def __init__(self, master):
		Frame.__init__(self, master)
		self.master = master
		master.title("File Inventory Tool Menu")
		master.geometry("300x200")


        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others

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
		my_gui = Parentpage(root)
		root.mainloop()

	def runinventory(self):
		dta = fm.operating(fm.dirlist)
		fm.output(dta)

	def viewfile(self):
		xl = Dispatch('Excel.Application')
		wb = xl.Workbooks.Open('//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST/BRG_FILE_INVENTORY.xlsx')
		xl.Visible = True    # optional: if you want to see the spreadsheet



class Parentpage(Frame):
	def __init__(self, master):
		Frame.__init__(self, master)
		self.master = master
		master.title("A simple GUI")
		master.geometry("500x500")


		#Label interface
		self.label = Label(master, text="Please select files to include/exclude")

		menubar = Menu(self.master)
		self.master.config(menu=menubar)
		fileMenu = Menu(menubar)
		menubar.add_command(label="Save Settings", command=self.onExit)
		menubar.add_command(label="Exit", command=self.onExit)

		r = 0
		for item in mdirs:
			r = r + 1
			self.dir_name_button=[]
			self.dir_name_button.append(Button(master, text=item, textvariable=item, command = lambda item=item: self.drill(item)).grid(row = r, column=2, sticky=W))
			self.dir_check_button = Checkbutton(master, variable=highestDir + '/' + item, onvalue = 1, offvalue = 0, height=2, width = 20, anchor=W)
			self.dir_check_button.grid(row = r, column=1, sticky=W)
			self.dir_check_button.select()
			
		#Select all/unselect all buttom
		
	def greet(self):
		print("Greetings!")

	def drill(self,item):
		print(item)

	def onExit(self):
		self.quit()

class Childpage(Frame):
	def __init__(self, master, controller):
		Frame.__init__(self, master)
		self.master = master
		master.title("A simple GUI")
		master.geometry("500x500")


		#Label interface
		self.label = Label(master, text="Please select files to include/exclude")

		menubar = Menu(self.master)
		self.master.config(menu=menubar)
		fileMenu = Menu(menubar)
		menubar.add_command(label="Save Settings", command=self.onExit)
		menubar.add_command(label="Exit", command=self.onExit)

		r = 0
		for item in mdirs:
			r = r + 1
			self.dir_name_button=[]
			self.dir_name_button.append(Button(master, text=item, textvariable=item, command = lambda item=item: self.drill(item)).grid(row = r, column=2, sticky=W))
			self.dir_check_button = Checkbutton(master, variable=highestDir + '/' + item, onvalue = 1, offvalue = 0, height=2, width = 20, anchor=W)
			self.dir_check_button.grid(row = r, column=1, sticky=W)
			self.dir_check_button.select()
			
		#Select all/unselect all buttom
		
	def greet(self):
		print("Greetings!")

	def drill(self,item):
		print(item)

	def onExit(self):
		self.quit()

# Query User for which directories to exclude
root = tk.Tk()
my_gui = mi.MENU_GUI(root)
root.mainloop()


