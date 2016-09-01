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
	GoblyGook


#List directories for path where executable is located (FOR TESTING POINT TO PERSONAL)
mdirs=[]
i=0
for dirName, subdirList, fileList in os.walk(highestDir):
	i = i + 1
	#Retrieve only subdirs
	if i != 1:
		mdirs.append(dirName)


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
		self.greet_button = Button(master, text="Greet Me", command=self.greet)
		self.refresh_button = Button(master, text="Refresh Inventory", command=self.runinventory)
		self.new_button = Button(master, text="Generate New Inventory", command=self.runinventory)
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
		master.geometry("1500x800")


		#Label interface
		self.label = Label(master, text="Please select files to include/exclude")
		self.label.pack()
		#Button interface
		self.greet_button = Button(master, text="Greet", command=self.greet)


		menubar = Menu(self.master)
		self.master.config(menu=menubar)
		fileMenu = Menu(menubar)
		menubar.add_command(label="New Inventory", command=self.run)
		menubar.add_command(label="Refresh Inventory", command=self.onExit)
		menubar.add_command(label="Exit", command=self.onExit)


		for item in mdirs:
			self.listbox.insert(END, item)
		self.greet_button.pack()
		
	def greet(self):
		print("Greetings!")

	def onExit(self):
		self.quit()

#--------------------------------------------------------------------------------------------------
#--USER INTERACTION SECTION
#--------------------------------------------------------------------------------------------------
# Query User for which directories to exclude
root = Tk()
my_gui = MENU_GUI(root)
root.mainloop()



	
#Notes on this section:
	# lbind allows for single and double clicking if single clicke then x if double then y
	# Walk through the director?
	# In State - If they select it it greys it out and they can no longer accces


#person = input('Enter your name: ')
#print('Hello', person)

#exclude directories mentioned by User
exclude = ['PBC']
#wait() # code used to pause the screen until the user interacts with it
