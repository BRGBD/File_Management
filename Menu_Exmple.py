import os
import re
import tkinter as tk 
from tkinter import Tk, Frame, Menu  # python3
from tkinter import ttk
import webbrowser
import pickle
from win32com.client import Dispatch


TITLE_FONT = ("Helvetica", 18, "bold")




#--------------------------------------------------------------------------------------------------
#--INPUT (WHAT DIRECTORIES WILL BE TRACKED, FOR TESTING ONLY)
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

#--------------------------------------------------------------------------------------------------
#--INPUT (WHAT DIRECTORIES WILL BE TRACKED, FOR TESTING ONLY)
#--------------------------------------------------------------------------------------------------



#FOR GREAT TUTORIAL OF THIS CONSTRUCT see https://www.youtube.com/watch?v=A0gaXfM1UN0&list=PLQVvvaa0QuDclKx-QpC9wntnURXVJqLyk&index=2

parent_dir='//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST'
cdir = parent_dir
class SampleApp(tk.Tk):

	def __init__(self, *args, **kwargs):
		tk.Tk.__init__(self, *args, **kwargs)

		tk.Tk.iconbitmap(self, default="favicon.ico")

		#can think of *args as passing through variables
		#**kwargs as passing through dicts
		#init means (run every time class if called)

		# the container is where we'll stack a bunch of frames
		# on top of each other, then the one we want visible
		# will be raised above the others
		container = tk.Frame(self) #define our window = container
		container.pack(side="top", fill="both", expand=True) #Two options in tkinter for adding objects - Pack (just shove it in) - Grid (...a grid)
		container.grid_rowconfigure(0, weight=1)
		container.grid_columnconfigure(0, weight=1)

		self.frames = {} #dict
		for F in (StartPage, Settings, Parentpage, Childpage):
			page_name = F.__name__
			frame = F(parent=container, controller=self)
			self.frames[page_name] = frame

			# put all of the pages in the same location;
			# the one on the top of the stacking order
			# will be the one that is visible.
			frame.grid(row=0, column=0, sticky="nsew")

		self.show_frame("StartPage")

	def show_frame(self, page_name):
		'''Show a frame for the given page name'''
		frame = self.frames[page_name]
		frame.tkraise()


class StartPage(tk.Frame):

	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)
		self.controller = controller
		label = ttk.Label(self, text="This is the start page", font=TITLE_FONT)
		label.grid(row=0, column=1)

		greet_button = ttk.Button(self, text="Motivate Me", command=self.greet)
		refresh_button = ttk.Button(self, text="Refresh Inventory", command=self.runinventory)
		new_button = ttk.Button(self, text="Settings", command=lambda: controller.show_frame("Settings"))
		greet_button.grid(row=1, column=1)
		refresh_button.grid(row=2, column=1)
		new_button.grid(row=3, column=1)

		#Menu 
		#Define menu this will persist in all frames when switch
		menubar = Menu(self.controller)
		fileMenu = Menu(menubar)
		menubar.add_cascade(label="Menu", command=lambda: controller.show_frame("StartPage"))
		menubar.add_cascade(label="View Inventory", command=self.viewfile)
		menubar.add_cascade(label="Exit", command=self.onExit)
		self.controller.config(menu=menubar)

	
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
		xl.Visible = True	# optional: if you want to see the spreadsheet

class Settings(tk.Frame):

	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)
		self.controller = controller
		label = ttk.Label(self, text="This is settings page", font=TITLE_FONT)
		label.pack(side="top", fill="x", pady=10)


		button1 = ttk.Button(self, text="Define Folders to Inventory",
						   command = lambda: controller.show_frame("Parentpage"))
		button2 = ttk.Button(self, text="Other Options",
				   command=lambda: controller.show_frame("Childpage"))
		button1.pack()
		button2.pack()  


	def viewfile(self):
		xl = Dispatch('Excel.Application')
		wb = xl.Workbooks.Open('//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST/BRG_FILE_INVENTORY.xlsx')
		xl.Visible = True	# optional: if you want to see the spreadsheet

	def onExit(self):
		self.quit()

#NEED TO MAKE THIS DYNAMIC
class Parentpage(tk.Frame):
	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)
		self.controller = controller
		label = ttk.Label(self, text="This is page 1 - Parent", font=TITLE_FONT)
		label.grid(row = 1, column=0,padx=(50, 50), sticky=tk.EW)
		childbutton = ttk.Button(self, text="Step Back",
				   command=lambda: controller.show_frame("Childpage"))
	   


		self.fnames={}

		# This will run on the first iteration
		i=0
		for dirName, subdirList, fileList in os.walk(parent_dir):
				self.fnames[dirName] = subdirList

		self.parental = tk.StringVar(self)
		self.childy = tk.StringVar(self)

		self.parental.trace('w', self.update_options)

		# print('test')
		# ck = [re.sub(parent_dir,'',element) for element in self.fnames.keys()]
		# for x in ck:
		# 	print(x)


		self.parental_menu = tk.OptionMenu(self, self.parental, *self.fnames.keys())
		self.parental_menu.grid(row = 0, column=0,padx=(50, 50), sticky=tk.EW)

		self.dir_name_button=[]


		if self.dir_name_button:
			r = 0
			for btn in self.dir_check_button:
				print('enter button l0op')
				r = r + 1
				btn.grid(row = r, column=1, sticky=W)
				btn.select()
		

		# This will run on the first iteration
		i=0
		for dirName, subdirList, fileList in os.walk(parent_dir):
			i = i + 1
			#Retrieve only subdirs
			if i == 1:
				mdirs = subdirList

		self.childy_menu = tk.OptionMenu(self, self.childy, '')
		self.childy_menu.grid(row = 1, column=0,padx=(50, 50), sticky=tk.EW)


	def update_options(self, *args):
		selected_dir=self.fnames[self.parental.get()]
		# print (selected_dir)

		if selected_dir:
			self.childy.set(selected_dir[0])
			menu = self.childy_menu['menu']
			menu.delete(0, 'end')

			for fn in selected_dir:
				self.dir_name_button.append(ttk.Button(self, text=fn, command=lambda nation=fn: self.childy.set(nation)))
				print(self.dir_name_button)
				menu.add_command(label=fn, command=lambda nation=fn: self.childy.set(nation))
		else:
			self.childy.set('')
			menu = self.childy_menu['menu']
			menu.delete(0,50)


	def drill(self,item):
		child_dir=highestDir + '/' + item
		i=0
		for dirName, subdirList, fileList in os.walk(child_dir):
			i = i + 1
			#Retrieve only subdirs
			if i == 1:
				cdirs=subdirList
				pickle.dump(cdirs, open("cdirs.pkl", "wb")) #Save list of children
				print(subdirList)
		if cdirs:
			Parentpage.update(self)
			Childpage.update(self)
			self.controller.show_frame("Childpage")
		else:
			print('NO CHILD')
		
class Childpage(tk.Frame):

	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)
		self.controller = controller
		label = ttk.Label(self, text="This is page 2 - Child", font=TITLE_FONT)
		label.grid(row = 1, column=0,padx=(50, 50), sticky=tk.EW)
		childbutton = ttk.Button(self, text="Step Back",
				   command=lambda: controller.show_frame("Parentpage"))
		try:
			cdirs = pickle.load(open( "cdirs.pkl", "rb" )) # Load the list
			print(cdirs)
			print('Loaded pickle')
		except:
			cdirs = []
			print('Did not load')

		r = 5
		for item in cdirs:
			r = r + 1
			dir_name_button=[]
			dir_name_button.append(ttk.Button(self, text=item,  command = lambda item=item: self.drill(item)).grid(row = r, column=0,padx=(50, 50), sticky=tk.EW))
			dir_check_button = tk.Checkbutton(self, variable=highestDir + '/' + item, onvalue = 1, pady=10,offvalue = 0)
			dir_check_button.grid(row = r, column=0, sticky=tk.W)
			dir_check_button.select()

		childbutton.grid(row = r+2, column=0,padx=(50, 50), sticky=tk.EW)


	def drill(self,item):
		child_dir=highestDir + '/' + item
		i=0
		for dirName, subdirList, fileList in os.walk(highestDir):
			i = i + 1
			#Retrieve only subdirs
			if i == 1:
				global mdirs
				mdirs=subdirList
		lambda: controller.show_frame("Childpage")
		print(item)

if __name__ == "__main__":
	app = SampleApp()
	app.mainloop()



