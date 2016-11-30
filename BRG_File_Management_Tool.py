import os
import re
import tkinter as tk 
from tkinter import Tk, Frame, Menu  # python3
from tkinter import ttk
import webbrowser
import pickle
from win32com.client import Dispatch
from tkinter import filedialog

TITLE_FONT = ("Calibri", 18, "bold")




#--------------------------------------------------------------------------------------------------
#--INPUT (WHAT DIRECTORIES WILL BE TRACKED, FOR TESTING ONLY)
#--------------------------------------------------------------------------------------------------

#FOR GREAT TUTORIAL OF THIS CONSTRUCT see https://www.youtube.com/watch?v=A0gaXfM1UN0&list=PLQVvvaa0QuDclKx-QpC9wntnURXVJqLyk&index=2

# READ IN PREVIOUSLY SET FILENAME
with open('dir_settings.txt', 'r') as f:
    first_line = f.readline()

global program_parent
program_parent = first_line
cdir = os.path.dirname(os.path.realpath(__file__))


global current_parent
current_parent = program_parent



class SampleApp(tk.Tk):

  def __init__(self, *args, **kwargs):
    tk.Tk.__init__(self, *args, **kwargs)

    tk.Tk.iconbitmap(self, default="brg_logo.ico")

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
    label = ttk.Label(self, text="FILE MANAGMENT TOOL: MAIN MENU", font=TITLE_FONT)
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
    xl.Visible = True # optional: if you want to see the spreadsheet

class Settings(tk.Frame):

  def __init__(self, parent, controller):
    tk.Frame.__init__(self, parent)
    self.controller = controller
    label = ttk.Label(self, text="FMT: SETTINGS PAGE", font=TITLE_FONT)
    label.pack(side="top", fill="x", pady=10)


    button0 = ttk.Button(self, text="Set Parent Directory",
           command = self.setdir)

    button1 = ttk.Button(self, text="Define Folders to Inventory",
               command = lambda: controller.show_frame("Parentpage"))

    button2 = ttk.Button(self, text="Other Options",
           command=lambda: controller.show_frame("Childpage"))

    button0.pack()
    button1.pack()
    button2.pack()


  def setdir(self):
    dir_opt = {}
    dir_opt['initialdir'] = cdir
    dir_opt['mustexist'] = False
    dir_opt['title'] = 'Please select directory'
    filename  = filedialog.askdirectory(**dir_opt)

    if filename:
      # Write filename to file, to be stored outside of application/semi-perm....
      f = open('dir_settings.txt', 'w')
      f.write(filename)
      f.close()
      print (filename)

  def viewfile(self):
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open('//brg-DC-fs1.brg.local/HOME/jlewis/Python_TEST/BRG_FILE_INVENTORY.xlsx')
    xl.Visible = True # optional: if you want to see the spreadsheet

  def onExit(self):
    self.quit()

#NEED TO MAKE THIS DYNAMIC
class Parentpage(tk.Frame):
  def __init__(self, parent, controller):
    tk.Frame.__init__(self, parent)
    self.controller = controller


    # Initial to run during init...not called again
    dirdisp = tk.StringVar(self, value = program_parent)
    i=0
    for dirName, subdirList, fileList in os.walk(program_parent): # Show Parent
        i = i + 1
        fnames = subdirList
        if i == 1:
          break


    # Refresh list of filenames to be tagged on top of buttons
    def refresh_filenames(target):
      i=0
      for dirName, subdirList, fileList in os.walk(target):
          i = i + 1
          rnames = subdirList
          if i == 1:
            break
      return(rnames)


    # Take one step down, refresh buttons
    def drill_down(j):
      print('Parent Prior to Update', current_parent)
      refreshdir = (current_parent + '/' + j)
      refresh_names = refresh_filenames(refreshdir)

      if refresh_names:
        dirdisp.set(current_parent + '/' + j)

        for child in self.winfo_children():
          child.destroy()
        gen_buttons(refresh_names)

        global prev_parent
        prev_parent = current_parent

        global current_parent
        current_parent = current_parent + '/' + j
        print('Parent PST Update', current_parent)

    # NEED TO UPDATE
    def step_out():
      print('Parent Prior to Update', prev_parent)
      refreshdir = program_parent.replace(prev_parent,'')
      refresh_names = refresh_filenames(refreshdir)

      if refresh_names:
        dirdisp.set(refreshdir)

        for child in self.winfo_children():
          child.destroy()
        gen_buttons(refresh_names)

        global current_parent
        current_parent = refreshdir
        print('Parent PST Update', refreshdir)

    #Return to parent folder structure, refresh buttons
    def parent_return():
      print('Parent Prior to Update', current_parent)
      dirdisp.set(program_parent)
      refresh_names = refresh_filenames(program_parent)

      if refresh_names:
        for child in self.winfo_children():
          child.destroy()
        gen_buttons(refresh_names)

        global current_parent
        current_parent = program_parent
        print('Parent PST Update', program_parent)


    #Generate all buttons on Frame ---- Call once during intit, wipe and refresh when click
    def gen_buttons(fnames):
      # Title
      label = ttk.Label(self, textvariable=dirdisp, font=TITLE_FONT)
      label.grid(row = 0, column=0,padx=(50, 50), sticky=tk.EW)

      # Menu Buttons
      ppbutton = ttk.Button(self, text="Program Parent Directory", command=parent_return)
      ppbutton.grid(row = 20, column=0,padx=(50, 50), sticky=tk.EW)

      outbutton = ttk.Button(self, text="Step Back", command=step_out)
      outbutton.grid(row = 21, column=0,padx=(50, 50), sticky=tk.EW)
      #Directory Buttons
      rw = 1

      for x in fnames:
        rw  = rw + 1
        vnameb = 'button_' + str(x).replace(' ','_')
        vnamec = 'button_' + str(x).replace(' ','_')
        vnamec = ttk.Checkbutton(self, variable= str(x).replace(' ','_'))
        vnameb = ttk.Button(self, text=x, command = lambda j = x: drill_down(j) )
        vnamec.grid(row=rw,column=1,padx=(1, 1), sticky=tk.W)
        vnameb.grid(row=rw,column=0,sticky=tk.EW)
    gen_buttons(fnames)
        

    
class Childpage(tk.Frame):

  def __init__(self, parent, controller):
    tk.Frame.__init__(self, parent)
    self.controller = controller
    label = ttk.Label(self, text="This is page 2 - Child - FILLER", font=TITLE_FONT)
    label.grid(row = 1, column=0,padx=(50, 50), sticky=tk.EW)
    childbutton = ttk.Button(self, text="Step Back",
           command=lambda: controller.show_frame("Parentpage"))
    childbutton.grid(row = 2, column=0,padx=(50, 50), sticky=tk.EW)

if __name__ == "__main__":
  app = SampleApp()
  app.mainloop()



