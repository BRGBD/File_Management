import os
import re
from time import strftime
import time
import datetime
import pandas as pd
import msvcrt as m
from tkinter import *

# print("Hello World")

#--------------------------------------------------------------------------------------------------
#--INPUT (WHAT DIRECTORIES WILL BE TRACKED?)
#--------------------------------------------------------------------------------------------------

dirlistAll = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/GS Work', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/PBC', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/zArchive']
highestDir = "//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY"
dirlist = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY']
#dirlist = ['//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/GS Work', '//brg-DC-fs1.brg.local/HOME/gsteele/Python DIY/zArchive']
#--------------------------------------------------------------------------------------------------
#--DEFINE Classes
#--------------------------------------------------------------------------------------------------
class MyFirstGUI:
	def __init__(self, master):
		self.master = master
		master.title("A simple GUI")
		master.geometry("1500x800")
		self.label = Label(master, text="This is our first GUI!")
		self.label.pack()
		self.greet_button = Button(master, text="Greet", command=self.greet)
		self.greet_button.pack()
	def greet(self):
		print("Greetings!")

#--------------------------------------------------------------------------------------------------
#--DEFINE FUNCTIONS
#--------------------------------------------------------------------------------------------------


def maximum_id(rootDir):
# Find the highest ID assinged, if no records have been assigned IDs set the count to 0
	all_ids=[]
	for dirName, subdirList, fileList in os.walk(rootDir):
		subdirList[:] = [d for d in subdirList if d not in exclude] 
		print('Found directory: %s' % dirName)
		for fname in fileList:
			#print('\t%s' % fname)
			t=fname[:5]
			if re.match('BRG_\d',fname[:5]) is not None: #if there already exists a BRG_##### file
				n=str(fname[4:12])
				n=n.lstrip('0') #strip 0s to get highest number i.e. 0095 = 95
				all_ids.append(int(n))
	print(all_ids)
	if not all_ids:
		ct=0
	else:
		ct=max(all_ids)

	return(ct)

def wait():
    m.getch()

def get_file_info_refresh(dirName,fname):
		try:
			st = os.stat(dirName+"/"+fname)
		except IOError:
			print ("Failed to get information about" +  fname)
		else:
			print('\n')
			print('Refreshed file ' + fname)
			refresh={}
			#print(st)
			BRG_ID = (str(fname[0:12]))
			refresh[BRG_ID]={}
			refresh[BRG_ID]["Orig_Path"]=''
			refresh[BRG_ID]["Orig_Title"]=''
			refresh[BRG_ID]["BRG_Path"]=dirName+"/"+fname
			refresh[BRG_ID]["BRG_Title"]=fname
			refresh[BRG_ID]["File_Size"]=str(st.st_size)
			refresh[BRG_ID]["MR_Access"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_atime))
			refresh[BRG_ID]["MR_Mod"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_mtime))
			refresh[BRG_ID]["TO_Create"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_ctime))
			refresh[BRG_ID]["Refresh_DT"]=str(datetime.datetime.now())
		return(refresh)	

def get_file_info_original(dirName,fname):
		try:
			st = os.stat(dirName+"/"+desc)
		except IOError:
			print ("Failed to get information about" +  dirName+"/"+desc)
		else:
			orig={}
			BRG_ID = (str(desc[0:12]))
			orig[BRG_ID]={}
			orig[BRG_ID]["Orig_Path"]=dirName+"/"+fname
			orig[BRG_ID]["Orig_Title"]=fname
			orig[BRG_ID]["BRG_Path"]=dirName+"/"+desc
			orig[BRG_ID]["BRG_Title"]=desc
			orig[BRG_ID]["File_Size"]=str(st.st_size)
			orig[BRG_ID]["MR_Access"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_atime))
			orig[BRG_ID]["MR_Mod"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_mtime))
			orig[BRG_ID]["TO_Create"]=strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_ctime))
			orig[BRG_ID]["Refresh_DT"]=str(datetime.datetime.now())
			#could add "level" information here
		return(orig)


def rename_file(dirName,fname,ct): 
	#Rename File With BRG ID
	ids=str(zs)+str(ct)
	brg="BRG_"+ids[-8:]
	desc=brg+" "+fname
	print("\nAdded File " + desc)
	os.rename(dirName+"/"+fname,dirName+"/"+desc)
	return(desc)

def refresh_naming(rootDir): # no longer still used?
	for dirName, subdirList, fileList in os.walk(rootDir):
		subdirList[:] = [d for d in subdirList if d not in exclude] 
		print('Found directory: %s' % dirName)
		for fname in fileList:
			if re.match('BRG_\d',fname[:5]) is not None:
				reset=re.sub('BRG_\d\d\d\d\d\d\d\d ','',fname)
				os.rename(dirName+"/"+fname,dirName+"/"+reset)
def clean():
	for x in dirlistAll: # loops through directories listed, resets the naming
		rootDir=x
		os.chdir(rootDir)
		refresh_naming(x)

#--------------------------------------------------------------------------------------------------
#--USER INTERACTION SECTION
#--------------------------------------------------------------------------------------------------
# Query User for which directories to exclude

rootDirUser=dirlist[0]
os.chdir(rootDirUser)
for dirName, subdirList, fileList in os.walk(rootDirUser):
	for subs in subdirList:
			print(subs)

root = Tk()
my_gui = MyFirstGUI(root)
root.mainloop()



#person = input('Enter your name: ')
#print('Hello', person)

#exclude directories mentioned by User
exclude = ['PBC']
#wait() # code used to pause the screen until the user interacts with it


#--------------------------------------------------------------------------------------------------
#--MASTER OPERATING PEICE
#--------------------------------------------------------------------------------------------------
# Daily refresh section
dta = []
for x in dirlist: # loops through directories listed
	rootDir=x
	os.chdir(rootDir)
	print(os.listdir("."))
	ct=maximum_id(rootDir) #ct is highest numbered BRG #### file i.e. if BRG 001 and BRG 002 in sub, then its 2
	for dirName, subdirList, fileList in os.walk(rootDir):
		subdirList[:] = [d for d in subdirList if d not in exclude] 
		print('Found directory: %s' % dirName.replace('\\','/'))
		for fname in fileList:
			if re.match('BRG_\d',fname[:5]) is not None:    #if there already exists a BRG_##### file
				# Already Been Indexed
				print(fname + " : Accounted For In System")
				dta.append(get_file_info_refresh(dirName.replace('\\','/'),fname)) #get info refresh returns all the info of a particular file; this step is concating all the file info into dta.
			else: # if there isnt a BRG 0000 name, then find out the highest count in the folder, add 1, rename the current file now have it.
				# Need To Add
				print(dirName.replace('\\','/')) 
				zs='00000000'
				ct=ct+1
				desc=rename_file(dirName.replace('\\','/'),fname,ct)
				dta.append(get_file_info_original(dirName.replace('\\','/'),fname))

# # # Run if Need to Reset All Files During Testing

#--------------------------------------------------------------------------------------------------
#--OUTPUT
#--------------------------------------------------------------------------------------------------
print("\t" + "\t" + "\t"+ "new stuff -----------------------------------------------------")


# Create data frame in pandas should simplify

frames = []
file_ids=[]
for y in dta:
	for brgid, info in y.items():
		file_ids.append(brgid)
		frames.append(pd.DataFrame.from_dict(y, orient='index'))
# Create data frame in pandas should simplify

dfFiles = pd.concat(frames)
pd.options.display.max_columns = 500

print(dfFiles.head(20))
#if os.path.exists(rootDir+'/BRG_00000001 FileSummary3.xlsx'): #still needs to account for BRG_000003, i.e. if File Summary isn't always the top file
 #   os.remove(rootDir+'/BRG_00000001 FileSummary3.xlsx')

highest_file = 0
highest_file_string = ""
print("working")
for x in dirlist: # loops through directories listed; removes any old file summaries
	rootDir=x
	os.chdir(rootDir)
	for dirName, subdirList, fileList in os.walk(rootDir):
		subdirList[:] = [d for d in subdirList if d not in exclude] 
		for fname in fileList:
			if re.match('.*FileSummary3.*',fname) is not None:    #if there already exists a BRG_##### file
				print(fname)
				print(subdirList)
				print(fileList)
				print(dirName)
				fileNumber = fname[4:12]
				removeZero = fileNumber.replace("0","") 
				if int(removeZero) > highest_file:
					highest_file_string = fileNumber
				os.remove(dirName+"/"+fname)
print(highest_file_string)
dfFiles.to_excel(highestDir+'/BRG_'+highest_file_string+' FileSummary3.xlsx', index=True, sheet_name = "File Summary") # as of now will always send it to the last folder location -> easy to change later though

print("directory work")


for dirName, subdirList, fileList in os.walk(rootDir):
	subdirList[:] = [d for d in subdirList if d not in exclude] 
	for fname in fileList:
			print(fname)

#clean()