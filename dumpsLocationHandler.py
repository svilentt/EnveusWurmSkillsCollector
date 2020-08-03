import tkinter as tk
from tkinter import filedialog
import glob
import os

DUMPS_FILE = 'dumpsLocation.txt'

def isFolderSet():
	with open(DUMPS_FILE,'r') as dumpsLocation:
		return dumpsLocation.read(1)

def askForFolder():
	root = tk.Tk()
	root.withdraw()
	folder_path = filedialog.askdirectory(title = "Select your Wurm dumps directory")
	with open(DUMPS_FILE,'w') as dumpsLocation:
		dumpsLocation.write(folder_path)

def getLatestDumpFile():
	with open(DUMPS_FILE,'r') as dumpsLocation:
		dumpsFolder = dumpsLocation.readline()
		list_of_files = glob.glob(dumpsFolder + '/skills*.txt')
		latest_file = max(list_of_files, key=os.path.getctime)
		return latest_file.replace('\\', '/')