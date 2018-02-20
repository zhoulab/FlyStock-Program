# This is a GUI program meant to print stock labels and query stocks
# For help ask Alain Garcia (alaingarcia@ufl.edu)

# For GUI functionality
from tkinter import Tk, ttk

import time # Only used to access current day
import zipfile # To unzip .odb and extract data from database
import os # Delete files
import shutil # For copy functions
from pathlib import Path # To get home directory (to later get dropbox directory)

# Import adder.py and finder.py
import sys
sys.path.append('Resources/')
import adder
import finder
import connect

# Python file path C:/.../LabelPrinter/
dir = os.path.abspath(os.path.dirname(__file__))+"\Resources"

# LabDB path
LabDBodb = str(Path.home()) + '\Dropbox (ZhouLab)\ZhouLab Team Folder\General\Databases\LabDB.odb'

# Delete previous mydb files (used during renameFiles)
def deletePrevious():
    for root, dirs, files in os.walk('Resources\LabDB', topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    if os.path.isdir('Resources\LabDB'):
        os.rmdir('Resources\LabDB')
    else:
        pass


    for root, dirs, files in os.walk('Resources\LabDBWrite', topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    if os.path.isdir('Resources\LabDBWrite'):
        os.rmdir('Resources\LabDBWrite')
    else:
        pass

# Function to rename database files
def renameFiles(readOrWrite):
    if(readOrWrite == "r"):
        files = os.listdir('Resources\LabDB\database')

        #Then, rename and 'create' new mydb
        for filename in files:
            os.rename(os.path.join('Resources\LabDB\database', filename), os.path.join('Resources\LabDB\database', "mydb."+filename))
    elif(readOrWrite == "w"):
        files = os.listdir('Resources\LabDBWrite\database')

        #Then, rename and 'create' new mydb
        for filename in files:
            os.rename(os.path.join('Resources\LabDBWrite\database', filename), os.path.join('Resources\LabDBWrite\database', filename.replace("mydb.","")))

#Uses zipfile library to unzip and extract HSQLDB database from .odb
def openFile():
    unzip = zipfile.ZipFile(LabDBodb, 'r')
    if os.path.isdir('Resources\LabDB'):
        pass
    else:
        os.makedirs('Resources\LabDB')
    unzip.extractall('Resources\LabDB')
    unzip.close()

# Copies LabDB, removes mydb. prefix, and removes prefix
# Then proceeds to zip the file up and then rename to .odb
def writeFile():

    if os.path.isfile('Resources\Lab.odb'):
        os.remove('Resources\Lab.odb')

    #Populate LabDBWrite
    shutil.copytree('Resources\LabDB','Resources\LabDBWrite')
    os.remove('Resources\LabDBWrite\database\mydb.lck')

    #Remove mydb. prefix for a write ("w")
    renameFiles('w')

    #Create zip file
    zipped = zipfile.ZipFile('Resources\LabDB.zip', 'w', zipfile.ZIP_DEFLATED)

    #Populate zip file
    for root, dirs, files in os.walk(dir+'Resources\LabDBWrite'):
        for file in files:
            os.chdir('Resources\LabDBWrite')
            filePath = os.path.join(root, file)
            filePath = filePath.replace(dir+'\LabDBWrite\\','')
            zipped.write(filePath)
    os.chdir(dir)
    zipped.close()
    os.rename(dir+'\LabDB.zip',dir+'\Lab.odb')

def main():
    deletePrevious()
    openFile()
    renameFiles('r')
    connect.readFlyStock('Resources\LabDB\database\mydb','Resources\hsqldb.jar')

    root = Tk()

    # Tabbed window implementation
    tabControl = ttk.Notebook(root)
    tab1 = ttk.Frame(tabControl)
    tab2 = ttk.Frame(tabControl)
    tabControl.add(tab1, text='Finder')
    tabControl.add(tab2, text='Add Stock')
    tabControl.grid()

    root.title('Zhou Lab - Label Printer and Stock Query') # Window title
    root.iconbitmap('Resources/print.ico') # Window icon

    # Gets time modified (in epoch format)
    modified = os.path.getmtime(LabDBodb)

    # Convert epoch format to MM/DD/YY Hours:Minutes:Seconds
    modified = time.strftime('%m/%d/%y %I:%M:%S', time.localtime(modified))

    # Populate tabs
    finder.Finder(root,tab1,modified)
    adder.Adder(root,tab2,modified)

    root.mainloop()


if __name__=="__main__":
    main()
