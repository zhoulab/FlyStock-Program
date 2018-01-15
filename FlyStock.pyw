# This is a GUI program meant to print stock labels and query stocks
# For help ask Alain Garcia (alaingarcia@ufl.edu)

# For GUI functionality
from tkinter import Tk, ttk, IntVar, StringVar, Label, Checkbutton, Text, Scrollbar, Listbox, messagebox, filedialog
from tkinter import END, INSERT, ACTIVE, Toplevel, RIGHT, LEFT, X, Y

import time # Only used to access current day
import zipfile # To unzip .odb and extract data from database
import os # Delete files
import subprocess # Run the java query subprocess
import csv # Read the csv that is read from java subprocess
import shutil
from win32com.client import Dispatch # For communicating to DYMO printer

# Python file path C:/.../LabelPrinter/
dir = os.path.abspath(os.path.dirname(__file__))+"\Resources"

# LabDB path
#labPath = os.path.join('C:\\Users\\lab\\Dropbox (ZhouLab)\\ZhouLab Team Folder\\General\\Databases','LabDB.odb')
labPath = os.path.join(dir,'LabDB.odb')

# Delete previous mydb files (used during renameFiles)
def deletePrevious():
    """#for loop to delete previous mydb files in \database
    for file in os.listdir(dir+'\LabDB\database'):
        #delete mydb files
        if file.startswith('mydb') and not file.endswith('.tmp'):
            os.remove(os.path.join(dir+'\LabDB\database', file))

        #.tmp files are considered directories, so rmdir is required instead of remove
        elif file.endswith('.tmp'):
            os.rmdir(os.path.join(dir+'\LabDB\database', file))"""
    for root, dirs, files in os.walk(dir+"\LabDB", topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    if os.path.isdir(dir+"\LabDB"):
        os.rmdir(dir+"\LabDB")
    else:
        pass


    for root, dirs, files in os.walk(dir+"\LabDBWrite", topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    if os.path.isdir(dir+"\LabDBWrite"):
        os.rmdir(dir+"\LabDBWrite")
    else:
        pass

# Function to rename database files
def renameFiles(readOrWrite):
    if(readOrWrite == "r"):
        files = os.listdir(dir+'\LabDB\database')

        #Then, rename and 'create' new mydb
        for filename in files:
            os.rename(os.path.join(dir+'\LabDB\database', filename), os.path.join(dir+'\LabDB\database', "mydb."+filename))
    elif(readOrWrite == "w"):
        files = os.listdir(dir+'\LabDBWrite\database')

        #Then, rename and 'create' new mydb
        for filename in files:
            os.rename(os.path.join(dir+'\LabDBWrite\database', filename), os.path.join(dir+'\LabDBWrite\database', filename.replace("mydb.","")))

#Uses zipfile library to unzip and extract HSQLDB database from .odb
def openFile():
    unzip = zipfile.ZipFile(labPath, 'r')
    if os.path.isdir(dir+"\LabDB"):
        pass
    else:
        os.makedirs(dir+"\LabDB")
    unzip.extractall(dir+'\LabDB')
    unzip.close()

#Uses java to connect to database and read it. Stores in csv file.
def readDatabase():

    #Runs "java C:\...\Query CATEGORY KEYWORD" and outputs it into "out"
    os.chdir(dir)
    proc = subprocess.Popen(['java','-cp',dir+";"+dir+"\\hsqldb.jar",'Query'],stdout=subprocess.PIPE, shell=True)

    #Receive output in 'out'
    (out, err) = proc.communicate()
    #Receives the java output which is encoded by latin-1 (comma delimited)
    out = str(out,'latin-1')

    #Writes it to a file
    file = open(dir+'\LabDB.csv','w')
    file.write(out)
    file.close()

def writeFile():

    if os.path.isfile(dir+"\Lab.odb"):
        os.remove(dir+"\Lab.odb")

    #Populate LabDBWrite
    shutil.copytree(dir+"\LabDB",dir+"\LabDBWrite")
    os.remove(dir+"\LabDBWrite\database\mydb.lck")
    os.remove(dir+"\LabDBWrite\database\mydb.properties")
    os.remove(dir+"\LabDBWrite\database\mydb.script")

    #Remove mydb. prefix for a write ("w")
    renameFiles("w")
    shutil.copy(dir+"\Functional\script",dir+"\LabDBWrite\database")
    shutil.copy(dir+"\Functional\properties",dir+"\LabDBWrite\database")

    #Create zip file
    zipped = zipfile.ZipFile(dir+"\LabDB.zip", 'w', zipfile.ZIP_DEFLATED)

    #Populate zip file
    for root, dirs, files in os.walk(dir+"\LabDBWrite"):
        for file in files:
            os.chdir(dir+"\LabDBWrite")
            filePath = os.path.join(root, file)
            filePath = filePath.replace(dir+"\LabDBWrite\\","")
            zipped.write(filePath)
    os.chdir(dir)
    zipped.close()
    os.rename(dir+"\LabDB.zip",dir+"\Lab.odb")

def main():
    deletePrevious()
    openFile()
    renameFiles("r")
    readDatabase()

    # Gets time modified (in epoch format)
    modified = os.path.getmtime(labPath)

    # Convert epoch format to MM/DD/YY Hours:Minutes:Seconds
    modified = time.strftime('%m/%d/%y %I:%M:%S', time.localtime(modified))

    #BEGIN GUI HERE
    root = Tk()

    #Tabbed window implementation
    tabControl = ttk.Notebook(root)
    tab1 = ttk.Frame(tabControl)
    tab2 = ttk.Frame(tabControl)
    tabControl.add(tab1, text="Finder")
    tabControl.add(tab2, text="Add Stock")
    tabControl.grid()

    #root.geometry("800x400+0+0") #Window Size
    root.title("Zhou Lab - Label Printer and Stock Query") #Window title
    root.iconbitmap(dir+'\print.ico') #Window icon

    #Function for the finder/printer
    def Finder():

        #GUI properties
        stockNumberRow = 1
        genotypeRow = 2
        descriptionRow = 3
        noteRow = 4
        resPersonRow = 5
        flybaseRow = 6
        projectRow = 7
        dayRow = 8
        textboxRow = 9

        buttonColumn=0
        labelColumn=1
        entryColumn=2

        entryWidth=50

        # Parameter object to make widget declaration less repetitive.
        # Gives a parameter object its own:
        # Checkbox, Label, Entry Box, and Find button if applicable
        class Parameter:

            def __init__(self, checkValue, labelName, rowNumber, searchable):

                self.checkValue = IntVar(value=checkValue)
                self.labelName = labelName

                self.checkUpdate = False

                self.Check = Checkbutton(tab1, variable = self.checkValue, command = lambda: updateTextbox('<KeyRelease>'))
                self.Check.grid(row=rowNumber, column=buttonColumn, padx=(10,0),stick='e')

                self.Label = ttk.Label(tab1, text=self.labelName)
                self.Label.grid(row=rowNumber, column=labelColumn, padx=(0,10), pady=(10),stick='w')

                self.String = StringVar()
                self.Entry = ttk.Entry(tab1, textvariable = self.String, width = entryWidth)
                self.Entry.grid(row=rowNumber, column=entryColumn, padx=(0,20))

                if searchable:
                    self.findButton = ttk.Button(tab1, text="Find", command = lambda: query(self.labelName,self.Entry.get()))
                    self.findButton.grid(row=rowNumber,column=entryColumn+1,padx=(0,20))


        # Declaration of parameter objects that will be used
        # Entry = Parameter(default checkbox value, label name, row number, searchable)
        stockNumber = Parameter(1, "Stock ID", stockNumberRow, True)
        day = Parameter(0, "Day (DD/MM/YY)", dayRow, False)
        genotype = Parameter(1, "Genotype", genotypeRow, True)
        description = Parameter(0, "Description", descriptionRow, True)
        note = Parameter(0, "Note", noteRow, True)
        resPerson = Parameter(0, "Responsible Person", resPersonRow, True)
        flybase = Parameter(0, "Flybase", flybaseRow, True)
        project = Parameter(0, "Project", projectRow, True)

        # Places parameter objects in a tuple to make it easy
        # to iterate through them. (e.g. update them all, clear them)
        table = [stockNumber, genotype, description, note, resPerson, flybase, project, day]

        # Textbox that display preview
        textbox= Text(tab1, height=6, width=35, font=("Arial",11))
        textbox.grid(row=textboxRow+2,column=labelColumn, columnspan=2,pady=(0,20))
        textbox.tag_configure("center", justify = "center")

        # Custom labels (one time use)

        modifiedLabel = Label(tab1, text="LabDB.odb updated: " + modified)
        modifiedLabel.grid(row=0, column=labelColumn, columnspan=2, pady=(5,0))

        previewLabel = Label(tab1, text="Preview ( 2\"x 3/4\" )")
        previewLabel.grid(row=textboxRow+1,column=labelColumn,columnspan=2,pady=(20,0))

        # Custom buttons (one time use)
        todayButton = ttk.Button(tab1, text="Today", command=lambda: updateToday())
        todayButton.grid(row=dayRow,column=3,padx=(0,20))

        def updateToday():
            day.Entry.delete(0,'end')
            day.Entry.insert(0,time.strftime('%x'))

        clearButton = ttk.Button(tab1, text = "Clear Entries", command = lambda: clearEntries())
        clearButton.grid(row=textboxRow, column = labelColumn, columnspan=2)

        printButton = ttk.Button(tab1, text="Print", command=lambda: printLabel())
        printButton.grid(row=textboxRow+3,column=0,padx=(0,20),pady=(0,10),columnspan=4)

        # Updates the textbox
        def updateTextbox(event):
            #Delete current Textbox
            textbox.delete(0.0,END)
            #Insert things into textbox
            textbox.insert(INSERT,printPreview(), "center")

        # Clears all entry boxes
        def clearEntries():
            for section in table:
                section.Entry.delete(0,'end')
            updateTextbox('<Return>')

        # Makes it so any key release updates the text box. Editing entries updates textbox.
        root.bind('<KeyRelease>', updateTextbox)

        # Text wrapping to prevent letters from shrinking if one entry is too long
        # DYMO does not have text wrap as far as I know
        def textWrap(length, string):
            stringLen = len(string)
            if stringLen >= length:
                newString = ''
                #If you want 25 characters per line and your string is 50 characters long, there should be two lines.
                splitString = [string[i:i+length] for i in range(0, len(string), length)]
                for i in splitString:
                    newString = newString + i + '\n'
                return newString


        #Creates the print preview and returns it as a string
        def printPreview():
            stringPreview = ""
            for column in table:
                if column.checkValue.get() == 1:
                    tempString = column.Entry.get()
                    #25 is the arbritrary value chosen to cut the text off and start the next line
                    if len(tempString) < 25:
                        stringPreview = stringPreview + tempString + "\n"
                    else:
                        stringPreview = stringPreview + textWrap(25, tempString)

            return stringPreview

        # This function uses java to query LabDB.odb
        # Java is necessary to connect to a HSQLDB
        # We will use this function to return java's output and update text boxes
        def query(category, keyword):
            #If there is a blank entry
            if keyword == "":
                error = messagebox.showerror("Find Error", "Blank entry detected. Please input a searchable term.")
            #An entry has been submitted, so it will now be searched
            else:
                try:
                    table2 = ["Stock ID", "Genotype", "Description", "Note", "Responsible Person", "Flybase", "Project"]
                    columnNumber = int(table2.index(category)) + 1
                    resultNumber = 0
                    queryResult = []

                    with open(dir+"\LabDB.csv", "r") as csvRead:
                        lab = csv.reader(csvRead)
                        for row in lab:
                            currentColumn = 0
                            for col in row:
                                currentColumn = currentColumn + 1
                                if currentColumn == columnNumber:
                                    if col.find(keyword) != -1:
                                        queryResult.append(row)
                                        resultNumber = resultNumber + 1
                    #No result
                    if (resultNumber == 0):
                        error = messagebox.showerror("Find Error", "Entry could not be found with %s \"%s\"" % (category,keyword))

                    #If there is only one stock result, do not show list
                    elif (resultNumber == 1):
                        for i, column in enumerate(table):
                        #delete current entry
                            column.Entry.delete(0,'end')
                        #update entry
                            try:
                                column.Entry.insert(0,(queryResult[0])[i])
                            except:
                                pass
                        updateTextbox('<Return>')

                    #If there are multiple results, let the user select from list
                    elif (resultNumber>1):
                        #declare new window popup for list selection
                        selectWindow = Toplevel(root)
                        selectWindow.iconbitmap(dir+'\print.ico') #Window icon

                        selectLabel = Label(selectWindow, text="Multiple stocks meet criteria. Please select one and then press \"Select\".")
                        selectLabel.pack()

                        selectScrollbar = Scrollbar(selectWindow)
                        selectScrollbar.pack(side=RIGHT, fill=Y)

                        #Different heights
                        if resultNumber<50:
                            selectListbox = Listbox(selectWindow, height=resultNumber, width=100, yscrollcommand=selectScrollbar.set)
                        else:
                            selectListbox = Listbox(selectWindow, height=50, width=100, yscrollcommand=selectScrollbar.set)

                        selectListbox.pack()

                        selectScrollbar.config(command=selectListbox.yview)

                        #declare listbox that shows all selections
                        #listbox = Listbox(selectionWindow,height=resultNumber,width=100)
                        #listbox.grid(row=1,column=0, padx = 20)

                        #populate listbox will results
                        for i in range(0,resultNumber):
                            selectListbox.insert(i, queryResult[i])
                        """for i in range(0,resultNumber):
                            for j in range(0,7):
                                selectListbox.insert(ACTIVE, queryResult[i][j])"""

                        #function for selection of listbox
                        def select(active,selectWindow):
                            for i, column in enumerate(table):
                                #delete current entry
                                column.Entry.delete(0,'end')
                                #update entry
                                try:
                                    column.Entry.insert(0,active[i])
                                except:
                                    pass
                            updateTextbox('<Return>')
                            selectWindow.destroy()

                        #declare select button
                        selectButton = ttk.Button(selectWindow,text="Select",command = lambda: select(selectListbox.get(ACTIVE),selectWindow))
                        selectButton.pack()

                        selectWindow.mainloop()

                #Something has gone wrong with finding using keyword
                except Exception as e:
                    print(e)
                    error = messagebox.showerror("Find Error", "Entry could not be found with %s \"%s\"" % (category,keyword))

        #PRINT
        def printLabel():

            #my.label should be in directory
            mylabel = os.path.join(dir,'my.label')

            #if my.label is not in directory
            if not os.path.isfile(mylabel):
                messagebox.showinfo('PyDymoLabel','Template file my.label does not exist')
                sys.exit(1)

            try:
                #Use COM to run DYMO Printer
                labelCom = Dispatch('Dymo.DymoAddIn')
                labelText = Dispatch('Dymo.DymoLabels')

                #Opens label and edits it
                isOpen = labelCom.Open(mylabel)
                labelCom.SelectPrinter('DYMO LabelWriter 450')
                labelText.SetField('TEXT', printPreview())

                #Print command
                labelCom.StartPrintJob()
                labelCom.Print(1,False)
                labelCom.EndPrintJob()

            except Exception as e:
                messagebox.showinfo('PyDymoLabel','An error occurred during printing.')
                sys.exit(1)

            messagebox.showinfo('PyDymoLabel','Label printed!')

    def Adder():
        availableG = []
        availableM = []

        #Finds first available stock entries
        def findMissing(L):
            start, end = L[0], L[-1]
            return sorted(set(range(start, end + 1)).difference(L))

        def findAvailable():
            listG = []
            listM = []

            with open(dir+"\LabDB.csv", "r") as csvRead:
                lab = csv.reader(csvRead)
                first = True
                for row in lab:
                    stockID = row[0]
                    if stockID.find("G") and first == False:
                        #Cuts off "G" and turns into integer
                        listG.append(int(stockID[1:]))
                    elif stockID.find("M") and first == False:
                        listM.append(int(stockID[1:]))
                    else:
                        first = False

            availableG = (findMissing(listG))
            availableM = (findMissing(listM))

        def addStock():
            os.chdir(dir)
            proc = subprocess.Popen(['java','-cp',dir+";"+dir+"\\hsqldb.jar",'Add'], stdout=subprocess.PIPE, shell=True)

            #Receive output in 'out'. For some reason necessary to keep code working.
            (out, err) = proc.communicate()

            #Removes lock
            os.remove(dir+"\LabDB\database\mydb.lck")

            #Updates database search with new entries
            readDatabase()


        modifiedLabel = Label(tab2, text="LabDB.odb updated: " + modified)
        modifiedLabel.grid(row=0,columnspan=2)

        bloomLabel = Label(tab2, text="Please enter a Bloomington Drosophila Stock Center .csv")
        bloomLabel.grid(row=1, columnspan=2)

        fileEntry = ttk.Entry(tab2,width=70)
        fileEntry.grid(row=2,column=0,padx=(20,10))

        def browser():
            browse = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
            fileEntry.insert(0, browse)

        fileBrowse = ttk.Button(tab2, text="Browse", command = lambda: browser())
        fileBrowse.grid(row=2,column=1)

        addButton = ttk.Button(tab2, text="Add stock", command = lambda: addStock())
        addButton.grid(row=3,columnspan=2)

        writeButton = ttk.Button(tab2, text="Write to Dropbox", command = lambda: writeFile())
        writeButton.grid(row=4,columnspan=2)

    Finder()
    Adder()
    root.mainloop()


if __name__=="__main__":
    main()
