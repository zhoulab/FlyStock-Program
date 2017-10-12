# This is a GUI program meant to print stock labels and query stocks
# For help ask Alain Garcia (alaingarcia@ufl.edu)

from tkinter import * # For GUI functionality
from tkinter import messagebox # Messagebox needs to be imported individually because it is considered its own module
import time # Only used to access current day
import zipfile # To unzip .odb and extract data from database
import os # Delete files
import subprocess # Run the java query subprocess
from win32com.client import Dispatch # For communicating to DYMO printer

# Python file path C:/.../LabelPrinter/
dir = os.path.abspath(os.path.dirname(__file__))

# LabDB path
labPath = os.path.join('C:\\Users\\lab\\Dropbox (ZhouLab)\\ZhouLab Team Folder\\General\\Databases','LabDB.odb')

# Delete previous mydb files (used during renameFiles)
def deletePrevious():
    #for loop to delete previous mydb files in \database
    for file in os.listdir(dir+'\database'):

        #delete mydb files
        if file.startswith('mydb') and not file.endswith('.tmp'):
            os.remove(os.path.join(dir+'\database', file))

        #.tmp files are considered directories, so rmdir is required instead of remove
        elif file.endswith('.tmp'):
            os.rmdir(os.path.join(dir+'\database', file))

# Function to rename database files
def renameFiles():

    files = os.listdir('database')

    #Then, rename and 'create' new mydb
    for filename in files:
        os.rename(os.path.join(dir+'\database', filename), os.path.join(dir+'\database', "mydb."+filename) )

#Uses zipfile library to unzip and extract HSQLDB database from .odb
def openFile():
    unzip = zipfile.ZipFile(labPath, 'r')
    for file in unzip.namelist():
        if file.startswith('database/'):
            unzip.extract(file,dir)
    unzip.close()


def main():
    deletePrevious()
    openFile()
    renameFiles()

    #BEGIN GUI HERE
    root=Tk()

    #root.geometry("800x400+0+0") #Window Size
    root.title("Zhou Lab - Label Printer and Stock Query") #Window title
    root.iconbitmap(dir+'\print.ico') #Window icon

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

            self.Check = Checkbutton(variable = self.checkValue, command = lambda: updateTextbox('<KeyRelease>'))
            self.Check.grid(row=rowNumber, column=buttonColumn, padx=(10,0),stick='e')

            self.Label = Label(text=self.labelName)
            self.Label.grid(row=rowNumber, column=labelColumn, padx=(0,10), pady=(10),stick='w')

            self.String = StringVar()
            self.Entry = Entry(textvariable = self.String, width = entryWidth)
            self.Entry.grid(row=rowNumber, column=entryColumn, padx=(0,20))

            if searchable:
                self.findButton = Button(text="Find", command = lambda: query(self.labelName+" ||| "+self.Entry.get()))
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
    textbox= Text(root,height=6,width=35, font=("Arial",11))
    textbox.grid(row=textboxRow+2,column=labelColumn, columnspan=2,pady=(0,20))
    textbox.tag_configure("center", justify = "center")

    # Custom labels (one time use)
    previewLabel = Label(root, text="Preview ( 2\"x 3/4\" )")
    previewLabel.grid(row=textboxRow+1,column=labelColumn,columnspan=2,pady=(20,0))

    # Custom buttons (one time use)
    todayButton = Button(root,text="Today", command=lambda: day.Entry.insert(0,time.strftime('%x')))
    todayButton.grid(row=dayRow,column=3,padx=(0,20))

    clearButton = Button(root, text = "Clear Entries", command = lambda: clearEntries())
    clearButton.grid(row=textboxRow, column = labelColumn, columnspan=2)

    printButton = Button(root,text="Print", command=lambda: printLabel())
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
    def query(keyword):
        keyword = keyword.split(" ||| ")
        #If there is a blank entry
        if keyword[1] == "":
            error = messagebox.showerror("Find Error", "Blank entry detected. Please input a searchable term.")
        #An entry has been submitted, so it will now be searched
        else:
            try:
                #Special case for responsible person,
                #keyword[0] is not the searchable keyword
                if keyword[0] == "Responsible Person":
                    keyword[0] = "Res_Person"

                #Runs "java C:\...\Query CATEGORY KEYWORD" and outputs it into "out"
                searchCommand =((keyword[0]).replace(" ","_")+" "+keyword[1])
                proc = subprocess.Popen(['java','-Dmydir='+dir,'Query',searchCommand],stdout=subprocess.PIPE, shell=True)

                #Receive output in 'out'
                (out, err) = proc.communicate()
                #String manipulation to work with java output, delimiter = ' ||| '
                out = out.decode('ascii')
                out = out.replace("\r\n"," ||| ")
                out = out.split(" ||| ")
                

                #Determine amount of results by dividing by 7
                #Each result has 7 parameters, so dividing by 7 gives result count
                lineCount = int(len(out)/7)

                #If there is only one stock result, do not show list
                if (lineCount == 1):
                    for i, column in enumerate(table):
                        #delete current entry
                        column.Entry.delete(0,'end')
                        #update entry
                        if out[i] != "null":
                            column.Entry.insert(0,out[i])
                        else:
                            column.Entry.insert(0," ")
                    updateTextbox('<Return>')

                #If there are multiple results, let the user select from list
                elif (lineCount>1):
                    #declare new window popup for list selection
                    selectionWindow = Toplevel(root)
                    selectionWindow.iconbitmap(dir+'\print.ico') #Window icon

                    selectLabel = Label(selectionWindow, text="Multiple stocks meet criteria. Please select one and then press \"Select\".")
                    selectLabel.grid(row=0,column=0, pady=20)

                    #declare listbox that shows all selections
                    listbox = Listbox(selectionWindow,height=lineCount,width=100)
                    listbox.grid(row=1,column=0, padx = 20)

                    #populate listbox will results
                    for i in range(0,lineCount):
                        begin = (i*7)
                        end = (i+1)*7-1
                        listbox.insert(i, out[begin:end])

                    #function for selection of listbox
                    def select(active,selectionWindow):
                        for i, column in enumerate(table):
                            #delete current entry
                            column.Entry.delete(0,'end')
                            #update entry
                            try:
                                if out[i] != "null":
                                    column.Entry.insert(0,active[i])
                                else:
                                    column.Entry.insert(0," ")
                            except:
                                pass
                        updateTextbox('<Return>')
                        selectionWindow.destroy()

                    #declare select button
                    selectButton = Button(selectionWindow,text="Select",command = lambda: select(listbox.get(ACTIVE),selectionWindow))
                    selectButton.grid(row=2,pady=20)

                    selectionWindow.mainloop()

            #Something has gone wrong with finding using keyword
            except Exception as e:
                print(e)
                error = messagebox.showerror("Find Error", "Entry could not be found with %s \"%s\"" % (keyword[0],keyword[1]))

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
            print(e)
            messagebox.showinfo('PyDymoLabel','An error occurred during printing.')
            sys.exit(1)

        messagebox.showinfo('PyDymoLabel','Label printed!')

    root.mainloop()


if __name__=="__main__":
    main()
