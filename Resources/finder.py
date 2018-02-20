from tkinter import (Tk, ttk, IntVar, StringVar, Label, Checkbutton, Text, Scrollbar, Listbox, messagebox,
                    END, INSERT, ACTIVE, Toplevel, RIGHT, HORIZONTAL, X, Y, TOP, BOTTOM)
import time # Only used to access current day
import os # Delete files
import pandas as pd # Read .csv
import math # Checking for NaN
import re # Query LabDB.csv
from win32com.client import Dispatch # For communicating to DYMO printer

# Python file path C:/.../LabelPrinter/
dir = os.path.abspath(os.path.dirname(__file__))

def Finder(root,tab,modified):

    #GUI properties
    order = {'stockNumber':1,'genotype':2,'description':3,'note':4,
            'resPerson':5,'flybase':6,'project':7,'day':8,'textbox':9,
            'buttonColumn':0,'labelColumn':1,'entryColumn':2}

    entryWidth=50

    # Parameter object to make widget declaration less repetitive.
    # Gives a parameter object its own:
    # Checkbox, Label, Entry Box, and Find button if applicable
    class Parameter:

        def __init__(self, checkValue, labelName, rowNumber, searchable, column):

            self.checkValue = IntVar(value=checkValue)
            self.labelName = labelName
            self.column = column

            self.Check = Checkbutton(tab, variable = self.checkValue, command = lambda: updateTextbox('<KeyRelease>'))
            self.Check.grid(row=rowNumber, column=order['buttonColumn'], padx=(10,0),stick='e')

            self.Label = ttk.Label(tab, text=self.labelName)
            self.Label.grid(row=rowNumber, column=order['labelColumn'], padx=(0,10), pady=(10),stick='w')

            self.String = StringVar()
            self.Entry = ttk.Entry(tab, textvariable = self.String, width = entryWidth)
            self.Entry.grid(row=rowNumber, column=order['entryColumn'], padx=(0,20))

            if searchable:
                self.findButton = ttk.Button(tab, text="Find", command = lambda: query(self.column,self.Entry.get()))
                self.findButton.grid(row=rowNumber,column=order['entryColumn']+1,padx=(0,20))


    # Declaration of parameter objects that will be used
    # Entry = Parameter(default checkbox value, label name, row number, searchable)
    stockNumber = Parameter(1, "Stock ID", order['stockNumber'], True, 1)
    day = Parameter(0, "Day (DD/MM/YY)", order['day'], False, None)
    genotype = Parameter(1, "Genotype", order['genotype'], True, 2)
    description = Parameter(0, "Description", order['description'], True, 3)
    note = Parameter(0, "Note", order['note'], True, 4)
    resPerson = Parameter(0, "Responsible Person", order['resPerson'], True, 5)
    flybase = Parameter(0, "Flybase", order['flybase'], True, 6)
    project = Parameter(0, "Project", order['project'], True, 7)

    # Places parameter objects in a tuple to make it easy
    # to iterate through them. (e.g. update them all, clear them)
    table = [stockNumber, genotype, description, note, resPerson, flybase, project, day]

    # Textbox that display preview
    textbox= Text(tab, height=6, width=35, font=("Arial",11))
    textbox.grid(row=order['textbox']+2,column=order['labelColumn'], columnspan=2,pady=(0,20))
    textbox.tag_configure("center", justify = "center")

    # Custom labels (one time use)
    modifiedLabel = Label(tab, text="LabDB.odb updated: " + modified)
    modifiedLabel.grid(row=0, column=order['labelColumn'], columnspan=2, pady=(5,0))

    previewLabel = Label(tab, text="Preview ( 2\"x 3/4\" )")
    previewLabel.grid(row=order['textbox']+1,column=order['labelColumn'],columnspan=2,pady=(20,0))

    # Custom buttons (one time use)
    todayButton = ttk.Button(tab, text="Today", command=lambda: updateToday())
    todayButton.grid(row=order['day'],column=3,padx=(0,20))

    def updateToday():
        day.Entry.delete(0,'end')
        day.Entry.insert(0,time.strftime('%x'))

    clearButton = ttk.Button(tab, text = "Clear Entries", command = lambda: clearEntries())
    clearButton.grid(row=order['textbox'], column = order['labelColumn'], columnspan=2)

    printButton = ttk.Button(tab, text="Print", command=lambda: printLabel())
    printButton.grid(row=order['textbox']+3,column=0,padx=(0,20),pady=(0,10),columnspan=4)

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

    def query(columnNumber, keyword):
        #If there is a blank entry
        if keyword == '':
            error = messagebox.showerror("Find Error", "Blank entry detected. Please input a searchable term.")
        #An entry has been submitted, so it will now be searched
        else:
            try:
                table2 = ["Stock_ID", "Genotype", "Description", "Note", "Res_Person", "Flybase", "Project"]
                # Get column name (minus 1 because 0 indexed)
                category = table2[columnNumber-1]

                # Let pandas read csv
                lab = pd.read_csv(dir+'\LabDB.csv', index_col = False)
                result = lab[lab[category].str.contains(keyword)]
                result = result.values.tolist()
                resultNumber = len(result)

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
                            column.Entry.insert(0,(result[0])[i])
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

                    yScrollbar = Scrollbar(selectWindow)
                    yScrollbar.pack(side=RIGHT, fill=Y)
                    xScrollbar = Scrollbar(selectWindow, orient=HORIZONTAL)
                    xScrollbar.pack(side=BOTTOM, fill=X)

                    #Different heights
                    if resultNumber<50:
                        selectListbox = Listbox(selectWindow, height=resultNumber, width=100, yscrollcommand=yScrollbar.set, xscrollcommand=xScrollbar.set)
                    else:
                        selectListbox = Listbox(selectWindow, height=50, width=100, yscrollcommand=yScrollbar.set, xscrollcommand=xScrollbar.set)

                    selectListbox.pack()

                    yScrollbar.config(command=selectListbox.yview)
                    xScrollbar.config(command=selectListbox.xview)

                    #declare listbox that shows all selections
                    #listbox = Listbox(selectionWindow,height=resultNumber,width=100)
                    #listbox.grid(row=1,column=0, padx = 20)

                    #populate listbox will results
                    for i in range(0,resultNumber):
                        selectListbox.insert(i, result[i])

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


        #if my.label is not in directory
        if not os.path.isfile('my.label'):
            messagebox.showinfo('PyDymoLabel','Template file my.label does not exist')
            sys.exit(1)

        try:
            #Use COM to run DYMO Printer
            labelCom = Dispatch('Dymo.DymoAddIn')
            labelText = Dispatch('Dymo.DymoLabels')

            #Opens label and edits it
            isOpen = labelCom.Open('my.label')
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
