from tkinter import (Tk, ttk, IntVar, StringVar, Label, Checkbutton, Text, Scrollbar, Listbox, messagebox, filedialog,
                    END, INSERT, ACTIVE, Toplevel, RIGHT, LEFT, HORIZONTAL, X, Y, TOP, BOTTOM)
import os # Delete files
import pandas as pd # Read .csv
import math # .ceiling and .isnan
import sys
import FlyStock
sys.path.append('/')
import connect


def Adder(root, tab, modified):

    browserFile = ''
    global openG, openM, lastG, lastM, lab

    #Finds first available stock entries
    def findMissing(L):
        start, end = L[0], L[-1]
        return sorted(set(range(start, end + 1)).difference(L))

    def findAvailable(Letter):
        global lastG, lastM, lab

        lab = pd.read_csv('Resources/LabDB.csv', index_col = False)
        stockList = lab['Stock_ID'][lab['Stock_ID'].str.contains(Letter)].str.replace(Letter,'').values.tolist()
        stockList = [int(x) for x in stockList]
        if Letter == 'G' and len(stockList)>0:
            lastG = stockList[-1]
        elif len(stockList)>0:
            lastM = stockList[-1]
        return findMissing(stockList)

    openG = findAvailable('G')
    openM = findAvailable('M')

    # Prepend letter and zeroes
    openG = ['G00'+str(x) if int(x) < 10 else 'G0'+str(x) if int(x) <100 else 'G'+str(x) for x in openG]
    openM = ['M00'+str(x) if int(x) < 10 else 'M0'+str(x) if int(x) <100 else 'M'+str(x) for x in openM]

    def showAvailable():
        global openG, openM
        selectWindow = Toplevel(root)
        selectWindow.iconbitmap('Resources/print.ico') #Window icon

        selectLabel = Label(selectWindow, text="Available \"G\" and \"M\" stocks displayed")
        selectLabel.grid(row=0,columnspan = 10)

        class stockDisplay:
            def __init__(self, frame, checkValue, labelName, rowNumber, buttonColumn):

                self.checkValue = IntVar(value=checkValue)
                self.labelName = labelName
                self.frame = frame

                self.Label = ttk.Label(self.frame, text=self.labelName, width=8, relief='groove')
                self.Label.grid(row=rowNumber, column=buttonColumn+1, padx=(5,5), pady=(10),stick='w')

        rows = math.ceil(max([len(openG),len(openM)])/7)

        gLabel = ttk.LabelFrame(selectWindow, text='G Stocks')
        gLabel.grid(row=1, stick='w')

        currentRow = 0
        currentColumn = 0

        available = []

        for entry in openG:
            if currentRow == rows:
                currentRow = 0
                currentColumn = currentColumn + 1
            available.append(stockDisplay(gLabel, 0, entry, currentRow, currentColumn))
            currentRow = currentRow + 1

        mLabel = ttk.LabelFrame(selectWindow, text='M Stocks')
        mLabel.grid(row=2, stick='w')

        currentRow = 0
        currentColumn = 0

        for entry in openM:
            if currentRow == rows:
                currentRow = 0
                currentColumn = currentColumn + 1
            available.append(stockDisplay(mLabel, 0, entry, currentRow, currentColumn))
            currentRow = currentRow + 1

    def addStock():

        global browserFile

        # Check before adding
        try:
            bloomington = pd.read_csv(browserFile, names = ['bun','stock','genotype','chrs','chr','insertion','blank','date','donor','notes','owner','po','rrid'])
            if len(bloomington.columns) == 13:
                addWindow = Toplevel(root)
                addWindow.iconbitmap('Resources/print.ico') #Window icon

                bloomingtonStocks = ttk.LabelFrame(addWindow, text='Stocks found')
                bloomingtonStocks.grid(row=0)

                addStockButtonFrame = ttk.LabelFrame(addWindow)
                addStockButtonFrame.grid(row=1)

                def updateSelected():
                    for stock in bloomStocks:
                        if stock.checkValue.get() == 0:
                            stock.String.set('')
                    selected = []
                    for stock in bloomStocks:
                        if stock.checkValue.get() == 1:
                            selected.append(stock)
                    amount_selected = str(len(selected))
                    matchSelectedButtonG = ttk.Button(addStockButtonFrame, text='Match Selected ('+amount_selected+') to first matching G', command = lambda: matchSelected('G', selected))
                    matchSelectedButtonG.grid(row=0,column=0)
                    matchSelectedButtonM = ttk.Button(addStockButtonFrame, text='Match Selected ('+amount_selected+') to first matching M', command = lambda: matchSelected('M', selected))
                    matchSelectedButtonM.grid(row=0,column=1)

                    addSelectedButton = ttk.Button(addStockButtonFrame, text='Add Selected ('+amount_selected+') to LabDB', command = lambda: addSelected(selected))
                    addSelectedButton.grid(row=1,columnspan=2)

                showAvailableButton = ttk.Button(addStockButtonFrame, text="Display unused stock numbers", command = lambda: showAvailable())
                showAvailableButton.grid(row=2,columnspan=2)


                class bloomStockEntry:
                    def __init__(self, checkValue, stock, genotype, chrs, notes, row):
                        self.stock = stock
                        self.genotype = genotype
                        self.chrs = chrs
                        self.notes = notes
                        self.row = row
                        self.checkValue = IntVar(value=checkValue)
                        self.String = StringVar()

                        if str(genotype) == 'nan':
                            self.genotype = ''
                        if str(chrs) == 'nan':
                            self.chrs = ''
                        if str(notes) == 'nan':
                            self.notes = ''

                        self.LabelStocks = ttk.Label(bloomingtonStocks, text=str(self.stock))
                        self.LabelStocks.grid(row=self.row, column=2, padx=(10,0), stick='w')

                        self.LabelChrs = ttk.Label(bloomingtonStocks, text=str(self.chrs))
                        self.LabelChrs.grid(row=self.row, column=3, stick='w')

                        if str(self.genotype) != 'nan':
                            self.LabelNotes = ttk.Label(bloomingtonStocks, text=str(self.genotype))
                            self.LabelNotes.grid(row=self.row, column=4, stick='w')

                        self.Check = Checkbutton(bloomingtonStocks, variable = self.checkValue, command = lambda: updateSelected())
                        self.Check.grid(row=self.row, column=0, padx=(10,0),stick='e')

                        self.Entry = ttk.Entry(bloomingtonStocks, textvariable = self.String, width = 5)
                        self.Entry.grid(row=self.row, column=1, padx=(0,10))

                bloomStocks = []
                for index, row in bloomington.iterrows():
                    if index != 1 and not math.isnan(row['stock']) :
                        bloomStocks.append(bloomStockEntry(1,int(row['stock']),row['genotype'],row['chrs'],row['notes'],index))

                updateSelected()

        except:
            error = messagebox.showerror("Error","Bloomington .csv not found")

    def matchSelected(Letter, selected):
        global openG, openM, lastG, lastM, lab
        # Appends letter and zeroes
        for index, selection in enumerate(selected):
            if Letter == 'G':
                try:
                    selection.String.set(openG[index])

                # More stocks to add than entries available, so make new ones
                except IndexError:
                    selectionString = 'G'+str(lastG+index-len(openG)+1)
                    while(len(selectionString)<4):
                        selectionString = selectionString.replace('G','G0')
                    selection.String.set(selectionString)

            elif Letter == 'M':
                try:
                    selection.String.set(openM[index])

                # More stocks to add than entries available, so make new ones
                except IndexError:
                    selectionString = 'M'+str(lastM+index-len(openM)+1)
                    while(len(selectionString)<4):
                        selectionString = selectionString.replace('M','M0')
                    selection.String.set(selectionString)

        #connect.addFlyStock('Resources\LabDB\database\mydb', 'Resources\hsqldb.jar', selection)

    def addSelected(selected):
        global openG, openM, lastG, lastM, lab
        # Appends letter and zeroes
        value = []
        for index, selection in enumerate(selected):
            value.append((selection.String.get(), selection.genotype, str(selection.chrs), str(selection.notes), '', '', ''))
        connect.writeFlyStock('Resources\LabDB\database\mydb','Resources/hsqldb.jar',value)
        FlyStock.writeFile()

    def browser():
        # Declare global to modify
        global browserFile

        browserFile = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
        fileEntry.delete(0, END)
        fileEntry.insert(0, browserFile)

        # Check if the file chosen is valid
        if browserFile:
            bloomington = pd.read_csv(browserFile)
            if len(bloomington.columns) == 13:
                successLabel = Label(bloomLabel, text='File read successful.')
                successLabel.grid(row=1,columnspan=2)
            else:
                failLabel = Label(bloomLabel, text='File read unsuccessful. Is this a Bloomington .csv?')
                failLabel.grid(row=1,columnspan=2)

    modifiedLabel = Label(tab, text="LabDB.odb updated: " + modified)
    modifiedLabel.grid(row=0,columnspan=10)

    bloomLabel = ttk.LabelFrame(tab, text="Please enter a Bloomington Drosophila Stock Center .csv")
    bloomLabel.grid(row=1,columnspan=10, padx=(20,0))

    fileEntry = ttk.Entry(bloomLabel,width=70)
    fileEntry.grid(row=0,column=0,stick='w')

    fileBrowse = ttk.Button(bloomLabel, text="Browse", command = lambda: browser())
    fileBrowse.grid(row=0,column=1,stick='e')

    addButton = ttk.Button(tab, text="Add stock", command = lambda: addStock())
    addButton.grid(row=3,columnspan=10)

    showButton = ttk.Button(tab, text="Display unused stock numbers", command = lambda: showAvailable())
    showButton.grid(row=4,columnspan=10)

    writeButton = ttk.Button(tab, text="Write to Dropbox", command = lambda: writeFile())
    writeButton.grid(row=5,columnspan=10)
