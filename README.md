# LabelPrinter

A simple Tkinter GUI program made to search through the Lab Stock Database and print them directly. 

Searches can be made by: 
1. Stock ID
2. Genotype
3. Description
4. Note
5. Responsible Person
6. Flybase
7. Project

The user can select the entries they want on the label with checkboxes. A print preview is displayed to show what will be printed.

If the search has multiple results (i.e. searching by someone's name with the "Responsible Person" entry), the user will be prompted to select a specific stock.

## Requirements

1. 'pypiwin32' (Necessary for printing through COM)
2. 'Python' (Confirmed working with Python 3.xx)
3. 'Java >= 8' (Necessary for connection to .odb)



