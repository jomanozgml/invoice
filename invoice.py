# Python 3.8.8
# Collecting all required information from imported 'daraz' CSV file
# Exporting required/collected information to spreedsheet file to fill in the invoice

# Importing required libraries
import pandas as pd
import string
import random
import os
import tkinter as tk
from num2words import num2words as n2w
from tkinter import filedialog

# Function definition for file selection after button is pressed
def openFile():
	filePathName = filedialog.askopenfilename(initialdir = "C:/Users/MYTH/Downloads",
                                          title = "Select an imported CSV File",
                                          filetypes = (("CSV files", "*.csv*"), ("all files", "*.*")))
	# Makes some modification on file like filtering and removing
	editFile(filePathName)
	# Closes the GUI after work completion
	window.destroy()
	# Starts Bill_Print word document for mail merge
	os.system('start ../Bill_Print.docx')

# GUI creation steps
window = tk.Tk()
window.title("Invoice Application")
window.geometry("200x300")
window.configure(background='lightsteelblue')
topLabel = tk.Label(text = 'MOONSTAR TAILORS', width = 36, height = 2, bg = 'silver', fg = 'black')
topLabel.config(font=('Helvetica bold', 12))
topLabel.pack(pady=(0,40))
buttonOpenFile = tk.Button(text = 'Select a CSV file', width = 18, height = 3, bg = 'teal', fg = 'white', command = openFile)
buttonOpenFile.config(font = ('Helvetica 10 bold'))
buttonOpenFile.pack(pady=15)
buttonCloseWindow = tk.Button(text = 'Close Application', width = 18, height = 3, bg = 'crimson', fg = 'white', command = window.destroy)
buttonCloseWindow.config(font = ('Helvetica 10 bold'))
buttonCloseWindow.pack(pady=15)
# Makes window appear topmost for greater visibiity
window.attributes("-topmost", True)

def editFile(fileName):
	# Reading CSV file and assigning to csvFile dataframe
	csvFile = pd.read_csv(fileName, sep = ';')

	# Empty list for storing total price in words
	amountInWords = []
	# Assigning csv file index total length to rowSize variable
	rowSize = len(csvFile.index)
	# Empty list for storing random internal invoice number
	invoiceNum = []

	# Creating list with column names for multiple order handling
	multipleOrderColName = ['Item2', 'Var2', 'Price2', 'Item3', 'Var3', 'Price3', 'Item4', 'Var4', 'Price4', 'Item5', 
		'Var5', 'Price5', 'Item6', 'Var6', 'Price6', 'Total Price', 'Amount in Words']
	# Creating list of only required column names from csv file
	cols = [
		'Created at',
		'Order Number',
		'Shipping Name',
		'Shipping Address',
		'Shipping Phone Number',
		'Item Name',
		'Variation',
		'Unit Price',
		'Tracking Code',
		'Invoice Number'
	]
	# Assigning None to all new cell with new column names
	for newColName in multipleOrderColName:
		csvFile[newColName] = [None] * rowSize
	# Introduing Total Price variable for multiple order handling
	csvFile['Total Price'] = csvFile['Unit Price']
	# Creating empty list to store indices of columns to delete
	listToDrop = []
	# Offset variable come in play if there more than 2 items in single order
	offset = 1

	# Loop to traverse row wise in dataframe of csvfile
	for i in csvFile.index:
		# Random invoice number stored in list
		invoiceNum.append(random.getrandbits(16))
		# Converting Total Price in words and storing in dataframe
		csvFile.at[i, 'Amount in Words'] = n2w(csvFile.iloc[i]['Total Price']) + ' rupees only'
		# i = 0 is skipped since i-1 will be used later
		if i > 0:
			# Checking if single order has mutiple items
			if csvFile.iloc[i]['Order Number'] == csvFile.iloc[i-offset]['Order Number']:
				# If yes row to later details is put in list to delete later
				listToDrop.append(i)
				# Variable x, y to go to specific location of new column names for multiple itemed order
				x, y = 0, 3
				# Checks for exact number of items
				for j in range(2,7):
					itemNum = 'Item' + str(j)
					if csvFile.iloc[i-offset][itemNum] == None:
						# List slice operation to copy to exact required cells
						for newColName, colName in zip(multipleOrderColName[x:y], cols[5:8]):
							csvFile.at[i-offset, newColName] = csvFile.iloc[i][colName]
						csvFile.at[i-offset, 'Total Price'] += csvFile.iloc[i]['Unit Price']
						csvFile.at[i-offset, 'Amount in Words'] = n2w(csvFile.iloc[i-offset]['Total Price']) + ' rupees only'
						offset = offset+1
						break
					else:
						# If new column name cells are already occupied pointer moves to next set of items location
						x, y = x+3, y+3
			else:
				offset = 1

	csvFile['Invoice Number'] = invoiceNum
	# Duplicate rows are deleted from dataframe
	csvFile = csvFile.drop(listToDrop)
	# Creating dataframe with only required column names
	updatedCSV = csvFile[cols + multipleOrderColName]
	# Exporting data from dataframe to CSV file
	updatedCSV.to_csv('../output.csv', index = False)
	# Adding new invoices to all invoice collection file
	updatedCSV.to_csv('../all_invoices.csv', mode='a', index=False, header=False)

window.mainloop()

# End of the line
