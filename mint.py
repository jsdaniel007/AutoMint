#!/usr/bin/env python

#imports
import openpyxl
import os

# Workhorse Functions

# Purpose: Take new data from the mint workbook, and transfer it to the destination workbook
# Notes: Prepend the newer entries, ignore repeat dates
def mintPull(mintBook, destBook):

	# Get basic solution working
	mintSheet = mintBook.active
	destSheet = destBook.active
	destDates = setFromColumn(destSheet, 'A')

	i = 2
	while i <= mintSheet.max_row:
		date = mintSheet.cell(row=i, column=1).value
		if not date in destDates:
			destSheet.insert_rows(1, amount=1) # Pre-pend blank row
			copyRow(mintSheet, destSheet, 2)
		i += 1

	#print_rows(destSheet)

# Helper Functions
def initWorkbook(src):
	workbookObj = openpyxl.load_workbook(src)
	return workbookObj

# Copy the row range and from one sheet to another -- PARALLEL ONLY
def copyRow(sheetCopy, sheetPaste, rowPlace):
	copyList = []
	# Copy to a list to copy to
	for row in sheetCopy.iter_rows(rowPlace):
		print()
		for cell in row:
			copyList.append(cell.value)
			print(cell.value)

	# Paste into the new sheet
	for row in sheetPaste.iter_rows(rowPlace):
		for cell in range(len(row)):
			print("cell:", cell, row[cell].value, copyList[cell])
			row[cell].value = copyList[cell]


# Create a set based on the Excel Column Name passed in
def setFromColumn(sheet, column):
	colSet = set({})
	for columnCell in sheet[column]:
		colSet.add(columnCell.value)
	return colSet

# Print the rows for the worksheet
def print_rows(sheet, isPrint=True):
	print("==============\nSheet Report for: ", sheet.title)
	print("Rows:", sheet.max_row, "Cols:", sheet.max_column)

	i = 0
	if isPrint == True:
		for row in sheet.iter_rows(values_only=True):
			print(f'Line {i}:', row)
			i += 1

# Driver Code
if __name__ == '__main__':
	# Universal Test Flags
	printFlag = True
	saveFlag = True

	# get CLI arguments <python3> <programName> <mintsrc> <destsrc>
	#mintsrc = sys.argv[1]
	#destsrc = sys.argv[2]
	mintsrc = 'test/mintTest.xlsx'
	destsrc = 'test/output.xlsx'

	mintBook = initWorkbook(mintsrc)
	destBook = initWorkbook(destsrc)

	mintPull(mintBook, destBook)

	#print_rows(destBook.active)
	if saveFlag:
		destBook.save('test/output.xlsx')
		pass
	print("Done")
