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

# Helper Functions
def initWorkbook(src):
	workbookObj = load_workbook(src)
	return workbookObj

# Copy the row range and from one sheet to another -- PARALLEL ONLY
def copyRow(sheetCopy, sheetPaste, row):
	for copyCell in sheetCopy.iter_rows(min_row=row, max_row=row):
		for pasteCell in sheetPaste.iter_rows(min_row=row, max_row=row):
			pasteCell.value = copyCell.value


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
if __name__ == "main":
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

	print_rows(destBook)

	if not saveFlag:
		destBook.save('test/output.xlsx')
		pass
