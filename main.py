import openpyxl as xl
import os
import json

# Get file path
filepath = input("Please type in the document name.") or "TT Brothers Information"
path = os.getcwd()
filepath = path + "/" + filepath + ".xlsx"

# Load source Excel sheet
wb1 = xl.load_workbook(filepath)
ws1 = wb1.worksheets[0]

# Get total rows/cols
maxRow = ws1.max_row
maxCol = ws1.max_column

# Set up JSON dict
json_data = {}

# Rip data and add to JSON dict
#   Headers: ID, Name (First Last), Class, Active (Y or N), LinkedIn, Major, Cabby/Exec (Y or N), Profile URL
for r in range(2, maxRow + 1):
    for c in range(1, maxCol + 1):
        print(ws1.cell(row=r, column=c).value)


# Write JSON file
with open("brother_info.json", "w") as outfile:
    json.dump(json_data, outfile)