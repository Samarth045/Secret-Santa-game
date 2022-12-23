import random
import openpyxl

# Open the Excel workbook
workbook = openpyxl.load_workbook('names.xlsx')

# Get the sheet with the names
sheet = workbook.get_sheet_by_name('Sheet1')

# Read the names from the sheet into a list
names = []
for row in sheet.rows:
    names.append(row[0].value)

# Shuffle the list of names
random.shuffle(names)

# Assign the names to participate in the Secret Santa game
randomname = {}
for i in range(len(names)):
    randomname[names[i]] = names[(i+1) % len(names)]

# Print the randomname
for name, secret_santa in randomname.items():
    print(f"{name}'s Secret Santa is {secret_santa}")
