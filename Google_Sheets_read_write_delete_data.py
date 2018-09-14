"""
To execute following code you need:
1. Create a service account and OAuth2 credentials from the Google API Console:
https://console.developers.google.com/
2. Download the JSON file.
3. Copy the JSON file to your code directory and rename it to client_secret.json
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("PythonPandasTest").sheet1

"""
Read data from Google Sheets
"""
# Extract and print all the data into a list of hashes
# [ATTENTION]: Need to have a data on the spreadsheet. If not -> raise the error "IndexError"
list_of_hashes = sheet.get_all_records()
print(list_of_hashes)

# Extract and print all the data into a list of lists
list_of_lists = sheet.get_all_values()
print(list_of_lists)

# Extract and print the data from one row (строка)
row = sheet.row_values(2)
print(row)

# Extract and print the data from one column (столбец)
column = sheet.col_values(2)
print(column)

# Extract and print the data from a single cell (клетка)
cell = sheet.cell(1, 1).value
print(cell)

# Ptint the title of the current sheet
print(sheet.title)

# Find out total number of rows in spreadsheet
sheet.row_count

"""
Write data to Google Sheets
"""
# Write to the spreadsheet by changing a specific cell
sheet.update_cell(1, 1, "I just wrote to a spreadsheet using Python!")
print("Your data have been written to the spreadsheet")

# Insert a row in the spreadsheet
row = ["I'm","inserting","a","row","into","a","Spreadsheet","with","Python"]
index = 7
sheet.insert_row(row, index)
print("Your data have been suckesfully inserted in the spreadsheets into the row " + str(index))

"""
Delete data from Goolge Sheets
"""
# Delete a row from the spreadsheet
sheet.delete_row(1)



