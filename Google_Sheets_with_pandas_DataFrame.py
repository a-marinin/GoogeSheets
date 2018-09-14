"""
To execute following code you need:
1. Create a service account and OAuth2 credentials from the Google API Console:
https://console.developers.google.com/
2. Download the JSON file.
3. Copy the JSON file to your code directory and rename it to client_secret.json
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import time

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("PythonPandasTest").sheet1

# Extract all the data into a list of lists
list_of_lists = sheet.get_all_values()

# Create a DataFrame of data from list of lists
df = pd.DataFrame(list_of_lists)

# Set the column labels to equal the values in the 0 row (index location 0)
df.columns = df.iloc[0]
# Drop a row by number(number 0) (Because now it's duplicated on column labels)
df.drop(df.index[0], inplace=True)

# Set the index to equal the values in the 'ID' column
df.set_index('ID', inplace=True)

# Import DataFrame to Excel
df.to_excel('test.xlsx')

# Select specific row
print(df.iloc[-1])

# Get the last row in DataFrame
print(df.tail(1))

# Get quantity of rows in DataFrame
print(len(df.index))

# Write pandas DataFrame to Google sheet
def pandasDataFrameToGoogleSheets():
for i in range(len(df.index)):
    # Conver each DataFrame row to flat list
    row_list = df.iloc[i].values.flatten().tolist()
    # Set the row where you would like to set your DataFrame
    index = 25
    # Insert the flat list in the spreadsheet
    sheet.insert_row(row_list, index)
    # Delay because Google Sheets API has following limits:
    # 500 requests per 100 seconds per project
    # 100 requests per 100 seconds prt user
    time.sleep(1.05)
	
# pandasDataFrameToGoogleSheets()