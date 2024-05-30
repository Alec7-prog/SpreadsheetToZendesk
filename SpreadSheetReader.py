import pandas as pd # type: ignore
import xlsxwriter # type: ignore
import requests
import requests 

#Only focus on first two colums, which include the first and last name, respectively
require_cols = [0, 1]

#create a dataframe using the information collected from the first and second columns
required_df = pd.read_excel('test.xlsx', usecols = require_cols, dtype= {'first name':str, 'last name':str})

#separate the dataframe into two lists, one with the first names and the other with the last names
first_names = required_df['first name'].tolist()
last_names = required_df['last name'].tolist()

#create a list that will hold the first and last names of each person together
check_names = []

#go through every name in first_names, concatenate the corresponding last_name to it, and store it in check_names
x = 0
while x < len(first_names):
    check_names.append(first_names[x] + " " + last_names[x])
    x += 1

'''Using the Zendesk Sell Contacts API, this function generates a dictionary, based on the check_names list, that only holds names as keys, 
   and attatches a value of 0 or 1 to it, depending on whether the name is shortened, contains a special char, or appears on Zendesk (1) or doesn't (0)
''' 
def contactsChecker(names):
    result = {}
    special_chars = ['.', '-', '\'']
    shortened_names = ['Nelly']
    headers = {'Accept' : 'application/json', 'Authorization' : 'Bearer 3cf9249545d38c1d94c16457731eed492b4ead4d618819de0d633f7c3cb09346'}

    for name in names:
        if any(x in name for x in special_chars) or any(x in name for x in shortened_names):
            result[name] = 1
            continue
        
        url = f'https://api.getbase.com/v2/contacts/?name={name}'
        response = requests.get(url, headers = headers)

        if not response.json()['items']:
            result[name] = 0
        elif response:
            result[name] = 1
        else:
            raise Exception(f"Non-success status code: {response.status_code}")
    return result

#createes a worksheet object using the xlsxwriter library to create a new excel spreadsheet with updated information
workbook = xlsxwriter.Workbook('testOut.xlsx')
worksheet = workbook.add_worksheet()

#writes in every name from result_names (should be all of the names that were originally given) in the first column
#also assigns a color to the row the name is in, yellow or green, depending on the value associated to it 
# 0 = green, 1 = yellow
result_names = contactsChecker(check_names)
x= 1
for key, val in result_names.items():
    worksheet.write(x, 0, key)
    if val < 1:
        worksheet.set_row(x, cell_format = workbook.add_format({'bg_color' : "#00FF00"}))
    else:
        worksheet.set_row(x, cell_format = workbook.add_format({'bg_color' : '#FFFF00'}))
    x += 1

#closes the workbook
workbook.close()