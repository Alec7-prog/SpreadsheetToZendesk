import pandas as pd # type: ignore
import xlsxwriter # type: ignore
import requests
import requests 

#Only focus on first two colums, which include the first and last name, respectively
require_cols = [0, 1]

#create a dataframe using the information collected from the first and second columns
required_df = pd.read_excel('mexico commonality list copy.xlsx', usecols = require_cols, dtype= {'first name':str, 'last name':str})

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

'''this function generates a list, based on the check_names list, that only holds names
 that show results when the Zendesk Sell Contacts API is checked''' 
def contactsChecker(names):
    result = {}
    headers = {'Accept' : 'application/json', 'Authorization' : 'Bearer 3cf9249545d38c1d94c16457731eed492b4ead4d618819de0d633f7c3cb09346'}

    for name in names:
        url = f'https://api.getbase.com/v2/contacts/?name={name}'
        response = requests.get(url, headers = headers)
        if not response.json()['items']:
            result[name] = 0
        elif response:
            result[name] = 1
        else:
            raise Exception(f"Non-success status code: {response.status_code}")
    return result

'''a dictionary that will associate a value of either 0/1 to the key (name). 
This value is determined on whether or not the name is present in names_present. 
If it is, the value associated is 1. If not, the value associated is 0.'''

#createes a worksheet object using the xlsxwriter library to create a new excel spreadsheet with updated information
#the file it will be written out to is called testOut.xlsx
workbook = xlsxwriter.Workbook('mexico commonality list results.xlsx')
worksheet = workbook.add_worksheet()

'''writes in every name from result_names (should be all of the names that were originally given) in the first column, 
and in the second column it writes in the value on whether or not it is present in ZenDesk'''
result_names = contactsChecker(check_names)
x= 0
for key, val in result_names.items():
    worksheet.write(x, 0, key)
    if val < 1:
        worksheet.set_row(x, cell_format = workbook.add_format({'bg_color' : "#00FF00"}))
    else:
        worksheet.set_row(x, cell_format = workbook.add_format({'bg_color' : '#FF0000'}))
    x += 1

#this highlights every name with the value 1 red, because it's already in the database
#all other names are highlighted green

""" red = workbook.add_format({'bg_color' : '#FF0000'})
green = workbook.add_format({'bg_color' : "#00FF00"})
worksheet.conditional_format('A1:A6', {'type' : 'cell', 'criteria' : '>', 'value' : 0, 'format' : red})
worksheet.conditional_format('A1:A6', {'type' : 'cell', 'criteria' : '<', 'value' : 1, 'format' : green}) """
#closes the workbook
workbook.close()