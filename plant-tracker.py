import xlsxwriter
import pandas as pd

# Variable arguments for the program

excel_file = 'filelocation/excel_file.xlsx'
export_filename = 'C:\Users\xanderinotoko\Documents'
# all_sheets = pd.read_excel(excel_file, sheet_name=None)
# sheets = all_sheets.keys()

"""
Need to use argv to determine name of xlsx file to import

"""

"""
ABOUT BELOW CODE:

I am using python3.x in Anaconda environment and In my case file name is 'INDIA-WMS.xlsx' 
having 40 different sheets below code will create 40 different csv files named as sheet 
name of excel file, as 'key.csv'. Hope this will help your issue.

For example if you have different sheets like 'Sheet1', 'Sheet2', 'Sheet3' etc. then above 
code will create different csv file as 'Sheet1.csv', 'Sheet2.csv', 'Sheet3.csv'. Here 'key' 
is the sheet name of your excel workbook. If you want to use data content inside sheets you 
can use the for loop as for key, value in df.items():
"""

# Create dataframe from excel file
df = pd.read_excel(excel_file, sheet_name=None)

# To iterate through the sheets

# create a csv for each of the worksheets in the excel file
for key in df.keys():
        df[key].to_csv('%s.csv' %key)

"""

Alternatives: 

df = pd.read_excel('file_name.xlsx', sheet_name=None)  
for key in df.keys(): 
    df[key].to_csv('{}.csv'.format(key))
    
OR

excel_file = 'data/excel_file.xlsx'
all_sheets = pd.read_excel(excel_file, sheet_name=None)
sheets = all_sheets.keys()

for sheet_name in sheets:
    sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
    sheet.to_csv("data/%s.csv" % sheet_name, index=False)

"""
# Create excel document
workbook= xlsxwriter.Workbook('master-tracking-data.xlsx')
# Create array to store each worksheet (definitely unnecessary)
worksheets_arr= []

#create worksheets for each day & add to array
day1 = workbook.add_worksheet('day1')
worksheets_arr = worksheets_arr.append(day1)

day2= workbook.add_worksheet('day2')
worksheets_arr = worksheets_arr.append(day2)

day3= workbook.add_worksheet('day3')
worksheets_arr = worksheets_arr.append(day3)

day4= workbook.add_worksheet('day4')
worksheets_arr = worksheets_arr.append(day4)

day5= workbook.add_worksheet('day5')
worksheets_arr = worksheets_arr.append(day5)

day6= workbook.add_worksheet('day6')
worksheets_arr = worksheets_arr.append(day6)

day7= workbook.add_worksheet('day7')
worksheets_arr = worksheets_arr.append(day7)

day8= workbook.add_worksheet('day8')
worksheets_arr = worksheets_arr.append(day8)

day9= workbook.add_worksheet('day9')
worksheets_arr = worksheets_arr.append(day9)

day10= workbook.add_worksheet('day10')
worksheets_arr = worksheets_arr.append(day10)

day11= workbook.add_worksheet('day11')
worksheets_arr = worksheets_arr.append(day11)

day12= workbook.add_worksheet('day12')
worksheets_arr = worksheets_arr.append(day12)

day13= workbook.add_worksheet('day13')
worksheets_arr = worksheets_arr.append(day13)

day14= workbook.add_worksheet('day14')
worksheets_arr = worksheets_arr.append(day14)

day15= workbook.add_worksheet('day15')
worksheets_arr = worksheets_arr.append(day15)

day16= workbook.add_worksheet('day16')
worksheets_arr = worksheets_arr.append(day16)

day17= workbook.add_worksheet('day17')
worksheets_arr = worksheets_arr.append(day17)

day18= workbook.add_worksheet('day18')
worksheets_arr = worksheets_arr.append(day18)

day19= workbook.add_worksheet('day19')
worksheets_arr = worksheets_arr.append(day19)

day20= workbook.add_worksheet('day20')
worksheets_arr = worksheets_arr.append(day20)

day21= workbook.add_worksheet('day21')
worksheets_arr = worksheets_arr.append(day21)

day22= workbook.add_worksheet('day22')
worksheets_arr = worksheets_arr.append(day22)

day23= workbook.add_worksheet('day23')
worksheets_arr = worksheets_arr.append(day23)

day24= workbook.add_worksheet('day24')
worksheets_arr = worksheets_arr.append(day24)

day25= workbook.add_worksheet('day25')
worksheets_arr = worksheets_arr.append(day25)

day26= workbook.add_worksheet('day26')
worksheets_arr = worksheets_arr.append(day26)

day27= workbook.add_worksheet('day27')
worksheets_arr = worksheets_arr.append(day27)

day28= workbook.add_worksheet('day28')
worksheets_arr = worksheets_arr.append(day28)

day29= workbook.add_worksheet('day29')
worksheets_arr = worksheets_arr.append(day29)

day30= workbook.add_worksheet('day30')
worksheets_arr = worksheets_arr.append(day30)

day31= workbook.add_worksheet('day31')
worksheets_arr = worksheets_arr.append(day31)

### Beginning of copy and paste ####

# What do we need to do?

"""
import csv made up of new plants. 
import csv made up of current plants and the tracking data associated with those plants.
Create new excel workbook.
Create worksheets for the workbook
Add current plants and tracking data from those plants to the appropriate worksheets
Format current plant and tracking data so that it is compatible with xlsxwriter.
Append new plant list to old plant list
Format new plants to make sure all data types are correct.
Create the main information sheet.
Export file.

FULL PLAN:
- Create a single exportable excel document on a single worksheet
- Make sure it works out and that all the data types are lined up.
- Import excel file, create all sheets as individual CSVs
- Import 'Day1' sheet into Pandas and figure out how to apply data into writeable xlswriter doc
- Once export of 'Day1' is tested, use the code in a loop to create all days. 


"""
##**********************************************##
##**********************************************##
""" Working with Adding and Updating Tables """
##**********************************************##
##**********************************************##

# Sample add table
# Need to convert contents of csv in the dataframe into this format
# Probably easy
data = [
    ['pb10', 'pancake breath #1', '','','',''],
    ['pb2', 'pancake breath #2', '', '','','']
]

# Determine the size of the table and create a string
table_size_string = 'B3:F7'
# Sample of adding a table to an individual worksheet. Can be used as test before looping.
for indiv_worksheet in worksheets_arr:
    indiv_worksheet.add_table(table_size_string, {'data': data, 'header_row': True, 'autofilter': False,
    'banded_columns': True, 'name': 'PlantTrackingData'})

# Trying to create a prototype method.
# Need to make sure I know what an autofilter is.
worksheets_arr[0].add_table(table_size_string, {'data': data, 'header_row': True, 'autofilter': False,
    'banded_columns': True, 'name': 'PlantTrackingData'})

""" PREPARING FORMATTING """
currency_format = workbook.add_format({'num_format': '$#,##0'})
wrap_format = workbook.add_format({'text_wrap': 1});
header_format = workbook.add_format()
header_format.set_font_size(14)

short_date_format  = workbook.add_format()
short_date_format.set_num_format('mm/dd/yy')
worksheet.write(5, 0, 36892.521, short_date_format)       # -> 01/01/01

dayofweek_format  = workbook.add_format()
dayofweek_format.set_num_format('ddd')
worksheet.write(5, 0, 36892.521, dayofweek_format)  # -> Tue

# THERE IS NO DEFAULT FORMULA FOR FOLIAR AND NUTES COLUMNS
# Probably need to manually add column headers so that we can enter a formula for each item in column
worksheets_arr[1].add_table(table_size_string, {'data': data,
                              'columns': [{'header': 'PlantID'},
                                          {'header': 'Plant'},
                                          {'header': 'Nutes',
                                           'formula': '=F5'},
                                          {'header': 'Foliar'},
                                          {'header': 'Loc'},
                                          {'header': 'Notes',
                                           'formula': '=F1DAODNAS'},
                                          {'header': 'Age',
                                           'formula': '=F1kASdASD'},
                                          {'header': 'FlowerDay',
                                           'formula': '=F1DAODNAS'},
                                          {'header': 'SeedStart',
                                           'formula': '=F1kASdASD'},
                                          {'header': 'Notes',
                                           'formula': '=F1DAODNAS'},
                                          {'header': 'Age',
                                           'formula': '=F1kASdASD'},
                                          {'header': 'FlowerDay',
                                           'formula': '=F1DAODNAS'},
                                          {'header': 'SeedStart'},
                                          {'header': 'FlowerStart'},
                                          {'header': "Group"},
                                          {'header': "Gender"},
                                          {'header': "Source"},
                                          {'header': "GroupNotes",
                                           'format': wrap_format,
                                           'formula': '=F5'}
                                          ]})
""" TO DO """
# Determine the length of the table based on the amount of plants in the dataframe.
# loop through each worksheet and use the add_table function.


# Some data we want to write to the worksheet.
expenses = (
    ['Rent', '2013-01-13', 1000],
    ['Gas', '2013-01-14', 100],
    ['Food', '2013-01-16', 300],
    ['Gym', '2013-01-20', 50],
)

# Start from the first cell below the headers.
row = 1
col = 0

for item, date_str, cost in (expenses):
    # Convert the date string into a datetime object.
    date = datetime.strptime(date_str, "%Y-%m-%d")

    worksheet.write_string(row, col, item)
    worksheet.write_datetime(row, col + 1, date, date_format)
    worksheet.write_number(row, col + 2, cost, money_format)
    row += 1

worksheet.write('A1','Hello world')

workbook.close()
