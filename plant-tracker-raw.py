import pandas as pd
import xlsxwriter as xlsxwriter

# df.keys()

"""
# create new directory to export 
for key in df.keys():
        df[key].to_csv('%s.csv' %key)
"""

""" Amount of rows to drop from the beginning. """
rows_to_drop = 19

""" Amount of columns in the table. """
max_columns = 14

""" path to excel file """
excel_file = r"feb-tracking-data.xlsx"

df = pd.read_excel(excel_file, sheet_name=None)

""" Extract feb2 sheet from dataframe and convert to dictionary """
feb1_dict = df['feb2']
#df2 = pd.DataFrame.from_dict(feb1_dict).transpose()
#df2 = df2.drop(df2.iloc[:, 14:], inplace = True, axis = 1)

""" Convert dictionary into a dataframe """
df2 = pd.DataFrame.from_dict(feb1_dict)
#df2 = df2.drop([:20])
# df3 = df2.drop(df2.iloc[:, 14:], inplace = True, axis = 1)

""" Drop columns not inlcuded in the table. """
df2.drop(df2.iloc[:, max_columns:], inplace = True, axis = 1)

""" Drop the first x amount of rows. """
df3 = df2.iloc[rows_to_drop:, :]

#df2 = df2.dropna(how='all')
# drop all NaN if it's under plant ID
#df2 = df2.dropna(subset=['PlantID'])

""" Rename columns based on the 0th row """
df4 = df3.rename(columns=df3.iloc[0])

""" Drop the first row because that is the row we're using as headers.
    We might actually want to include this since we are exporting a list. """
df5 = df4.iloc[1:, :]

""" Drop all rows that have a PlantID as NaN"""
df6 = df5.dropna(subset=['PlantID'])

""" Create a list of lists for inserting into a table. """
df6_list = [df6[i].tolist() for i in df6.columns]
