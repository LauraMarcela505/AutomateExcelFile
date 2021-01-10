# puts multiple excel files as sheets in one workbook

import pandas as pd
import glob
import os

location = 'C:\\Users\\laura\\Desktop\\Sleep Data\\*.xls'
excel_files = glob.glob(location)

writer = pd.ExcelWriter('C:\\Users\\laura\\Desktop\\Combined Sleep Data\\multiple_sheets.xlsx')

for excel_file in excel_files:
    sheet = os.path.basename(excel_file)
    df1 = pd.read_excel(excel_file)
    df1.fillna(value='N/A', inplace=True)
    df1.to_excel(writer, sheet_name=sheet, index=False)

writer.save()


print(df1)