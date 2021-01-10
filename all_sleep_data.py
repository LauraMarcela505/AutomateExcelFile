# concatenates all excel files in one worksheet

import pandas as pd
import glob

location = 'C:\\Users\\laura\\Desktop\\Sleep Data\\*.xls'
excel_files = glob.glob(location)

pd.set_option('display.max_rows', 91)

df1 = pd.DataFrame()

for excel_file in excel_files:
    df2 = pd.read_excel(excel_file)
    df1 = pd.concat([df1, df2], ignore_index=True)

df1.fillna(value="N/A", inplace=True)
df1.to_excel('C:\\Users\\laura\\Desktop\\Combined Sleep Data\\all_sleep_data.xlsx')


print(df1)