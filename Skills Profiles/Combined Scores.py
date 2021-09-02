import openpyxl
import os
import xlrd
import pandas as pd
import glob

#glob.glob('/Users/adamlewis/Documents/Work/Business Consultancy/Skills and Development/Combined Skills Profiles 2020/*.xlsx')
all_data = pd.DataFrame()
for f in glob.glob('/Users/adamlewis/Documents/Work/Business Consultancy/Skills and Development/Combined Skills Profiles 2021/*.xlsx'):
    # Get persons name
    book = xlrd.open_workbook(f)
    first_sheet = book.sheet_by_index(0)
    PersonName = first_sheet.cell(1,1).value
    PersonName = str(PersonName)
    # Get score data
    df_Score = pd.read_excel(f, sheet_name='Details', skiprows=12, nrows=5, usecols="B:J")
    df_Score.fillna('', inplace=True)
    # Get target scores and format
    df_Targets = pd.read_excel(f, sheet_name="Details", skiprows=19, nrows=3, usecols="B:J")
    df_Targets = df_Targets.rename(columns={'Unnamed: 1': 'Overall', 'Unnamed: 2': 'SIaaS', 'Unnamed: 3': 'Pure Ins', 'Unnamed: 4': 'SSP Broker', 'Unnamed: 5': 'IQH', 'Unnamed: 6': 'ACT', 'Unnamed: 7': 'Common Components', 'Unnamed: 8': 'Domain', 'Unnamed: 9': 'BA'})
    df_Targets["SIaaS"] = 100 * df_Targets["SIaaS"]
    df_Targets["Pure Ins"] = 100 * df_Targets["Pure Ins"]
    df_Targets["SSP Broker"] = 100 * df_Targets["SSP Broker"]
    df_Targets["IQH"] = 100 * df_Targets["IQH"]
    df_Targets["ACT"] = 100 * df_Targets["ACT"]
    df_Targets["Common Components"] = 100 * df_Targets["Common Components"]
    df_Targets["Domain"] = 100 * df_Targets["Domain"]
    df_Targets["BA"] = 100 * df_Targets["BA"]
    # Combine Scores and targets into one dataframe
    df_Overall = df_Score.append([df_Targets], ignore_index=True, sort=False)
    df_Overall.drop('Overall', axis=1, inplace=True)
    df_Overall.iat[5,0] = 'Score'
    df_Overall.iat[6,0] = 'Target'
    df_Overall.iat[7,0] = 'Variance'
    df_Overall.insert(0, "Name", [PersonName, PersonName, PersonName, PersonName, PersonName, PersonName, PersonName, PersonName])
    all_data = all_data.append(df_Overall,ignore_index=True)
    pd.options.display.float_format = "{:,.0f}".format
    print(df_Overall)
    
# Save Result
all_data.to_excel("output.xlsx", index = False)
