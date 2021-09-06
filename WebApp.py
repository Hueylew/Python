from matplotlib import colors
import openpyxl
import os
import xlsxwriter
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image
import numpy as np
import glob
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.utils import get_column_interval
import re
import fpdf
from fpdf import FPDF

def load_workbook_range(range_string, ws):
	col_start, col_end = re.findall("[A-Z]+", range_string)
	
	data_rows = []
	for row in ws[range_string]:
		data_rows.append([cell.value for cell in row])
		
	return pd.DataFrame(data_rows, columns=get_column_interval(col_start, col_end))

# Read all spreadsheets in the named folder below
all_data = pd.DataFrame()
f = ('/Users/adamlewis/Documents/Work/Business Consultancy/Skills and Development/Combined Skills Profiles 2021/Billy May - Skills Profile.xlsx')

# Get the persons name
	
from openpyxl import load_workbook
wb = load_workbook(f, data_only=True)
sh = wb["Details"]
PersonName = sh["B2"].value
PersonPos = sh["B3"].value
PersonScore = sh["G22"].value
PersonTarget = sh["G23"].value
PersonVariance = sh["G24"].value
PersonRole = sh["B3"].value
WorkIn = sh["B4"].value
Location = sh["B5"].value
YearsGI = sh["B6"].value
YearsBA = sh["B7"].value
YearsSSP = sh["B8"].value

#Create the data frame
from itertools import islice
ws = wb["Details"]
df_Scores = load_workbook_range('B22:K24', ws)
df_Scores = df_Scores.rename(columns={'B': '','C': 'Select','D': 'Pure Ins', 'E': 'SSP Broker', 'F': 'IQH', 'G': 'ACT', 'H': 'Common Components', 'I': 'Domain', 'J': 'BA', 'K': 'Architecture'})
print(df_Scores)
	
#Create the data frame
from itertools import islice
ws = wb["ACT"]
df_BA = load_workbook_range('A1:C77', ws)
	
# Create and Populate the score column based on answer provided
df_BA = df_BA.fillna(value=0)
df_BA["C"].replace({"Basic": 1}, inplace=True)
df_BA["C"].replace({"Intermediate": 2}, inplace=True)
df_BA["C"].replace({"Advanced": 3}, inplace=True)
df_BA["C"].replace({"SME": 4}, inplace=True)
df_BA = df_BA.rename(columns={'A': 'Category'})
df_BA = df_BA.rename(columns={'B': 'Topic'})
df_BA = df_BA.rename(columns={'C': 'Skill Level'})
df_BA.at[2:7,'Category']='Analytical thinking and problem solving'
df_BA.at[9:11,'Category']='Communication skills'
df_BA.at[13:15,'Category']='Interaction skills'
df_BA.at[17:21,'Category']='KCM Author'
df_BA.at[23:29,'Category']='Pure In-House Document production'
df_BA.at[31:40,'Category']='Pre Rules Author'
df_BA.at[42:50,'Category']='Pure Product Dev - Back office'
df_BA.at[52:59,'Category']='Pure Product Dev - PB2 Basic'
df_BA.at[61:64,'Category']='Pure Product Dev - PB2 Advanced'
df_BA.at[66:70,'Category']='Programming languages'
df_BA.at[72,'Category']='Integration'
df_BA.at[74:75,'Category']='BDX'

all_data = df_BA

all_data.fillna('', inplace=True)
new_header = all_data.iloc[0] 
all_data = all_data[1:] 
all_data.columns = new_header
dfCategory = all_data.groupby('Category')['Skill Level'].mean().sort_values(ascending=False)
dfCategory.columns = ['Category']
print(dfCategory)
list(dfCategory.columns)

figScores = px.bar(df_Scores, x='ACT', y='')
figCategory = px.bar(dfCategory, x='Category', y='Skill Level')

df = all_data
fig = px.bar(df, x="Category", y="Skill Level", orientation='v')
# fig.show()
		
all_data.fillna('', inplace=True)
new_header = all_data.iloc[0] 
all_data = all_data[1:] 
all_data.columns = ['Category', 'Topic', 'Skill Level']
dfCategory = all_data.groupby('Category')['Skill Level'].mean().sort_values(ascending=False)

st.set_page_config(page_title='Skills Profiles')
st.header(' ACT Skills Profiles Results 2021')
st.subheader(PersonName)
st.dataframe(df_Scores)
st.dataframe(all_data)
st.dataframe(dfCategory)
st.plotly_chart(fig)
st.plotly_chart(figScores)
st.plotly_chart(figCategory)