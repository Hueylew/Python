import openpyxl
import os
import xlwings as xw
from openpyxl import workbook, load_workbook

excel_files = ('/Users/adamlewis/BC Team Overall.xlsx')
wb = load_workbook(excel_files)
ws = wb['Architecture']
ws.delete_rows(3)
ws.delete_rows(2)
wb.save(excel_files)

sheet = xw.Book('/Users/adamlewis/BC Team Overall.xlsx').sheets('Architecture')
sheet.range('A3:A36').value = "Underlying Competencies"
sheet.range('A38:A63').value = "Qualifications / Experience"
sheet.range('A65:A67').value = "Infrastructure"
sheet.range('A69:A73').value = "Accountability"
sheet.range('A75:A76').value = "Preferences"
sheet.range('B3:B8').value = "Analytical thinking and problem solving"
sheet.range('B10:B12').value = "Communication skills"
sheet.range('B14:B16').value = "Interaction skills"
sheet.range('B18:B24').value = "Tools"
sheet.range('B26:B29').value = "Architectural"
sheet.range('B31:B36').value = "Process"
sheet.range('B38:B42').value = "Architecture"
sheet.range('B44:B49').value = "Progamming"
sheet.range('B51:B54').value = "Databases"
sheet.range('B56').value = "Methodology"
sheet.range('B58:B63').value = "Leadership"
sheet.range('B65').value = "Cloud"

wb.save(excel_files)
os.system("open -a 'Microsoft Excel.app' '%s'" % excel_files)