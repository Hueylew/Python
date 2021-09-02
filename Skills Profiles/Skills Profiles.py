#!/usr/bin/env python3
import os
import xlsxwriter

workbook = xlsxwriter.Workbook(r'/Users/adamlewis/BC Team Overall.xlsx')
workbook.close()

exec(open('/Users/adamlewis/Documents/Work/Python/ACT Team.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/Architecture.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/BA Skills.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/Domain Skills.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/IQH.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/Pure Insurance.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/Select.py').read())
exec(open('/Users/adamlewis/Documents/Work/Python/SSP Broker.py').read())

file_path = '/Users/adamlewis/BC Team Overall.xlsx'
os.system("open -a 'Microsoft Excel.app' '%s'" % file_path)