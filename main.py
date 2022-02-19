# -*- coding: utf-8 -*-


import os 
from pathlib import Path
from tkinter import E
import pandas as pd 
import xlrd
import glob
import openpyxl
from xls2xlsx import XLS2XLSX

def get_project_root():
    """Returns project root file path"""
    return str(Path(__file__).parent)

project_root=get_project_root()
print(project_root)
Month_folder=os.path.join(project_root,"JAN-22")
X=[]
for subdir,dir,file in os.walk(Month_folder):
    for i in dir:
        X.append(os.path.join(Month_folder,i))

# All the excel

Excel_sheets=[]
for i in X:
  
 for a,b,c in os.walk(i):
    for n in c:
     Excel_sheets.append((os.path.join(i,n)))

 
# start working on excel sheet
# must put data only as a true to get the values not the formula
wb=openpyxl.load_workbook(Excel_sheets[0],data_only=True)
# sheets=wb.get_sheet_names()
# it will give the list of sheet names
sheets=wb.sheetnames
# we are interested in reading first sheet only so we will get working sheet by following way
sheet=wb.get_sheet_by_name(sheets[0])
print(sheet['K8'].value)
# unfreeze the cells
sheet.freeze_panes=None
# unmerging cells
for items in sorted(sheet.merged_cell_ranges):
    print(items)
    sheet.unmerge_cells(str(items) )

# get highest rows and columns
rows=sheet.max_row
columns=sheet.max_column
print(rows,columns)
print(sheet['AI6'].value)
Datas={'Date':[]}
for j in range(10,columns):
    if sheet.cell(row=10,column=j).value !=None:
       Datas["Date"].append(sheet.cell(row=6,column=j).value)
print(Datas)
#ws.Columns.EntireColumn.Hidden=False
sheet.Columns.EntireColumn.Hidden=False
wb.save('freezeExample.xlsx')

