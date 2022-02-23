# -*- coding: utf-8 -*-


import os 
from pathlib import Path
from tkinter import E
import pandas as pd 
from datetime import datetime


def get_project_root():
    """Returns project root file path"""
    return str(Path(__file__).parent)

project_root=get_project_root()
print(project_root)
# 'JAN--22',give your folder name
Month_folder=os.path.join(project_root,"JAN-22")
# to get the month and year of the date
# a,b,Month1=Month_folder.split('\\')
Month1="JAN-22"
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
# reading excel file 
df_for_all=[]
for i in range(len(Excel_sheets)):

 df=pd.read_excel(Excel_sheets[i],skiprows=4)
 checking_for_sum=df['Unnamed: 2'].values
 if 'sum' not in checking_for_sum:
    for comp in df.columns:
     if 'Component' in comp:
        component=comp
     else:
         pass
        
        #  component='Component ID-.'
# get the Month
    for month in df.columns:
      if 'month' in month.lower():
        x,Month=month.split(':')
# reading the columns from 6th rows
    df_new=pd.read_excel(Excel_sheets[i],skiprows=5)
# forward filling
    #df_new.loc[:,'Date'].ffill()
    df_new.loc[:,'Date'].ffill(inplace=True)
    df_new.drop_duplicates(subset = ['Date'], keep = 'first', inplace = True) 
    # print(df_new)
    Datas={"Date":[]}  
# getting types of rejections 
    rej_col=df_new.columns
    for i in rej_col:
      if type(i)==int:
       Datas['Date'].append(i)
       key,value=component.split('-')
       Datas.setdefault(key,[])
       Datas[key].append(value)
    for i in range(len(df_new)):
      for j in range(3,int(len(Datas['Date'])+3)):
       Datas.setdefault(df_new.iloc[i,1],[])
       Datas[df_new.iloc[i,1]].append(df_new.iloc[i,j])
    final_df=pd.DataFrame(Datas)
    final_df['Date']=final_df['Date'].apply(lambda x: str(x) +'-'+str(Month1))
    df_for_all.append(final_df)  
 else:
# get the componet name 
  for comp in df.columns:
    if 'Component' in comp:
        component=comp
    else:
        # component='Component ID-.'
        pass

# get the Month
  for month in df.columns:
    if 'month' in month.lower():
        x,Month=month.split(':')
# reading the columns from 6th rows
  df_new=pd.read_excel(Excel_sheets[i],skiprows=5)
# forward filling
  df_new.ffill(inplace=True)
# getting only those columns which contains sum
  df_new=df_new[df_new['Unnamed: 2']=='sum']
  df_new.drop_duplicates(subset = ['Date'], keep = 'first', inplace = True) 
# for second format 
  Datas={"Date":[]}
  df_second_format=[]
# getting types of rejections 
  rej_col=df_new.columns
  for i in rej_col:
    if type(i)==int:
     Datas['Date'].append(i)
     key,value=component.split('-')
     Datas.setdefault(key,[])
     Datas[key].append(value)
  for i in range(len(df_new)):
    for j in range(3,int(len(Datas['Date'])+3)):
     Datas.setdefault(df_new.iloc[i,1],[])
     Datas[df_new.iloc[i,1]].append(df_new.iloc[i,j])
  final_df=pd.DataFrame(Datas)
  final_df['Date']=final_df['Date'].apply(lambda x: str(x) +'-'+str(Month1))
  df_for_all.append(final_df)
# file_path=os.path.join(r'C:\Automation\Excel_automation-master\JAN-22','converted.xlsx')

file_path=os.path.join(r'C:\Automation\JAN-22','converted.xlsx')
df_of_all_excel=pd.concat(df_for_all,axis=0,ignore_index=True)
# final_df.to_excel(file_path,index=False)
# x.to_excel(file_path,index=False)
# print(Excel_sheets)

name_excel=pd.read_excel(r"C:\Automation\Castex Sapphire Converted Data_Nomeclture (18).xlsx",sheet_name = 4,header=None)
rename={}
chr=['.','/'] 
name_excel['new_col']=name_excel.iloc[:,1].apply(lambda X:(''.join([e for e in X if e not in chr]).lower()))


for col in df_of_all_excel.columns:
  for i in range(len(name_excel)):
    if (''.join([e for e in col if e not in chr])).lower() in name_excel.iloc[i,2]:
        rename.setdefault(col,[])
        rename[col].append(name_excel.iloc[i,0])
df_of_all_excel.rename(columns=rename,inplace=True)
df_of_all_excel.columns = [x[0] if type(x)==list else x for x in df_of_all_excel.columns]
#df_of_all_excel.rename(columns={'Mould Broken':'Mould Broken (no)','W/Jkt Br.':'W/GR/Wrong Grind (no)','Wrong Grind':'W/GR/Wrong Grind (no)','M/ Scab.':'Mould Scab (no)','"Core Lift',"Cor"},inplace=True)
#df_of_all_excel.columns =df_of_all_excel.columns.astype(str)
# removing unwanted columns 
df_of_all_excel.drop(['CHECK','% Rejection','Total Checked','Checked'],axis=1,inplace=True)
# summing the columns with same name 
# df_of_all_excel.fillna(0).groupby(df_of_all_excel.columns, axis=1).sum()
# df_of_all_excel.groupby(level=0, axis=1).sum()
# print(df_of_all_excel.info)
df_of_all_excel.to_excel(file_path,index=False)
# last_cols=list(set(df_of_all_excel.columns))
req_file=df_of_all_excel.groupby(level=0, axis=1).sum()
req_file.insert(0,'Date',req_file.pop('Date'))
req_file.insert(1,'Component ID ',req_file.pop('Component ID '))
req_file.insert(2,'Production',req_file.pop('Production'))
req_file['Date']=req_file['Date'].apply(lambda x: datetime.strptime(x,"%d-%b-%y"))

# adding component weight,and cavity
compoent_wt=[]
cavity=[]
component_master=pd.read_excel(r'C:\Automation\Sapphire Component Master.xlsx')
req_file_1=pd.merge(req_file,component_master,on='Component ID',how='left')

# req_file=req_file[last_cols] 'Component ID '
req_file_1.to_excel(file_path,index=False)








# start working on excel sheet
# must put data only as a true to get the values not the formula
# wb=openpyxl.load_workbook(Excel_sheets[1],data_only=True)
# # sheets=wb.get_sheet_names()
# # it will give the list of sheet names
# sheets=wb.sheetnames
# # we are interested in reading first sheet only so we will get working sheet by following way
# sheet=wb.get_sheet_by_name(sheets[0])
# print(sheet['K8'].value)
# # unfreeze the cells
# sheet.freeze_panes=None
# # unmerging cells
# for items in sorted(sheet.merged_cell_ranges):
#     print(items)
#     sheet.unmerge_cells(str(items) )

# # get highest rows and columns
# rows=sheet.max_row
# columns=sheet.max_column
# print(rows,columns)
# print(sheet['AI6'].value)

#  # get the componet id and and month

# for i in range(1,columns):
#   if 'Component' in  str(sheet.cell(row=5,column=i).value) :
#       component=str(sheet.cell(row=5,column=i).value)
# # get the month name 
# for i in range(1,columns):
#     if 'Month' in str(sheet.cell(row=5,column=i).value) :
#         Month=str(sheet.cell(row=5,column=i).value)
#         a,month=Month.split(':')

      

# Datas={}
# for j in range(4,columns):
#     # column j+1 is added to remove the last None value 
#     if sheet.cell(row=6,column=j+1).value !=None:
       
#        Datas.setdefault("Date",[])
#        Datas['Date'].append((sheet.cell(row=6,column=j).value))
#        # get the component IDs
#        key,value=component.split('-')
#        Datas.setdefault(key,[])
#        Datas[key].append(value)


# # data frame with two features
# df1=pd.DataFrame(Datas)
# for i in range(7,rows):
#     for j in range(4,4+len(df1)):
#         if sheet.cell(row=i,column=2).value==None:
#             pass
#         else:
#          Datas.setdefault(sheet.cell(row=i,column=2).value,[])# to get the key value
#          if sheet.cell(row=i,column=j).value==None:
#               Datas[sheet.cell(row=i,column=2).value].append('')
#          else:
#           Datas[sheet.cell(row=i,column=2).value].append(sheet.cell(row=i,column=j).value)
# #ws.Columns.EntireColumn.Hidden=False
# # sheet.Columns.EntireColumn.Hidden=False
# # wb.save('freezeExample.xlsx')
# df=pd.DataFrame(Datas)
# df['Date']=df['Date'].apply(lambda x: str(x) +'-'+month)
# file_path=os.path.join(r'C:\Automation\JAN-22','converted.xlsx')
# df.to_excel(file_path,index=False)


# start working on excel sheet
# must put data only as a true to get the values not the formula
# wb=openpyxl.load_workbook(Excel_sheets[1],data_only=True)
# # sheets=wb.get_sheet_names()
# # it will give the list of sheet names
# sheets=wb.sheetnames
# # we are interested in reading first sheet only so we will get working sheet by following way
# sheet=wb.get_sheet_by_name(sheets[0])
# print(sheet['K8'].value)
# # unfreeze the cells
# sheet.freeze_panes=None
# # unmerging cells
# for items in sorted(sheet.merged_cell_ranges):
#     print(items)
#     sheet.unmerge_cells(str(items) )

# # get highest rows and columns
# rows=sheet.max_row
# columns=sheet.max_column
# print(rows,columns)
# print(sheet['AI6'].value)

#  # get the componet id and and month

# for i in range(1,columns):
#   if 'Component' in  str(sheet.cell(row=5,column=i).value) :
#       component=str(sheet.cell(row=5,column=i).value)
# # get the month name 
# for i in range(1,columns):
#     if 'Month' in str(sheet.cell(row=5,column=i).value) :
#         Month=str(sheet.cell(row=5,column=i).value)
#         a,month=Month.split(':')

      

# Datas={}
# for j in range(4,columns):
#     # column j+1 is added to remove the last None value 
#     if sheet.cell(row=6,column=j+1).value !=None:
       
#        Datas.setdefault("Date",[])
#        Datas['Date'].append((sheet.cell(row=6,column=j).value))
#        # get the component IDs
#        key,value=component.split('-')
#        Datas.setdefault(key,[])
#        Datas[key].append(value)


# # data frame with two features
# df1=pd.DataFrame(Datas)
# for i in range(7,rows):
#     for j in range(4,4+len(df1)):
#         if sheet.cell(row=i,column=2).value==None:
#             pass
#         else:
#          Datas.setdefault(sheet.cell(row=i,column=2).value,[])# to get the key value
#          if sheet.cell(row=i,column=j).value==None:
#               Datas[sheet.cell(row=i,column=2).value].append('')
#          else:
#           Datas[sheet.cell(row=i,column=2).value].append(sheet.cell(row=i,column=j).value)
# #ws.Columns.EntireColumn.Hidden=False
# # sheet.Columns.EntireColumn.Hidden=False
# # wb.save('freezeExample.xlsx')
# df=pd.DataFrame(Datas)
# df['Date']=df['Date'].apply(lambda x: str(x) +'-'+month)
# file_path=os.path.join(r'C:\Automation\JAN-22','converted.xlsx')
# df.to_excel(file_path,index=False)



