import os 
from pathlib import Path
import pandas as pd 

def get_project_root():
    """Returns project root file path"""
    return str(Path(__file__).parent)

project_root=get_project_root()
print(project_root)
# 'JAN--22',give your folder name
Month_folder=os.path.join(project_root,"JAN-22")
# to get the month and year of the date
# a,b,month=Month_folder.split('\\')
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
df=pd.read_excel(Excel_sheets[11],skiprows=4)
checking_for_sum=df['Unnamed: 2'].values




for comp in df.columns:
    if 'Component' in comp:
        component=comp
    else:
        pass

# get the Month
for month in df.columns:
    if 'month' in month.lower():
        x,Month=month.split(':')
# reading the columns from 6th rows
df_new=pd.read_excel(Excel_sheets[11],skiprows=5)
# forward filling
df_new.ffill(inplace=True)
# getting only those columns which contains sum
df_new=df_new[df_new['Unnamed: 2']=='sum']
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
# for key,value in Datas.items():
#     print(key,len(value))
# print(df_new.groupby(df_new['Date']).aggregate)
# print(df_new.iloc[:,1].values)

# final_df=pd.DataFrame(Datas)
# final_df['Date']=final_df['Date'].apply(lambda x: str(x) +'-'+Month)
# df_for_all.append(final_df)
# file_path=os.path.join(r'C:\Automation\Excel_automation-master\JAN-22','converted.xlsx')

# file_path=os.path.join(r'C:\Automation\Excel_automation-master\JAN-22','converted.xlsx')
# x=pd.concat(df_for_all,axis=0,ignore_index=True)
# # final_df.to_excel(file_path,index=False)
# x.to_excel(file_path,index=False)
# print(Datas)
# print(Excel_sheets)
name_excel=pd.read_excel(r"C:\Automation\Excel_automation-master\Castex Sapphire Converted Data_Nomeclture (18).xlsx",sheet_name = 4,header=None)
rename={}
chr=['.','/'] 
name_excel['new_col']=name_excel.iloc[:,1].apply(lambda X:(''.join([e for e in X if e not in chr]).lower()))
# for col in df.cols:
#   for i in range(len(name_excel)):
#     if ''.join([e for e in col if e not in chr]).lower() in name_excel.iloc[i:,2]:
#         rename.setdefault(col,name_excel.iloc[i:,0])
# for col in df.columns:

#  chr=['.','/'] 
#  for i in range(len(name_excel)):
#   if  (''.join([e for e in col if e not in chr])) in n
# print(str2.lower())
# for i in range(5) :  
#  print(name_excel.iloc[:,2])

mn=set()
# print(x.join('name'))

mn.update('z')
print(mn)
print(str({'abc'}))
from datetime import datetime

strdate="16-Oct-20" 
datetimeobj=datetime.strptime(strdate,"%d-%b-%y")
print(datetimeobj)
# df['date'] = pd.to_datetime(df['date']).dt.date
component_master=pd.read_excel(r'C:\Automation\Excel_automation-master\Sapphire Component Master.xlsx')
print(component_master.iloc[:,0])