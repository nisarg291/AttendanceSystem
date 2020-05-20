import xlwt;
from datetime import datetime;
from xlrd import open_workbook;
from xlwt import Workbook;
from xlutils.copy import copy
from pathlib import Path
from scripts import mail as m
import pandas as pd;
df=pd.read_excel('firebase/attendance_files/Students.xls', sheet_name='class1')
df=df.dropna()
    #print(df)
df.index=[i for i in range(0,len(df))]
df1=pd.read_excel('firebase/attendance_files/attendance'+str(datetime.now().date())+'.xls', sheet_name='class1')
df1=df1.dropna()
df1.index=[i for i in range(0,len(df1))]
yes=[]
rollno=[]
for i in range(0,len(df1)):
    rollno.append(str(df1.loc[i,'Rollno']))

for i in range(0,len(df)):
    receiver=str(df1.loc[i,'Email'])
    p=df1.loc[i,'Present']
    print(p)
    if p=='yes':
        msg = "you are attended today's lecs"
    else:
        msg="you are not attended today's lecs"
    sender='nisargadalja24680@gmail.com'
    line=['hi',msg,sender,receiver]
    #m.send_mail(line)



