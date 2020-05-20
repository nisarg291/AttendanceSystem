from datetime import datetime;
from tkinter import *
from PIL import Image, ImageTk
import os
import xlwt;
from datetime import datetime;
from xlrd import open_workbook;
from xlwt import Workbook;
from xlutils.copy import copy
from pathlib import Path
import pandas as pd;


from scripts import mail as m
def attend():
    # for open fun4 attendance shit
    os.startfile(os.getcwd() + "/firebase/attendance_files/attendance" + str(datetime.now().date()) + '.xls')
def month_attend_sheet():
    month=datetime.now().month
    year=datetime.now().year
    print(month)
    if month<10:
        month='0'+str(month)
    os.startfile(os.getcwd() + "/firebase/attendance_files/" + str(year)+'-'+str(month) + '.xls')
def month_attend():
    dd = 0;
    count = []
    mail=[]
    date1=str(datetime.now().date())
    date1=date1.split('-')
    year=date1[0]
    month=date1[1]
    day=date1[2]
    df1 = pd.read_excel('firebase/attendance_files/attendance' + str(datetime.now().date()) + '.xls',sheet_name='class1')
    df1 = df1.dropna()
    df1.index = [i for i in range(0, len(df1))]
    book = xlwt.Workbook()
    sh = book.add_sheet('class1')
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
    col1_name = 'Rollno'
    col2_name = 'Name'
    col3_name = 'Enrollno',
    col4_name = 'Course',
    col5_name = 'Sem',
    col6_name = 'Section',
    col7_name = 'Contact',
    col8_name = 'Email'
    col9_name = 'Present(%)'
    sh.write(0, 0, col1_name, style0)
    sh.write(0, 1, col2_name, style0)
    sh.write(0, 2, col3_name, style0)
    sh.write(0, 3, col4_name, style0)
    sh.write(0, 4, col5_name, style0)
    sh.write(0, 5, col6_name, style0)
    sh.write(0, 6, col7_name, style0)
    sh.write(0, 7, col8_name, style0)
    sh.write(0, 8, col9_name, style0)

    for z in range(0,221):
        count.append(0);
    i=day
    f=0;
    while i !='00':
        print(year,str(month),i)
        my_file = Path('firebase/attendance_files/attendance'+ str(year)+'-'+str(month)+'-'+ str(i) + '.xls');
        if my_file.is_file():
            f=f+1;
            df=pd.read_excel(my_file,sheet_name='class1')
            df=df.dropna()
            df.index=[i for i in range(0,len(df))]
            for j in range(0,len(df)):
                if f==1:
                    rollno3=int(df.loc[j,'Rollno'])
                    sh.write(j+1,0,int(rollno3))
                    sh.write(j+1,1,str(df.loc[j, 'Name']))
                    sh.write(j+1,2, int(df.loc[j, 'Enrollno']))
                    sh.write(j+1,3, str(df.loc[j, 'Course']))
                    sh.write(j+1,4, str(df.loc[j, 'Sem']))
                    sh.write(j+1,5, str(df.loc[j, 'Section']))
                    sh.write(j+1,6, str(df.loc[j, 'Contact']))
                    sh.write(j+1,7, str(df.loc[j, 'Email']))
                    if df.loc[j,'Present']=='yes':
                        count[j]=count[j]+1
                        print(count[j])
                elif df.loc[j,'Present']=='yes' and f>1:
                    count[j]=count[j]+1;
                    print(count[j])
            dd = dd + 1;
        i=int(i)
        i=i-1;
        if i<10:
            i='0'+str(i)
    for i in range(0,len(df1)):
        line=['hi',f'your percentage of attendance for {dd} days are {str(round(float((count[i]*100)/dd),4))}','nisargadalja24680@gmail.com',df1.loc[i,'Email']]
        print("mail is sent")
        #m.send_mail(line)
        print(f'your percentage of attendance for {dd} days are {str(round(float((count[i]*100)/dd),4))}')
        sh.write(i+1,8,round(float((count[i]*100)/dd),4))
        book.save('firebase/attendance_files/'+str(year)+'-'+str(month)+'.xls')

def mail():
    os.system('python accessStu.py')

root2 = Tk()
root2.geometry("400x500")
image = Image.open("GUI/register.jpg")
image=image.resize((580,600), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(image)
root2.configure(background='white')
root2.title("Auto Attendance..")

Button(root2, text="Attendance Sheet", font=('times new roman', 20), bg="#0D47A1", fg="white", command=attend).grid(row=2, columnspan=8, sticky=N + E + W + S, padx=5, pady=5)
Button(root2, text="send mail of Monthly Attendance", font=('times new roman', 20), bg="#0D47A1", fg="white", command=month_attend).grid(row=4, columnspan=8, sticky=N + E + W + S, padx=5, pady=5)
Button(root2, text="Send mail of today's attendance", font=('times new roman', 20), bg="#0D47A1", fg="white", command=mail).grid(row=6, columnspan=8, sticky=N + E + W + S, padx=5, pady=5)
Button(root2, text="open monthly attendance sheet", font=('times new roman', 20), bg="#0D47A1", fg="white", command=month_attend_sheet).grid(row=8, columnspan=8, sticky=N + E + W + S, padx=5, pady=5)
root2.mainloop()