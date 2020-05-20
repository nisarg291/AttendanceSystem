# import xlwt;
# from datetime import datetime;
# from xlrd import open_workbook;
# from xlwt import Workbook;
# from xlutils.copy import copy
# from pathlib import Path
#
# '''style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
#     num_format_str='#,##0.00')
# style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
# wb = xlwt.Workbook()
# ws = wb.add_sheet('A Test Sheet')
# ws.write(0, 0, 1234.56, style0)
# ws.write(1, 0, datetime.now(), style1)
# ws.write(2, 0, 1)
# ws.write(2, 1, 1)
# ws.write(2, 2, xlwt.Formula("A3+B3"))
# wb.save('example.xls')
# '''
#
# def details(filename, sheet,rollno,name,enrollment,course,section,sem,contact,email):
#     print("hii")
#     style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
#                          num_format_str='#,##0.00')
#     my_file = Path('firebase/attendance_files/' + filename +"_"+ str(course)+"_"+str(sem)+"_"+str(section)+'.xls');
#     if my_file.is_file():
#         rb = open_workbook('firebase/attendance_files/'+ filename +"_"+ str(course)+"_"+str(sem)+"_"+str(section)+'.xls');
#         book = copy(rb);
#         sh = book.get_sheet(0)
#     else:
#         book = xlwt.Workbook()
#         sh = book.add_sheet(sheet)
#     col1_name = 'rollno',
#     col2_name = 'Name',
#     col3_name='enroll_no',
#     col4_name='course',
#     col5_name='section',
#     col6_name='sem',
#     col7_name='contact',
#     col8_name='email'
#     sh.write(1, 0, col1_name, style0);
#     sh.write(1, 1, col2_name, style0);
#
# def output(filename, sheet,num, name, present):
#     my_file = Path('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
#     if my_file.is_file():
#         rb = open_workbook('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
#         book = copy(rb);
#         sh = book.get_sheet(0)
#         # file exists
#     else:
#         book = xlwt.Workbook()
#         sh = book.add_sheet(sheet)
#     style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
#                          num_format_str='#,##0.00')
#     style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
#
#     #variables = [x, y, z]
#     #x_desc = 'Display'
#     #y_desc = 'Dominance'
#     #z_desc = 'Test'
#     #desc = [x_desc, y_desc, z_desc]
#     sh.write(0,0,datetime.now().date(),style1);
#     col1_name = 'Name'
#     col2_name = 'Present'
#     sh.write(1,0,col1_name,style0);
#     sh.write(1, 1, col2_name,style0);
#     sh.write(num+1,0,name);
#     sh.write(num+1, 1, present);
#     #You may need to group the variables together
#     #for n, (v_desc, v) in enumerate(zip(desc, variables)):
#     fullname=filename+str(datetime.now().date())+'.xls';
#     book.save('firebase/attendance_files/'+fullname)
#     return fullname;
import xlwt;
from datetime import datetime;
from xlrd import open_workbook;
from xlwt import Workbook;
from xlutils.copy import copy
from pathlib import Path
import pandas as pd;
from pandas import ExcelWriter,ExcelFile
'''style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
ws.write(0, 0, 1234.56, style0)
ws.write(1, 0, datetime.now(), style1)
ws.write(2, 0, 1)
ws.write(2, 1, 1)
ws.write(2, 2, xlwt.Formula("A3+B3"))
wb.save('example.xls')
'''
def details(filename, sheet,rollno,name,enrollment,course,section,sem,contact,email):
    print("hii")
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                         num_format_str='#,##0.00')
    my_file = Path('firebase/attendance_files/'+str(filename)+""+ str(course)+"_"+str(sem)+"_"+str(section)+'.xls');
    my_file1=Path('firebase/attendance_files/Students.xls')
    if my_file.is_file():
        rb = open_workbook('firebase/attendance_files/'+ filename +""+ str(course)+"_"+str(sem)+"_"+str(section)+'.xls');
        book = copy(rb);
        sh = book.get_sheet(0)

    if my_file1.is_file():
        rb1 = open_workbook('firebase/attendance_files/Students.xls')
        book1 = copy(rb1);
        sh1 = book1.get_sheet(0)

    elif my_file.is_file()== False:
        book = xlwt.Workbook()
        sh = book.add_sheet(sheet)
    elif my_file1.is_file()== False:
        book1 = xlwt.Workbook()
        sh1 = book1.add_sheet(sheet)
    col1_name = 'Rollno',
    col2_name = 'Name',
    col3_name='Enrollno',
    col4_name='Course',
    col5_name='Sem',
    col6_name='Section',
    col7_name='Contact',
    col8_name='Email'
    rollno=int(rollno)
    sh.write(0, 0, col1_name, style0);
    sh.write(0, 1, col2_name, style0);
    sh.write(0, 2, col3_name, style0);
    sh.write(0, 3, col4_name, style0);
    sh.write(0, 4, col5_name, style0);
    sh.write(0, 5, col6_name, style0);
    sh.write(0, 6, col7_name, style0);
    sh.write(0, 7, col8_name, style0);
    sh.write(rollno,0,rollno);
    sh.write(rollno, 1,name);
    sh.write(rollno,2,enrollment);
    sh.write(rollno, 3,course);
    sh.write(rollno,4,sem);
    sh.write(rollno, 5,section);
    sh.write(rollno,6,contact);
    sh.write(rollno,7,email);
    sh1.write(0, 0, col1_name, style0);
    sh1.write(0, 1, col2_name, style0);
    sh1.write(0, 2, col3_name, style0);
    sh1.write(0, 3, col4_name, style0);
    sh1.write(0, 4, col5_name, style0);
    sh1.write(0, 5, col6_name, style0);
    sh1.write(0, 6, col7_name, style0);
    sh1.write(0, 7, col8_name, style0);
    sh1.write(rollno, 0, rollno);
    sh1.write(rollno, 1, name);
    sh1.write(rollno, 2, enrollment);
    sh1.write(rollno, 3, course);
    sh1.write(rollno, 4, sem);
    sh1.write(rollno, 5, section);
    sh1.write(rollno, 6, contact);
    sh1.write(rollno, 7, email);
    book1.save('firebase/attendance_files/Students.xls')
    fullname=str(filename)+""+ str(course)+"_"+str(sem)+"_"+str(section)+'.xls';
    book.save('firebase/attendance_files/'+fullname)
    return fullname;



def output(filename, sheet,rollno, name,enrollno,course,sem,section,contact,email,present):
    my_file = Path('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
    if my_file.is_file():
        rb = open_workbook('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
        book = copy(rb);
        sh = book.get_sheet(0)
        # file exists
    else:
        book = xlwt.Workbook()
        sh = book.add_sheet(sheet)
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

    #variables = [x, y, z]
    #x_desc = 'Display'
    #y_desc = 'Dominance'
    #z_desc = 'Test'
    #desc = [x_desc, y_desc, z_desc]
    df = pd.read_excel('firebase/attendance_files/Students.xls', sheet_name='class1')
    #print(rollno)
    col1_name='Rollno'
    col2_name = 'Name'
    col3_name = 'Enrollno',
    col4_name = 'Course',
    col5_name = 'Sem',
    col6_name = 'Section',
    col7_name = 'Contact',
    col8_name = 'Email'
    col9_name = 'Present'
    sh.write(0,0,col1_name,style0);
    sh.write(0, 1, col2_name,style0);
    sh.write(0, 2, col3_name, style0);
    sh.write(0, 3, col4_name, style0);
    sh.write(0, 4, col5_name, style0);
    sh.write(0, 5, col6_name, style0);
    sh.write(0, 6, col7_name, style0);
    sh.write(0, 7, col8_name, style0);
    sh.write(0, 8, col9_name, style0);
    sh.write(rollno,0,rollno);
    sh.write(rollno,1,name)
    sh.write(rollno,2,int(enrollno))
    sh.write(rollno,3,str(course))
    sh.write(rollno,4,str(sem))
    sh.write(rollno,5,str(section))
    sh.write(rollno,6,str(contact))
    sh.write(rollno,7,str(email))
    sh.write(rollno,8,present)
    #You may need to group the variables together
    #for n, (v_desc, v) in enumerate(zip(desc, variables)):
    fullname=filename+str(datetime.now().date())+'.xls';
    book.save('firebase/attendance_files/'+fullname)
    my();
    return fullname;

def my():
    df = pd.read_excel('firebase/attendance_files/Students.xls', sheet_name='class1')
    df = df.dropna()
    # print(df)
    df.index = [i for i in range(0, len(df))]
    df1 = pd.read_excel('firebase/attendance_files/attendance' + str(datetime.now().date()) + '.xls',
                        sheet_name='class1')
    df1 = df1.dropna()
    df1.index = [i for i in range(0, len(df1))]
    yes = []
    rollno = []
    for i in range(0, len(df1)):
        rollno.append(str(df1.loc[i, 'Rollno']))

    for i in range(0, len(df)):
        receiver = df.loc[i, 'Email']
        roll = str(df.loc[i, 'Rollno'])
        #print(type(roll))
        #print(type(rollno))
        if roll in rollno:
            msg = "you are attend today's lecs"
        else:
            my_file = Path('firebase/attendance_files/attendance' + str(datetime.now().date()) + '.xls');
            if my_file.is_file():
                rb = open_workbook('firebase/attendance_files/attendance' + str(datetime.now().date()) + '.xls');
                book = copy(rb);
                sh = book.get_sheet(0)
                roll = roll.split('.');
                roll = int(roll[0])
                print(roll)
                sh.write(roll, 0, roll);
                sh.write(roll, 1, str(df.loc[i, 'Name']))
                sh.write(roll, 2, int(df.loc[i, 'Enrollno']))
                sh.write(roll, 3, str(df.loc[i, 'Course']))
                sh.write(roll, 4, str(df.loc[i, 'Sem']))
                sh.write(roll, 5, str(df.loc[i, 'Section']))
                sh.write(roll, 6, str(df.loc[i, 'Contact']))
                sh.write(roll, 7, str(df.loc[i, 'Email']))
                sh.write(roll, 8, 'no')
                fullname = 'attendance' + str(datetime.now().date()) + '.xls';
                book.save('firebase/attendance_files/' + fullname)



