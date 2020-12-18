# import module from tkinter for UI
import os,xlwrite
from datetime import datetime;
from tkinter import *
import os
from pathlib import Path
from playsound import playsound
# creating instance of TK
root = Tk()
#root.geometry("580x550")
root.configure(background='white')


def function1():
    os.system("python gui2.py")
    os.system("python training_dataset.py")
def function2():
    os.system("python training_dataset.py")
def function3():
    os.system("python recognizer2.py")
    #playsound('sound.mp3')

def function5():
    os.system("python gui3.py")
    #os.startfile(os.getcwd() + "/developers/diet1frame1first.html");

def function6():
    root.destroy()

def attend():
    my_file = Path('firebase/attendance_files/attendance' + str(datetime.now().date()) + '.xls');
    if my_file.is_file():
    # for open fun4 attendance shit
        os.startfile(os.getcwd() + "/firebase/attendance_files/attendance" + str(datetime.now().date()) + '.xls')
    else:
        l3 = Label(root, text="attendance is not taken yet.Take attendance then try again", bg="orange", fg="blue", font="lucida 10 bold",width="60", pady="4").grid(columnspan=3, row=12, pady="5")
# stting title for the window
root.title("ATTENDANCE MANAGEMENT USING FACE RECOGNITION")
# creating a text label
root.geometry("510x510")
Label(root, text="ATTENDANCE MANAGEMENT SYSTEM", font=("times new roman", 20), fg="white", bg="maroon",height=2).grid(row=0, rowspan=4, columnspan=2, sticky=N + E + W + S,padx=5, pady=5)
# creating first button here fg is font color and bg for background color
Button(root, text="Registeration", font=("times new roman", 20), bg="#0D47A1", fg='white', command=function1).grid(row=5, columnspan=2, sticky=N + E + W + S, padx=5, pady=5)
# creating second button
Button(root, text="Train Dataset", font=("times new roman", 20), bg="#0D47A1", fg='white', command=function2).grid(row=6, columnspan=2, sticky=N + E + W + S, padx=5, pady=5)
# creating third button
Button(root, text="Recognize + Attendance", font=('times new roman', 20), bg="#0D47A1", fg="white",command=function3).grid(row=7, columnspan=2, sticky=N + E + W + S, padx=5, pady=5)
# creating attendance button
Button(root, text="Attendance Sheet", font=('times new roman', 20), bg="#0D47A1", fg="white", command=attend).grid(row=8, columnspan=2, sticky=N + E + W + S, padx=5, pady=5)
Button(root, text="Teacher's login page", font=('times new roman', 20), bg="#0D47A1", fg="white", command=function5).grid(row=9,columnspan=2,sticky=N + E + W + S,padx=5,pady=5)
Button(root, text="Exit", font=('times new roman', 20), bg="maroon", fg="white", command=function6).grid(row=10,columnspan=2,sticky=N + E + W + S,padx=5, pady=5)
root.mainloop()
