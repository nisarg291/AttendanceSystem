from tkinter import *
from PIL import Image, ImageTk
import tkinter.messagebox as tmsg
import os,time,xlwrite
import cv2;

root1 = Tk()
root1.geometry("400x500")
#root1.geometry("580x550")
#root1.maxsize(580,550)
#root1.minsize(580,550)
image = Image.open("GUI/register.jpg")
image=image.resize((580,600), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(image)
flag=0;
uname1=StringVar()
upass1=StringVar()
def login_details():
    global flag,uname1,upass1;
    if uname1.get()=='LJIET' and upass1.get()=='LJIET@1234':
        os.system('python gui4.py')
    else:
        print('uu')
        l3 = Label(root1, text="Username or password is incorrect", bg="orange", fg="blue", font="lucida 10 bold",width="35", pady="4").grid(columnspan=3, row=8, pady="15")
        flag=1;


root1.title("Auto Attendance..")
l0=Label(root1,text="Techer's login Page",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")

l1=Label(root1,text="UserName: ",font="lucida 10 bold").grid(column=0,row=2,pady="4")

e1=Entry(root1,textvariable=uname1,width="28").grid(column=1,row=2)
l2=Label(root1,text="Password: ",font="lucida 10 bold").grid(column=0,row=4,pady="4")

e2=Entry(root1,textvariable=upass1,width="28").grid(column=1,row=4)

btn=Button(root1,text="Submit",bg="green",fg="white",width="10",font="lucida 10 bold",command=login_details)
btn.grid(column=0,row=6,pady="20")
root1.mainloop()