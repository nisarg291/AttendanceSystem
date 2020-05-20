from tkinter import *
from PIL import Image, ImageTk
import tkinter.messagebox as tmsg
import os,xlwrite
import cv2;
count=1;
def dataset_capture1(rollno,name):
    def assure_path_exists(path):
        dir = os.path.dirname(path)
        if not os.path.exists(dir):
            os.makedirs(dir)

    face_id = rollno
    user_name = name
    # Start capturing video
    vid_cam = cv2.VideoCapture(0)

    # Detect object in video stream using Haarcascade Frontal Face
    face_detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

    # Initialize sample face image
    count = 0

    assure_path_exists("dataset/")

    # Start looping
    while (True):

        # Capture video frame
        _, image_frame = vid_cam.read()

        # Convert frame to grayscale
        gray = cv2.cvtColor(image_frame, cv2.COLOR_BGR2GRAY)

        # Detect frames of different sizes, list of faces rectangles
        faces = face_detector.detectMultiScale(gray, 1.3, 5)

        # Loops for each faces
        for (x, y, w, h) in faces:
            # Crop the image frame into rectangle
            cv2.rectangle(image_frame, (x, y), (x + w, y + h), (255, 0, 0), 2)

            # Increment sample face image
            count += 1

            # Save the captured image into the datasets folder
            # only faces which are detected in video it is store atle khali faces je store thase clothes nai aave
            cv2.imwrite("dataset/." + user_name + "." + str(face_id) + '.' + str(count) + ".jpg",
                        gray[y:y + h, x:x + w])

            # Display the video frame, with bounded rectangle on the person's face
            cv2.imshow('frame', image_frame)

        # To stop taking video, press 'q' for at least 100ms
        if cv2.waitKey(100) & 0xFF == ord('q'):
            break

        # If image taken reach 100, stop taking video
        elif count >= 30:
            print("Successfully Captured")
            root1.destroy()
            break

    # Stop video
    vid_cam.release()

    # Close all started windows
    cv2.destroyAllWindows()

def cancel():
    pass
def update_details():
    print(uname.get(), ucourse.get(), ucls.get(), uenrollment.get(), ucontact.get(), uemail.get())
    filename=xlwrite.details('User', 'class1', rollno.get(), uname.get(),uenrollment.get(),ucourse.get(),ucls.get(),sem.get(), ucontact.get(), uemail.get());
    dataset_capture1(rollno.get(),uname.get());
root1 = Tk()
root1.geometry("400x500")
root1.maxsize(580,550)
root1.minsize(580,550)
image = Image.open("GUI/register.jpg")
image=image.resize((580,600), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(image)
root1.title("Auto Attendance..")
l = Label(image=photo)
f31=Frame(l,pady="25",padx="25")
l0=Label(root1,text="Update Details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")
l01=Label(root1,text="RollNo : ",font="lucida 10 bold").grid(column=0,row=1,pady="4")
rollno=StringVar()
e01=Entry(root1,textvariable=rollno,width="28").grid(column=1,row=1)
l1=Label(root1,text="Name : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
uname=StringVar()
e1=Entry(root1,textvariable=uname,width="28").grid(column=1,row=2)

l2=Label(root1,text="Enrollment No : ",font="lucida 10 bold").grid(column=0,row=3,pady="4")

uenrollment=StringVar()
e2=Entry(root1,textvariable=uenrollment,width="28").grid(column=1,row=3)
l3=Label(root1,text="Course : ",font="lucida 10 bold").grid(column=0,row=4,pady="4")

ucourse=StringVar()
e3=Entry(root1,textvariable=ucourse,width="28").grid(column=1,row=4)
l4=Label(root1,text="Class: ",font="lucida 10 bold").grid(column=0,row=5,pady="4")
ucls=StringVar()
e4=Entry(root1,textvariable=ucls,width="28").grid(column=1,row=5)

l5=Label(root1,text="Sem",font="lucida 10 bold").grid(column=0,row=6,pady="4")
sem = StringVar()
sem.set("1st sem") # default value
w1 = OptionMenu(root1,sem,"1st sem","2nd sem","3rd sem","4th sem","5th sem","6th sem","7th sem","8th sem").grid(column=1,row=6,pady="4")
l6=Label(root1,text="Contact No : ",font="lucida 10 bold").grid(column=0,row=7,pady="4")
ucontact=StringVar()
e6=Entry(root1,textvariable=ucontact,width="28").grid(column=1,row=7)
l7=Label(root1,text="Email : ",font="lucida 10 bold").grid(column=0,row=8,pady="4")

uemail=StringVar()
e7=Entry(root1,textvariable=uemail,width="28").grid(column=1,row=8)
btn=Button(root1,text="Submit",bg="green",fg="white",width="10",font="lucida 10 bold",command=update_details)
btn.grid(column=0,row=9,pady="20")
btn1=Button(root1,text="cancel",bg="green",fg="white",width="10",font="lucida 10 bold",command=cancel)
btn1.grid(column=1,row=9,pady="20")
f31.pack(pady="100")
root1.mainloop()