import cv2
import numpy as np
import xlwrite
import firebase.firebase_ini as fire
import time
from playsound import playsound
import pyttsx3 #pip install pyttsx3
import speech_recognition as sr

import pandas as pd
from pandas import ExcelWriter,ExcelFile;
import sys
import os,cv2;
import numpy as np
from PIL import Image;
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

start = time.time()
period = 8
face_cas = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
cap = cv2.VideoCapture(0);
recognizer = cv2.face.LBPHFaceRecognizer_create();
recognizer.read('trainer/trainer.yml');
df = pd.read_excel('firebase/attendance_files/Students.xls', sheet_name='class1')
#df['Name']
len1=df['Rollno'].value_counts()
print(len1.index)
print(df['Rollno'][0])
flag = 0;
id = 0;
filename = 'filename';
dict = {
    'item1': 0
}
#font = cv2.InitFont(cv2.cv.CV_FONT_HERSHEY_SIMPLEX, 5, 1, 0, 1, 1)
font = cv2.FONT_HERSHEY_SIMPLEX
flg=0;
l=[]
while True:
    match1 = 0;
    ret, img = cap.read();
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY);
    faces = face_cas.detectMultiScale(gray, 1.3, 5);
    #df = pd.read_excel('firebase/attendance_files/UserC.S._6th sem_A.xls', sheetname='class1')
    for (x, y, w, h) in faces:
        roi_gray = gray[y:y + h, x:x + w]
        cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 2);
        id,conf = recognizer.predict(roi_gray)
        cv2.imshow('frame', img);
        #cv2.imshow('frame', img);
        if (conf < 500):
            confidence = int(100 * (1 - (conf) / 300))
            print(confidence)
            if confidence >=78:
                filt=df['Rollno']==id
                name=df[filt]['Name'].values
                Enrollno = df[filt]['Enrollno'].values
                Course = df[filt]['Course'].values
                Sem = df[filt]['Sem'].values
                Section = df[filt]['Section'].values
                Contact = df[filt]['Contact'].values
                Email = df[filt]['Email'].values
                cv2.putText(img, str(id) + " " + str(conf), (x, y - 10), font, 0.55, (120, 255, 120), 1)
                filename = xlwrite.output('attendance','class1',id, name[0],Enrollno[0],Course[0],Sem[0],Section[0],Contact[0],Email[0],'yes');
                dict[str(id)] = str(id);
                match=1;
                match1=match1+1;
                #if flg==len(faces):
                speak('Thanks'+name[0]+ 'your attandnce successfully done')
                flg=flg+1;
                break;
            else:
                cv2.putText(img,"face is not recognize",(x, y - 10), font, 0.55, (120, 255, 120), 1)
                speak('face is not recognize you are not register please register then try again')
        else:
            id = 'Unknown, can not recognize'
            speak('you are not register please register then try agian')
            flag = flag + 1
            break

        #cv2.cv.PutText(cv2.cv.fromarray(img),str(id),(x,y+h),font,(0,0,255));

    if match1==len(faces) and flag>=5:
        break;
    if flag == 10:
        print("Transaction Blocked")
        break;
    if time.time() > start + period:
        break;
    if cv2.waitKey(100) & 0xFF == ord('q'):
        break;

cap.release();
cv2.destroyAllWindows();


# recognizer = cv2.face.LBPHFaceRecognizer_create()
# detector= cv2.CascadeClassifier("haarcascade_frontalface_default.xml");
# Ids=[]
# Names=[]
# def getImagesAndLabels(path):
#     global Ids, Names
#     #get the path of all the files in the folder
#     imagePaths=[os.path.join(path,f) for f in os.listdir(path)]
#     #create empth face list
#     faceSamples=[]
#     #create empty ID list
#
#     #now looping through all the image paths and loading the Ids and the images
#     for imagePath in imagePaths:
#         #loading the image and converting it to gray scale
#         pilImage=Image.open(imagePath).convert('L')
#         #Now we are converting the PIL(pillow) image into numpy array
#         imageNp=np.array(pilImage,'uint8')
#         #getting the Id from the image
#         Id=int(os.path.split(imagePath)[-1].split(".")[2])
#         Name=(os.path.split(imagePath)[-1].split(".")[1])
#         # extract the face from the training image sample
#         faces=detector.detectMultiScale(imageNp)
#         #If a face is there then append that in the list as well as Id of it
#         for (x,y,w,h) in faces:
#             faceSamples.append(imageNp[y:y+h,x:x+w])
#             Ids.append(Id)
#             Names.append(Name)
#     return faceSamples,Ids
#
# faces,Ids = getImagesAndLabels('dataSet')
# s = recognizer.train(faces, np.array(Ids))
# print("Successfully trained")
# recognizer.write('trainer/trainer.yml')
# print(Ids)
# print(Names)
# print(len(Ids))
# print(len(Names))