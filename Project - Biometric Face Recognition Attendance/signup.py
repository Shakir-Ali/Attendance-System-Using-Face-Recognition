import numpy as np
import cv2
import pandas as pd
import face_recognition as fc
import time
import random as rd
import smtplib
import xlrd

fcc=0
v=cv2.VideoCapture(0)
fd=cv2.CascadeClassifier(r"C:\Users\HP\AppData\Local\Programs\Python\Python36\Lib\site-packages\cv2\data\haarcascade_frontalface_alt2.xml")
def cap():
        ret,i=v.read()
        j=cv2.cvtColor(i,cv2.COLOR_BGR2GRAY)
        f=fd.detectMultiScale(j)
        if len(f)==1:
            for(x,y,w,h) in f:
                image=i[y:y+h,x:x+w].copy()
                fl=fc.face_locations(image)
                fcl=fc.face_encodings(image,fl)
                cv2.imshow('image',image)
                k= cv2.waitKey(5)
        
                return fcl
                break
        else:
            print("Face not Detected")
    
def genotp():
    ran=rd.random()
    otp=ran*10000
    return int(otp)

def enterdata():
    name=input("Enter Name: ")
    roll=input("Enter Roll No.: ")
    number=int(input("Enter MObile Number: "))
    email=input("Enter E-Mail: ")
    print("Hold Still The Camera will initialize to detect your face in few seconds")
    print("Name:",name,"\nRoll",roll,"\nNumber",number,"\nEmail",email)
    time.sleep(2)
    q=0
    while(q!=1):
        try:
            fcc=cap()
            if len(fcc) != 0:
                print("Successfully Entered Data OTP is sent to your email")    
                q=1
                return name,roll,number,email,fcc
        except:
            pass

def sendmail(email,otp):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login("Enter your mail","Enter your password")
    msg="Subject: OTP is "+str(otp)+" \nWelcome to our Institute!\nTo Complete Registration Please Enter the following OTP:"+str(otp)+" \nThank you for enrolling with us."
    server.sendmail("Leonardo Da Vinci",email,msg)



Data = pd.read_excel("Data.xlsx")
df = pd.DataFrame(Data)
Data1 = pd.read_excel("Attendance.xlsx")
df1 = pd.DataFrame(Data1)
name,roll,number,email,fcc = enterdata()
v.release()
otp=genotp()
sendmail(email,otp)
q=0
dataf=pd.DataFrame({"Name":[name],
                    "Roll":[roll],
                    "Number":[number],
                    "Email":[email],
                    "Encoding":list(fcc)})
dataf1=pd.DataFrame({'Name':[name],
                     'Email':[email]})
while(q!=1):
    eotp=int(input("Enter OTP"))
    if eotp==otp:
        q=1
        df=df.append(dataf,ignore_index=True,sort=False)
        df.to_excel("Data.xlsx",index=False)
        
        df1=df1.append(dataf1,ignore_index=True,sort=False)
        df1.to_excel("Attendance.xlsx",index=False)
        print("Success")
    else :
        print("Re-Enter OTP")
