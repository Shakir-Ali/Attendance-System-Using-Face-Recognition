import numpy as np
import cv2
import pandas as pd
import face_recognition as fc
import datetime
import random as rd
import smtplib
import xlrd

x=[]
Data = pd.read_excel("Data.xlsx")
df = pd.DataFrame(Data)
y=datetime.datetime.now()
dates=y.strftime("%x")
Data1 = pd.read_excel("Attendance.xlsx")
df1 = pd.DataFrame(Data1)
if dates not in df1.columns:
    l=df1.shape[1]
    df1.insert(l,dates,"")
v=cv2.VideoCapture(0)
fd=cv2.CascadeClassifier(r"C:\Users\HP\AppData\Local\Programs\Python\Python36\Lib\site-packages\cv2\data\haarcascade_frontalface_alt2.xml")
img=[]

while(1):
    try:
        ret,i=v.read()
        j=cv2.cvtColor(i,cv2.COLOR_BGR2GRAY)
        f=fd.detectMultiScale(j)
        for(x,y,w,h) in f:
            img=i[y:y+h,x:x+w].copy()
        fm=fc.face_locations(img)
        fcm=fc.face_encodings(img,fm)
        if len(fcm) == 0:
            continue
        cv2.imshow('face',img)
        k=cv2.waitKey(5)
        if k==ord('q'):
            break
        elif k==ord('r'):
            
            length=len(df['Encoding'])
            for w in range(length):
                column=df['Encoding'][w]
                a=column.split(' ')
                newlist=[]
                r=[x for x in a if x!='']
                for x in r:
                    if '\n' in x:
                        f=float(x[:-2])
                        newlist.append(f)
                    elif '[' in x:
                        f=float(x[1:])
                        newlist.append(f)
                    elif ']' in x:
                        f=float(x[:-2])
                        newlist.append(f)
                    else:
                        newlist.append(float(x))
                new_array=np.array(newlist,dtype='float64')
                d=fc.compare_faces([new_array],fcm[0])
                if d==[False]:
                    continue
                else:
                    print(df['Name'][w])
                    z=datetime.datetime.now()
                    times=z.strftime("%X")
                    df1[dates][w]=times
                    df1.to_excel("Attendance.xlsx",index=False)
    except: 
        print("Unable To Detect Face")

                    
