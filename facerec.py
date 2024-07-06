import face_recognition as fr
import cv2 as cv
import pandas as pd
from tkinter import Tk
import openpyxl
from tkinter.filedialog import askopenfilename
from datetime import date
from datetime import datetime
import os
import xlsxwriter 

Tk().withdraw()
load_image=askopenfilename()


target_image=fr.load_image_file(load_image)
target_enc=fr.face_encodings(target_image)

ls=[]

for i in range(0,50):
    ls.append(0)
    

def encode_faces(folder):
    list_ppl_enc=[]
    for filename in os.listdir(folder):
        known_image=fr.load_image_file(f'{folder}{filename}')
        known_enc=fr.face_encodings(known_image)[0]
        list_ppl_enc.append((known_enc,filename))
    return list_ppl_enc


def find_target_face():
    face_location=fr.face_locations(target_image)
    for person in encode_faces('people/'):
        encode_face=person[0]
        filename=person[1]
        is_tar_face=fr.compare_faces(encode_face, target_enc, tolerance=0.50)
        name=""
        for it in filename:
            if it=='.':
                break
            else:
                name+=it
        rn=0
        rn+=(ord(name[-1])-ord('0'))
        rn+=(10*(ord(name[-2])-ord('0')))
        for i in is_tar_face:
            if i==True:
                ls[rn-1]=1

        print(f'{is_tar_face}{filename}')
        
        if face_location:
            face_number=0
            for location in face_location:
                if is_tar_face[face_number]:
                    label=filename
                    create_frame(location,label)
                face_number+=1

def numOfDays(date1, date2):
    if date2 > date1:
        return (date2-date1).days
    else:
        return (date1-date2).days

def create_frame(location, label):
    top,right,bottom,left=location
    cv.rectangle(target_image,(left,top),(right,bottom),(255,0,0),2)
    cv.rectangle(target_image,(left,bottom+20),(right,bottom),(255,0,0),cv.FILLED)
    cv.putText(target_image, label, (left+3, bottom+14), cv.FONT_HERSHEY_DUPLEX, 0.4,(255,255,255),1)

def render_image():
    rgb_img=cv.cvtColor(target_image, cv.COLOR_BGR2RGB)
    cv.imshow('Face Recognition',rgb_img)
    cv.waitKey(0)

find_target_face()
render_image()
day=int(input("Enter date: "))
month=int(input("Enter month: "))
year=int(input("Enter year: "))
workbook = xlsxwriter.Workbook('sample.xlsx') 
start = date(2024,5,1)
today = date(year, month, day)
delta=numOfDays(today,start)
df1=pd.read_excel('Book2.xlsx')
df2=pd.DataFrame(ls, columns=[today])
with pd.ExcelWriter(
    "Book2.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="overlay",
) as writer:
    df1.to_excel(writer, sheet_name="Sheet1", index=False)
    df2.to_excel(writer, sheet_name="Sheet1", startcol=delta+1, index=False)\  