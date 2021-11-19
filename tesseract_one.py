import pytesseract
import cv2
import numpy as np
import re
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\Shubham\AppData\Local\Tesseract-OCR\tesseract.exe'

img = cv2.imread('./Crop_Image/crop_4.png')
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
blurred = cv2.GaussianBlur(gray, (5, 5), 0)
#edged = cv2.Canny(blurred, 50, 200, 255)
custom_config = r'--oem 3 --psm 6'
val=(pytesseract.image_to_string(blurred, config=custom_config))
values=[]
j=''
for x in val:
    j+=x
    if ' ' in x:
        values.append(j)
        j=""
for x in range(0,len(values)):
    if '\n' in values[x]:
        m=values[x].split('\n')
        values[x]= m[0]
        values.insert(x+1,m[1])
dic={'seat_no':' ','name':' ' ,'status':''}
temp=""
nos=[[],[],[]]
sgpi={}
k=0
j=p=q=r=0
list2count=0
list3count=0
list4count=0
list5count=0
for x in range(0,len(values)):
    if x == 0:
        dic['seat_no']=values[x]
    elif values[x].find('(') == -1 and x <= 6 and not(re.search('[0-9]',values[x])):
        temp=temp+' '+(values[x])
        dic['name']=temp
    elif values[x].find('(') != -1 and  x <=9 :
        while values[x].find(')') == -1:
            x+=1
        k=x
    else:
        while j != 10:
            k+=1 
            nos[0].append(values[k])
            j+=1
        list2count=k
        while p!= 10:
            list2count+=1 
            nos[1].append(values[list2count])
            p+=1
            dic['status']=values[list2count+1]
        if(dic['status']=="F"):
            list3count = list2count + 1   
            while q!= 10:
                list3count+=1 
                nos[2].append(values[list3count])
                q+=1
        else:
            list3count = list2count + 1
            while q!= 10:
                list3count+=1 
                if  (re.search("[a-zA-Z]",values[list3count])): 
                    dic['status']+= " " + values[list3count]
                else:
                    nos[2].append(values[list3count])
                    q+=1
        list4count = list3count 
        while True:
            if values[list4count].find(')') != -1  or values[list4count].find('SGPI')!=-1:
                r=list4count
                break
            else:
                list4count+=1 
       
        if (re.search("[a-zA-Z]",values[r-1])) and values[r-1].find('(')!=-1:
            sgpi['result'] =  values[r-4] + ' ' +values[r-3] + ' ' +values[r-2] 
        else:
            sgpi['result'] =  values[r-3] + ' ' +values[r-2] + ' ' +values[r-1] 
        list5count = list4count + 1
        while True:
            list5count+=1
            if values[list5count].find('CGPI') == -1:
                if re.search("[a-zA-Z]",values[list5count]):
                    sgpi[values[list5count]]=""
                elif re.search("[0-9]",values[list5count]):
                    sgpi[values[list5count-1]] = values[list5count]
            else:
                break
print(values)
print("dictionary")
print(dic)
print(nos)
print(sgpi)
