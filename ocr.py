import easyocr
import re
import pytesseract
reader = easyocr.Reader(['en']) 
result = reader.readtext('./Crop_Image/crop_2.png')
values=[]

text = pytesseract.image_to_string('./Crop_Image/crop_2.png')
for x in result:
    values.append(x[1])

dic={'seat_no':' ','name':' ', 'subject' :' ' }
temp=""
nos=[[],[],[]]
j=0
k=0
l=0
for x in range(0,len(values)):
    if x == 0:
        dic['seat_no']=values[x]
    elif x == 1 or x == 2 or x ==3 or x==4 :
        temp=temp+' '+(values[x])
        dic['name']=temp
    elif (x==5 or  x==6 ):
        dic['subject'] += ' '+values[x]
    elif  x==7 :
        if re.search('[a-zA-Z]',values[x]):
            dic['subject'] += ' '+values[x]
        else:
            nos[0].append(values[x])
            j=1
    elif  j==1  and  8 <= x <= 16 :
        nos[0].append(values[x])
        k=0
    elif j==0  and  8 <= x <= 17 :
        nos[0].append(values[x])
        k=1
    elif k==1  and  16<= x <= 24 :
        nos[1].append(values[x])
    elif k==0  and  17<= x <= 25 :
        nos[1].append(values[x])
print(text)