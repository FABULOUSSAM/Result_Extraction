from PIL import Image
from pdf2image import convert_from_path
import os,os.path
import cv2
import numpy as np
def line_detect(x):
    for number in range(1,x+1):
        img = cv2.imread("./Page_Image/page_"+str(number)+".png", cv2.IMREAD_COLOR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray, 50, 200)
        lines = cv2.HoughLinesP(edges, 1, np.pi/180, 700,lines=6,minLineLength=400, maxLineGap=50)
        newline=[]
        for line in lines:
            x1, y1, x2, y2 = line[0]
            single_line=[x1,y1,x2,y2]
            newline.append(single_line)
        finalvalue=[newline[0]]
        output=[newline[0]];
        for i in range(1,len(newline)):
            flag=True;
            for j in range(0,len(output)):
                x,y=output[j][1]+2,output[j][1]-2;
                z=newline[i][1];
                if(z==x or z==y):
                    flag=False;
                    break;
            if(flag):
                output.append(newline[i]);
        n = len(output)
        for i in range(n):
            for j in range(0, n-i-1):
                if output[j][1] > output[j+1][1] :
                    output[j], output[j+1] = output[j+1], output[j]
        final_output=[]
        for i in range(0,len(output)):
            if(i>=2 and i<len(output)-2): 
                final_output.append(output[i])
        croping(final_output,number)

def croping(output,i):
    im = Image.open(r"./Page_Image/page_"+str(i)+".png")
    index=1
    for j in range(len(output)-1):
        left = output[j][0]  
        right = output[j][2]
        top=output[j][1]
        bottom=output[j+1][1]
        im1 = im.crop((left, top, right, bottom))
        im1 = im1.save("./Crop_Image/crop_" + str(i)+"_"+str(index)+".png")
        index=index+1


if __name__ == "__main__":
    DIR = './Page_Image'
    x=len([name for name in os.listdir(DIR) if os.path.isfile(os.path.join(DIR, name))])
    line_detect(x)