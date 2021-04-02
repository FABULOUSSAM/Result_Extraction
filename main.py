
import pdf2image
from PIL import Image
import time
from pdf2image import convert_from_path
import os,os.path
import cv2
import numpy as np

PDF_PATH = "1T00718C.pdf"
DPI = 200
OUTPUT_FOLDER ="Page_Image"
FIRST_PAGE = 1
LAST_PAGE = 5
FORMAT = 'jpg'
THREAD_COUNT = 1
USERPWD = None
USE_CROPBOX = False
STRICT = False

def pdftopil():
    start_time = time.time()
    pil_images = pdf2image.convert_from_path(PDF_PATH, dpi=DPI,  first_page=FIRST_PAGE, last_page=LAST_PAGE, fmt=FORMAT, thread_count=THREAD_COUNT, userpw=USERPWD, use_cropbox=USE_CROPBOX, strict=STRICT)
    return pil_images
    
def save_images(pil_images):
    index = 1
    for image in pil_images:
        image.save("./Page_Image/page_" + str(index) + ".png")
        index += 1

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
        
count=1;

def croping(output,i):
    im = Image.open(r"./Page_Image/page_"+str(i)+".png")
   
    for j in range(len(output)-1):
        left = output[j][0]  
        right = output[j][2]
        top=output[j][1]
        bottom=output[j+1][1]
        im1 = im.crop((left, top, right, bottom))
        global count
        im1 = im1.save("./Crop_Image/crop_"+str(count)+".png")
        count=count+1;


if __name__ == "__main__":
    pil_images = pdftopil()
    save_images(pil_images)
    DIR = './Page_Image'
    x=len([name for name in os.listdir(DIR) if os.path.isfile(os.path.join(DIR, name))])
    line_detect(x)
