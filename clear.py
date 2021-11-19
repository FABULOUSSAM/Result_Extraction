import os
 
dir = './Crop_Image'
for f in os.listdir(dir):
    os.remove(os.path.join(dir, f))

dir = './Header_crop'
for f in os.listdir(dir):
    os.remove(os.path.join(dir, f))

dir = './Page_Image'
for f in os.listdir(dir):
    os.remove(os.path.join(dir, f))
