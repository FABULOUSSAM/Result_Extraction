from PIL import Image
from PIL import Image, ImageEnhance

image = Image.open('./Crop_Image/crop_7.png')
new_image = image.resize((2250, 225))
# sharpness = ImageEnhance.Sharpness(new_image)
# sharpness.enhance(2).save('changes.png')
contrast = ImageEnhance.Contrast(new_image)
contrast.enhance(1.5).save('changes.png')


print(image.mode)
print(image.size) # Output: (1920, 1280)
print(new_image.size) # Output: (400, 400)