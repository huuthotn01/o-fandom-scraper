from PIL import Image
import os
import PIL
import glob

image = Image.open("test.png")
img_size = image.size

image = image.resize((img_size[0] * 3, img_size[1] * 3))
image.save("resized.png", optimize=True, quality=100)