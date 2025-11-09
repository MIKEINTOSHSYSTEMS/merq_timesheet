# convert_to_ico.py
from PIL import Image
import os

def convert_png_to_ico():
    try:
        img = Image.open('src/merq.png')
        # Create multiple sizes for the icon
        sizes = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
        img.save('src/merq.ico', sizes=sizes)
        print("ICO file created successfully!")
    except Exception as e:
        print(f"Error creating ICO: {e}")

if __name__ == '__main__':
    convert_png_to_ico()