from PIL import Image, ImageDraw, ImageColor
import numpy as np
import xlsxwriter
import PIL

def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % rgb

filename = 'images/cat.jpg'
image = Image.open(filename)  # open image file
draw = ImageDraw.Draw(image)  # open tool for drawing
width = image.size[0]  # get width
height = image.size[1]  # get heigth
pix = image.load()  # load values of pixels

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()
for x in range(height):
    for y in range(width):
        cell_format = workbook.add_format()
        cell_format.set_pattern(1)  # This is optional when using a solid fill.
        r = pix[y, x][0]  # get red
        g = pix[y, x][1]  # get green
        b = pix[y, x][2]  # get blue
        cell_format.set_bg_color(rgb_to_hex((r, g, b)))
        worksheet.set_column(0, width, 2)  # Width of columns from 1 to width set to 2 size.
        worksheet.write(x, y,' ', cell_format)

workbook.close()
