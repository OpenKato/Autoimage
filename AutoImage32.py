# -*- coding: utf-8 -*-
from PIL import Image
import glob, os
from openpyxl import load_workbook


size = 190, 190

allowed_extensions = set(['png', 'jpg', 'gif',"PNG","JPG","GIF","jpeg","JPEG"])

for infile in glob.glob("*.jpg"):
    file, ext = os.path.splitext(infile)
    im = Image.open(infile)
    im.thumbnail(size)
    im.save(file + ".r.JPG", "JPEG",quality=100, optimize=True)

wb = load_workbook(filename = '1.xlsx')
ws = wb.active

from openpyxl.drawing.image import Image

img1 = Image('N.r.JPG')
img2 = Image('F.r.JPG')
img3 = Image('L.r.JPG')
img4 = Image('R.r.JPG')
img5 = Image('B.r.JPG')

ws.add_image(img1, 'D15')
ws.add_image(img2, 'L15')
ws.add_image(img3, 'D27')
ws.add_image(img4, 'L27')
ws.add_image(img5, 'D39')
wb.save('2.xlsx')
