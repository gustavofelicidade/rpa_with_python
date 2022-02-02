# =====================================================================
#  this libraries below is to handle rpa and xlsx.
# =====================================================================

import rpa as r
import openpyxl
from openpyxl import Workbook  # to work with excel files
from openpyxl.worksheet.table import TableStyleInfo

# =====================================================================
# this libraries below is to test OCR.
# =====================================================================

import numpy as np
import matplotlib.pyplot as plt
import cv2  # opencv
import pytesseract
from pytesseract import Output
# import imutils
import re

# =====================================================================
# # RUN THIS SCRIPT
# =====================================================================


# Simple example of RPA with Python:
# Here I will demonstrate common RPA tasks
# Such as:
# Fill a form and submit
# Extract Text from image
# and the next, with pdf, desktop and so on...


# =============================================================
#  First we gonna to create an
#  Excel file to populate with
#  data for the form
# =============================================================


wb = Workbook()
sheet = wb.active  # activate the current sheet.

# this way we write the cells
sheet["A1"] = "*First name:"
sheet["A2"] = "*Last name:"
sheet["A3"] = "*Email:"
sheet["A4"] = "*Company:"
sheet["A5"] = "City:"
sheet["A6"] = "*Country:"
sheet["A7"] = "Phone:"
sheet["A8"] = "Please tell us how can we help you"

# the cell values will be inputs to the form


sheet["B1"] = "Gustavo"
sheet["B2"] = "Felicidade"
sheet["B3"] = "gustavofelicidade@outlook.com"
sheet["B4"] = "Upwork"
sheet["B5"] = "Rio de Janeiro"
sheet["B6"] = "Brazil"
sheet["B7"] = "+5521981176975"
sheet["B8"] = "accept the deal"

wb.save(filename="form.xlsx")  # Saving excel file.

# =============================================================
# Now we will use rpa library to submit the online form#
# =============================================================


url = 'https://clickdimensions.com/form/default.html'
# web.open(url)    # <- use webbrowser to open
r.init(visual_automation=True)
r.url('https://clickdimensions.com/form/default.html')

# This is an example of address of useful Xpath that we collect on the browser.
# <input id="txtFirstName" size="30" name="txtFirstName" zn_id="91" xpath="1">

# Note that we pass the value from the .xlsx file created

r.type('//input[@id = "txtFirstName"]', sheet["B1"].value)
r.type('//input[@id = "txtLastName"]', sheet['B2'].value)
r.type('//input[@id = "txtFormEmail"]', sheet['B3'].value)
r.type('//input[@id="txtCompany"]', sheet['B4'].value)
r.type('//input[@id="txtCity"]', sheet['B5'].value)
r.type('//input[@id="txtPhone"]', sheet['B6'].value)
r.type('//tbody/tr[10]/td[2]', sheet['B7'].value)

r.click('//input[@id = "txtFirstName"]')

# =============================================================
# Now we gonna work with Image and text
# I will do an example of OCR use.
# =============================================================

# later, we can improve this lines for few lines using list comprehensions.
sheet["D1"] = "Month"
sheet["D2"] = "Jan"
sheet["D3"] = "Feb"
sheet["D4"] = "Mar"
sheet["D5"] = "Apr"
sheet["D6"] = "May"
sheet["D7"] = "Jun"
sheet["D8"] = "Jul"
sheet["D9"] = "Ago"
sheet["D10"] = "Sep"
sheet["D11"] = "Oct:"
sheet["D12"] = "Nov:"
sheet["D13"] = "Dez"

sheet["E1"] = "Value"
sheet["E2"] = "2400"
sheet["E3"] = "2300"
sheet["E4"] = "2700"
sheet["E5"] = "2700"
sheet["E6"] = "2950"
sheet["E7"] = "2500"
sheet["E8"] = "2110"
sheet["E9"] = "2840"
sheet["E10"] = "3200"
sheet["E11"] = "3500"
sheet["E12"] = "3000"
sheet["E13"] = "3000"

wb.save(filename="form.xlsx")  # Saving excel file.

# define a table style
mediumStyle = TableStyleInfo(name='TableStyleMedium2',
                             showRowStripes=True)
# create a table
table = openpyxl.worksheet.table.Table(ref='D1:E13',
                                       displayName='Payments',

                                       tableStyleInfo=mediumStyle)
# add the table to the worksheet
sheet.add_table(table)

wb.save(filename="form.xlsx")  # Saving excel file.'''

#  Get a print screen
r.wait(6.6)
r.snap('page', 'results.png')
#r.echo(r.read('results.png'))  # reading image with python RPA func
r.close()

# Image processing

img = cv2.imread('results.png')
image = cv2.imread('results.png')
# Convert the image to gray scale
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# Adding custom options (see options with 'tesseract --help')
custom_config = r'--oem 3 --psm 6'
pytesseract.image_to_string(img, config=custom_config)  # convert image to string text


# get grayscale image
def get_grayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)


# noise removal
def remove_noise(image):
    return cv2.medianBlur(image, 5)


# thresholding
def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]


# dilation
def dilate(image):
    kernel = np.ones((5, 5), np.uint8)
    return cv2.dilate(image, kernel, iterations=1)


# erosion
def erode(image):
    kernel = np.ones((5, 5), np.uint8)
    return cv2.erode(image, kernel, iterations=1)


# opening - erosion followed by dilation
def opening(image):
    kernel = np.ones((5, 5), np.uint8)
    return cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel)


# canny edge detection
def canny(image):
    return cv2.Canny(image, 100, 200)


# skew correction
def deskew(image):
    coords = np.column_stack(np.where(image > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)

    else:
        angle = -angle
        (h, w) = image.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotated


# template matching
def match_template(image, template):
    return cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)


#  After preprocessing with the following code
'''
image = cv2.imread('results.png')

gray = get_grayscale(image)
thresh = thresholding(gray)
opening = opening(gray)
canny = canny(gray)

cv2.imshow('grayscale', gray)

cv2.waitKey(0)
cv2.imshow('thresholding', thresh)
cv2.waitKey(0)
cv2.imshow('opening', opening)
cv2.waitKey(0)
cv2.imshow('canny-edge', canny)
cv2.waitKey(0)
cv2.destroyAllWindows()'''

# Plot original image

image = cv2.imread('results.png')
b, g, r = cv2.split(image)
rgb_img = cv2.merge([r, g, b])
plt.imshow(rgb_img)

plt.title('ORIGINAL IMAGE')
plt.show(block=False)
plt.pause(3)
plt.close



# Preprocess image

gray = get_grayscale(image)
thresh = thresholding(gray)
opening = opening(gray)
canny = canny(gray)
images = {'gray': gray,
          'thresh': thresh,
          'opening': opening,
          'canny': canny}

# Plot images after preprocessing

fig = plt.figure(figsize=(13, 13))
ax = []

rows = 2
columns = 2
keys = list(images.keys())
for i in range(rows*columns):
    ax.append(fig.add_subplot(rows, columns, i+1))
    ax[-1].set_title('original - ' + keys[i])
    plt.imshow(images[keys[i]], cmap='gray')

# =============================================================
# Get OCR output using Pytesseract
# =============================================================

custom_config = r'--oem 3 --psm 6'
print('-----------------------------------------')
print('TESSERACT OUTPUT --> ORIGINAL IMAGE')
print('-----------------------------------------')
print(pytesseract.image_to_string(image, config=custom_config))
print('\n-----------------------------------------')
print('TESSERACT OUTPUT --> THRESHOLDED IMAGE')
print('-----------------------------------------')
print(pytesseract.image_to_string(image, config=custom_config))
print('\n-----------------------------------------')
print('TESSERACT OUTPUT --> OPENED IMAGE')
print('-----------------------------------------')
print(pytesseract.image_to_string(image, config=custom_config))
print('\n-----------------------------------------')
print('TESSERACT OUTPUT --> CANNY EDGE IMAGE')
print('-----------------------------------------')
print(pytesseract.image_to_string(image, config=custom_config))

# If you don't have tesseract executable in your PATH, include the following:
path = r"D:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = path


# =============================================================
# Bounding box information using Pytesseract
# =============================================================

# Plot original image

image = cv2.imread('results.png')
b, g, r = cv2.split(image)
rgb_img = cv2.merge([r, g, b])

plt.figure(figsize=(16, 12))
plt.imshow(rgb_img)

plt.title('SAMPLE INVOICE IMAGE')

plt.show(block=False)
plt.pause(3)
plt.close()



# =============================================================
# Plot character boxes on image using pytesseract.image_to_boxes() function
# =============================================================

image = cv2.imread('results.png')
h, w, c = image.shape
boxes = pytesseract.image_to_boxes(image)
for b in boxes.splitlines():
    b = b.split(' ')
    image = cv2.rectangle(image, (int(b[1]), h - int(b[2])), (int(b[3]), h - int(b[4])), (0, 255, 0), 2)

b, g, r = cv2.split(image)
rgb_img = cv2.merge([r, g, b])

plt.figure(figsize=(16, 12))
plt.imshow(rgb_img)
plt.title('SAMPLE INVOICE WITH CHARACTER LEVEL BOXES')

plt.show(block=False)
plt.pause(3)
plt.close()

