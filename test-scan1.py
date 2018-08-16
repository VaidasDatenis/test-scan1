from os import listdir
from os.path import isfile, join
import pandas as pd
import pyzbar.pyzbar as pyzbar
import numpy as np
import cv2
import glob
import os
import xlwt
from tempfile import TemporaryFile

#File path to images
folder = "C:\\Users\\tager88\\Desktop\\test-scan1\\"
onlyfiles = [img for img in listdir(folder) if isfile(join(folder, img))]
barcode = []
name = []

def decode(im) : 
  # Find barcodes and QR codes
  decodedObjects = pyzbar.decode(im)
 
  # Print results
  for obj in decodedObjects:
     
    print('Type : ', obj.type)
    print('Number : ', obj.data)
    print('Image : ', img.strip(folder), '\n')
    
    try:
      barcode.append(int(obj.data))
    except Exception as e:
      print(e)

    try:
      name.append(str(img.strip(folder)))
    except Exception as e:
      print(e)
    
  return decodedObjects

# Display barcode and QR code location  
def display(im, decodedObjects):
 
  # Loop over all decoded objects
  for decodedObject in decodedObjects: 
    points = decodedObject.polygon
 
    # If the points do not form a quad, find convex hull
    if len(points) > 4 : 
      hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
      hull = list(map(tuple, np.squeeze(hull)))
    else : 
      hull = points;
     
    # Number of points in the convex hull
    n = len(hull)
 
    # Draw the convext hull
    for j in range(0,n):
      cv2.line(im, hull[j], hull[ (j+1) % n], (255,0,0), 3)
 
  # Display results 
  #cv2.imshow("Results", im);
  #cv2.waitKey(0);
 
   
# Main 
if __name__ == '__main__':
 
  # Read image
  for img in glob.glob(folder + "*.jpg"):
    im = cv2.imread(img)

    decodedObjects = decode(im)
    display(im, decodedObjects)
    
#Converting lists to DataFrames
df_onlyfiles = pd.DataFrame(onlyfiles)
df_barcode = pd.DataFrame(barcode)
df_name = pd.DataFrame(name)
frames = pd.concat([df_name, df_barcode], axis = 1, join = 'inner')
frames.columns = ['Image Name', 'Barcode nr.']

#Saves the file to Excel file.
writer = pd.ExcelWriter('barkodai.xlsx', engine = 'xlsxwriter')
frames.to_excel(writer, sheet_name = 'Barkodai')
workbook = writer.book
worksheet = writer.sheets['Barkodai']

#Formating Excel file's columns
format1 = workbook.add_format({'num_format': '#'})
worksheet.set_column('B:B', 18)
worksheet.set_column('C:C', 20, format1)
writer.save()
