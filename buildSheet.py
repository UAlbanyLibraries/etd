import os
import csv
import xlwt

workingDir = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"

outputFile = os.path.join(workingDir, "output.csv")

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')

rowCount = 0
with open(outputFile, newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        
        print (row[0])
        
        ws.write(rowCount, 0, row[0])
        ws.write(rowCount, 1, row[1])
        
        rowCount += 1
        



wb.save(os.path.join(workingDir, 'example.xls'))