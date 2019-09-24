#Playing with Openpyxl


import os
import openpyxl

dir = "C:\\Users\\v-ganmir\\Desktop\\"
filename = "Test1.xlsx"
file = dir + filename
#print(file)
os.chdir(dir)

wb = openpyxl.load_workbook(file)
sheet = wb.worksheets[0]
#print(sheet)
