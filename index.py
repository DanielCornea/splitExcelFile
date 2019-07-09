import os
import pandas as pd 
import datetime as dt
import win32com.client as win32
from  win32com.client import Dispatch


# read the excel file with items 
items = pd.read_excel('SampleData.XLSX')

# strip white spaces in column name 
int_list = []
for column in items.columns: 
    int_list.append(column.replace(" ", ""))
items.columns = int_list


# removing duplicates 
iters = list(set(items['Rep']))


# create the output folder 
os.mkdir("output")

# split the excel file 
for iter in iters : 
    aux = items[items['Rep'] == iter]
    title = "output" + "\\" + str(iter) + ".xlsx"  
    aux.to_excel(title, index = False)

# get the curent directory 
dirpath = os.getcwd()


print("current directory is : " + dirpath)

# Autofit columns 
for fisier in os.listdir(dirpath + "\\output") :
    title = dirpath + "\\output\\" + fisier    
    excel = Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(title)
    #Activate second sheet
    excel.Worksheets(1).Activate()
    #Autofit column in active sheet
    excel.ActiveSheet.Columns.AutoFit()
    print("processing " + fisier + " now...")
    wb.Save()
    wb.Close()


# success message
print("All good, files split :) ")