#!/usr/bin/env python3

import glob
import os
import pandas 
import subprocess
import sys
#import xlsxwriter
import openpyxl
import csv

#convert lb and clean column to .txt
excel_files = glob.glob(os.getcwd() + '/*.xlsx') 

for excel in excel_files:
    base_name = excel.split('.')[0]
    convex = base_name + '_CONV.txt'
#apply clean-skript_V3.py
    df = pandas.read_excel(excel) 
    df[['lb','dipl']].to_csv(convex, index=False, sep='\t') 
    output = convex[:-9] + '_CLEANED.csv'
    subprocess.call(['python3', 'clean-skript_V3.py', convex, output])
#extract cleaned 'dipl' column as string and list
    cleaned = base_name + '_CLEANDIPL.csv' 
    col = pandas.read_csv(output, error_bad_lines=False, sep='\t') 
    cleaned_col = col[['dipl']].to_csv(cleaned, index = None, header=True)
#bring edited 'clean' column back to origine
    wb = openpyxl.load_workbook(excel)
    sheet= wb.active

    cleanlist = pandas.read_csv(cleaned)
    li = cleanlist.values.tolist()
    print(type(li))
    
    with open(cleaned) as f:
        reader = csv.reader(f, delimiter= '\t')
        for column in reader:
            sheet['B1'] = li #need list not str)
            #sheet.append(column)
    wb.save(output[:-5] + '.xlsx') #hier setzte ich später base_name statt output[... ein

 
   


# next step: bring col to reader         
# TODO: read csv file (output of clean script) line by line and set value for clean cell
# TODO: how and when merge clean cells??? E.g. by using the empty lines

#wb = openpyxl.load_workbook(excel)
    #sheet= wb.active

    #with open(output) as f:
        #reader = csv.reader(f, delimiter= '\t')
        #for row in reader:
            #sheet.append(row)

    
   # wb.save(output[:-5] + '.xlsx') #hier setzte ich später base_name statt output[... ein


    
