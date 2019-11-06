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
#extract cleaned 'dipl' column as string 
    cleaned = base_name + '_CLEANDIPL.csv' 
    col = pandas.read_csv(output, error_bad_lines=False, sep='\t') 
    cleaned_col = col[['dipl']].to_csv(cleaned, index = None, header=True)
#1)bring edited 'clean' column back to origine
    wb = openpyxl.load_workbook(excel)
    ws = wb.active
    dipl_col_index = None
    clean_col_index = None

    for e, col in enumerate(ws.iter_cols()):     
        if col[0].value == 'clean':
            clean_col_index = e+1
    if clean_col_index is not None:
        ws.delete_cols(clean_col_index)
    for e, col in enumerate(ws.iter_cols()):
        if col[0].value == 'dipl':
            dipl_col_index = e+1
    print('dipl_col_index=', dipl_col_index)
    ws.insert_cols(dipl_col_index+1)
    #find merged cells
    for r in ws.merged_cells.ranges:
        #find unmerged cells
        ws.unmerge_cells(r)

    
    with open(cleaned) as f:
        reader = csv.reader(f, delimiter=';')
        for i, row in enumerate(reader):
                for j, cell in enumerate(row): 
                    ws.cell(row=i+1, column=dipl_col_index+1).value = 'ben'
                    

    wb.save('output.xlsx') #wb.save(basename + '.xlsx') 
    
   # wb.save(output[:-5] + '.xlsx') #hier setzte ich später base_name statt output[... ein


    
