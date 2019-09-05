#!/usr/bin/env python
# -*- coding: cp936 -*-
import easygui as g
import win32com.client as win32com
import os
import os.path
import csv
import re
import codecs,sys
import tkinter
import win32timezone
import win32timezone
import datetime
import time
now_1 = str(time.strftime('%Y'))
def ph_QC_Chart_2018():    
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s'%now_1, title="Select files",filetypes=[("csvfile", "*.csv")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./QC Chart _pH_66-01-2018-006.xlsx'))        
        maxcolumn=1
        while excel.Sheets("Data 1").Cells(2,maxcolumn).Value is not None:
            maxcolumn+=1  
        ph_date = csv.reader(open('%s'%path,'r'))
        print ph_date
        k=0
        for each in ph_date:
            if ('CC' in each):
                now= re.findall('\d{4}-\d{2}-\d{2}',each[0])
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).Value=now[0]
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).NumberFormat= "yyyy/m/d"
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).Value=each[5]
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).NumberFormat="0.000"            
                k+=1
def ph_QC_Chart_2014():    
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s'%now_1, title="Select files",filetypes=[("csvfile", "*.csv")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./QC Chart _pH_66-01-2014-015.xlsx'))        
        maxcolumn=1
        while excel.Sheets("Data 1").Cells(2,maxcolumn).Value is not None:
            maxcolumn+=1  
        ph_date = csv.reader(open('%s'%path,'r'))
        print ph_date
        k=0
        for each in ph_date:
            if ('CC' in each):
                now= re.findall('\d{4}-\d{2}-\d{2}',each[0])
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).Value=now[0]
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).NumberFormat= "yyyy/m/d"
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).Value=each[5]
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).NumberFormat="0.000"            
                k+=1                
option=g.ccbox('which test item',title='pH QC data input', choices=('2014','2018'))
if option==1:
        print '2014'
        ph_QC_Chart_2014()
else:
        print '2018'
        ph_QC_Chart_2018()
g.msgbox('End ')

