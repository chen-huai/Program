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
def HCHO_QC_Chart():    
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/QC chart/%s'%now_1, title="Select files",filetypes=[("csvfile", "*.csv")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'Z:/QC chart/%s/QC Chart_HCHO_CARY60.xlsx'%now_1))        
        maxcolumn=8
        while excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn).Value is not None:
            maxcolumn+=1  
        data = csv.reader(open('%s'%path,'r'))
        k=0
        for each in data:
            if re.findall('\d{1,2}/\d{1,2}/\d{4}',each[1]) !=[]:
                now=re.findall('\d{1,2}/\d{1,2}/\d{4}',each[1])
            if ('QC' in each[0]):
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).Value=now[0]
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).NumberFormat= "yyyy/m/d"
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(3,maxcolumn+k).Value='QC'
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).Value=each[1]
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).NumberFormat="0.0000"            
                k+=1
def Cr_QC_Chart():    
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/QC chart/%s'%now_1, title="Select files",filetypes=[("csvfile", "*.csv")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'Z:/QC chart/%s/QC Chart_Cr_CARY100.xlsx'%now_1))        
        maxcolumn=8
        while excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn).Value is not None:
            maxcolumn+=1  
        data = csv.reader(open('%s'%path,'r'))
        k=0
        for each in data:
            if re.findall('\d{1,2}/\d{1,2}/\d{4}',each[1]) !=[]:
                now=re.findall('\d{1,2}/\d{1,2}/\d{4}',each[1])
            if ('QC' in each[0]):
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).Value=now[0]
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).NumberFormat= "yyyy/m/d"
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(3,maxcolumn+k).Value='QC'
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).Value=each[1]
                excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).NumberFormat="0.0000"            
                k+=1               
option=g.ccbox('which test item',title='UV QC data input', choices=('Cr','HCHO'))
if option==1:
        print 'Cr'
        Cr_QC_Chart()
else:
        print 'Formal'
        HCHO_QC_Chart()
g.msgbox('End ')

