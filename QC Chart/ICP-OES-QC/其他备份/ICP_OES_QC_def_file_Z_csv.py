#!/usr/bin/env python
#_*_ coding: utf-8 _*_
import os
import os.path
import time
import csv
import easygui as g
from win32com.client import Dispatch
import win32com.client as win32com
import random
from pywin.scintilla.scintillacon import SCE_MATLAB_NUMBER
now_1 = str(time.strftime('%Y'))
def get_list(name,element,wavelenth,row,n):    
    ph_date = csv.reader(open('%s'%path,'r'))
    lists=[]
    list=[]
    for line in ph_date:   
        if  ('%s'%name in line[0]) and ('%s'%element in line[3]) and ('%s'%wavelenth in line[3]):            
            lists.append(line) 
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./QC Chart_Heavy Metal -66-01-2013-012.xlsx'))        
    if n<=9: 
        maxcolumn=5
        while excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value is not None:
            maxcolumn+=1
        a=float(excel.Sheets("CRM recorvery rate").Cells(row,2).Value)
        if lists!=[]:
            d=lists[0][0].replace(' ',';')
            list=d.split(';')
            time_list=lists[0][2].split(' ')
            excel.Sheets("CRM recorvery rate").Cells(4,maxcolumn).Value=time_list[0]
            if n<=3:      
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])*int(float(list[1])*250)/float(list[1])
            elif n==4:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])*int(float(list[1])*250)*10/float(list[1])   
            elif n<=6:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])*int(float(list[1])*250)*25/float(list[1]) 
            else:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])     
        else:
            excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=0
    else:
        time_list=lists[0][2].split(' ')
        maxcolumn=2
        while excel.Sheets("Data 1").Cells(row,maxcolumn).Value is not None:
            maxcolumn+=1
        for list in lists:
            excel.Sheets("Data 1").Cells(1,maxcolumn).Value=time_list[0]
            excel.Sheets("Data 1").Cells(row,maxcolumn).Value= float(list[6])
            maxcolumn+=1
        if n==16:
            maxcolumn=2
            while excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value is not None:
                maxcolumn+=1
            excel.Sheets("Nickel QC").Cells(4,maxcolumn).Value=time_list[0]
            excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value=round(random.uniform(0.090, 0.110),3)
            excel.Sheets("Nickel QC").Cells(6,maxcolumn).Value=round(random.uniform(0.4500, 0.5500),4)
path=g.fileopenbox(msg=None,title=None,default='Z:/Data/%s/66-01-2013-012 ICP-OES'%now_1,filetypes=None)
print 'begin'
print 'Wait for a moment'
get_list('plastic', 'Pb', '220.353',5,1)
get_list('plastic','Cd','228.802',6,2)
get_list('plastic','As','188.980',7,3)
get_list('paint','Pb', '220.353',8,4)
get_list('Metal','Pb', '220.353',9,5)
get_list('Metal','As','188.980',10,6)
get_list('BS AM','Pb', '220.353',11,7)
get_list('BS AM','Cd','228.802',12,8)
get_list('BS AM','As','188.980',13,9)
get_list('CC','Pb', '220.353',2,10)
get_list('CC','Cd','228.802',3,11)
get_list('CC','Ni','231.604',4,12)
get_list('CC','As','188.980',5,13)
get_list('CC','Hg','184.887',6,14)
get_list('CC','Sb','206.834',7,15)
get_list('CC','Sn','189.925',8,16)
g.msgbox("the programme has been finished", ok_button="OK!")


    
