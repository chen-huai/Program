#!/usr/bin/env python
# -*- coding: cp936 -*-
import os
import os.path
import time
import easygui as g
from win32com.client import Dispatch
import win32com.client as win32com
now_1 = str(time.strftime('%Y'))
def get_list(name,element,wavelenth,row,n):
    f2 = open(path,"r")
    lines = f2.readlines()
    lists=[]
    for line in lines:
        if ('%s'%name in line) and ('%s'%element in line)and ('%s'%wavelenth in line):
            line=line.replace(',',' ')
            lists.append(line.split())     
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'Z:/QC chart/%s/QC Chart_Heavy Metal -66-01-2013-012.xlsx'%now_1))        
    print lists 
    if n<=9: 
        maxcolumn=5
        while excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value is not None:
            maxcolumn+=1
        a=float(excel.Sheets("CRM recorvery rate").Cells(row,2).Value)
        if lists!=[]:
            excel.Sheets("CRM recorvery rate").Cells(4,maxcolumn).Value=lists[0][10]
            if n<=3:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][5])*50/float(lists[0][1])
            elif n==4:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])*50*10/float(lists[0][1])
            elif n<=6:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][6])*50*25/float(lists[0][1]) 
            else:
                excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=float(lists[0][5])     
        else:
            excel.Sheets("CRM recorvery rate").Cells(row,maxcolumn).Value=0
    else:
        maxcolumn=2
        while excel.Sheets("Data 1").Cells(row,maxcolumn).Value is not None:
            maxcolumn+=1
        for list in lists:
            excel.Sheets("Data 1").Cells(1,maxcolumn).Value=list[11]
            excel.Sheets("Data 1").Cells(row,maxcolumn).Value= float(list[6])
            maxcolumn+=1     
path=g.fileopenbox(msg=None,title=None,default='C:\Users\chenhuai\Desktop\DATE\2018\*.doc',filetypes=None)
get_list('plastic', 'Pb', '220.353',5,1)
get_list('plastic','Cd','228.802',6,2)
get_list('plastic','As','188.980',7,3)
get_list('paint','Pb', '220.353',8,4)
get_list('Metal','Pb', '220.353',9,5)
get_list('Metal','As','188.980',10,6)
get_list('BS AM','Pb', '220.353',11,7)
get_list('BS AM','Cd','228.802',12,8)
get_list('BS AM','As','188.980',13,9)
get_list('Cont Calib Verif','Pb', '220.353',2,10)
get_list('Cont Calib Verif','Cd','228.802',3,11)
get_list('Cont Calib Verif','Ni','231.604 ',4,12)
get_list('Cont Calib Verif','As','188.980',5,13)
get_list('Cont Calib Verif','Hg','184.887',6,14)
get_list('Cont Calib Verif','Sb','206.834',7,15)
get_list('Cont Calib Verif','Sn','189.925',8,16)
g.msgbox("the programme has been finished", ok_button="OK!")


    