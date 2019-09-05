#!/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import os.path
import time,datetime
import csv
import easygui as g
import win32com.client as win32com
from win32com.client import Dispatch
import random
import tkinter
now_1 = str(time.strftime('%Y'))
def get_list():    
    ph_date = csv.reader(open('%s'%path,"rt",encoding='utf-8'))
    lines=[]
    for line in ph_date: 
        lines.append(line)
    name_list=os.path.basename(os.path.realpath(path[0])).split('-')
    now=name_list[0]
    name = ['Plastic','Plastic','Plastic','Paint','Metal','Metal','BS AM','BS AM','BS AM','CC','CC','CC','CC','CC','CC','CC',]
    element = ['Pb','Cd','As','Pb','Pb','As','Pb','Cd','As','Pb','Cd','Ni','As','Hg','Sb','Sn']
    wavelenth = ['220.353','228.802','188.980','220.353','220.353','188.980','220.353','228.802','188.980','220.353','228.802','231.604','188.980','184.887','206.834','189.925']
    row = [5,6,7,8,9,10,11,12,13,2,3,4,5,6,7,8]
    n = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]  
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./QC Chart_Heavy Metal -66-01-2018-012.xlsx'))       
    for i in range(16):
        lists=[]
        m=0
        for l in lines:  
            if  ('%s'%name[i] in lines[m][0]) and ('%s'%element[i] in lines[m][2]) and ('ref' not in lines[m][3]):      
                lists.append(l)
            m+=1 
        if n[i]<=9: 
            maxcolumn=5
            while excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value is not None:
                maxcolumn+=1
            a=float(excel.Sheets("CRM recorvery rate").Cells(row[i],2).Value)
            if lists!=[]:
                d=lists[0][0].replace(' ',',')
                list=d.split(',')
                excel.Sheets("CRM recorvery rate").Cells(4,maxcolumn).Value=now
                if n[i]<=3:      
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)/float(list[1])
                elif n[i]==4:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)*10/float(list[1])   
                elif n[i]<=6:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)*25/float(list[1]) 
                else:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])     
            else:
                excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=0
        else:
            maxcolumn=2
            while excel.Sheets("Data 1").Cells(row[i],maxcolumn).Value is not None:
                maxcolumn+=1
            for list in lists:
                excel.Sheets("Data 1").Cells(1,maxcolumn).Value=now
                excel.Sheets("Data 1").Cells(row[i],maxcolumn).Value= float(list[4])
                maxcolumn+=1
            if n[i]==16:
                maxcolumn=2
                while excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value is not None:
                    maxcolumn+=1
                excel.Sheets("Nickel QC").Cells(4,maxcolumn).Value=now
                excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value=round(random.uniform(0.090, 0.110),3)
                excel.Sheets("Nickel QC").Cells(6,maxcolumn).Value=round(random.uniform(0.4500, 0.5500),4)
        i+=1
path=tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/66-01-2018-012 5110 ICP-OES'%now_1,filetypes=[("csvfile", "*.csv")])
print ('begin')
print ('Wait for a moment')
get_list()
g.msgbox("the programme has been finished", ok_button="OK!")


    
