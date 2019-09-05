#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import os.path
import time
import easygui as g
from win32com.client import Dispatch
import win32com.client as win32com
import random
import csv
import tkinter
now_1 = str(time.strftime('%Y'))
def get_list():  
    lab_number = input('please fill in the lab number:')
    quality=float(input('please fill in the sample quality:'))
    volume=float(input('please fill in the volume of the sample:'))
    print ('begin')
    print ('Wait for a moment')
    name = lab_number.replace('/', '_') 
    elements=['Zn','B','As','Sb','Cr','Pb','Co','Si','Ba','Sn','Mo','Zr','Al','Ti','Cd','Na','K','Sr','Ca','Mg']
    #wavelenths=['202.548','249.772','188.980','206.834','267.716','220.353','228.615','230.424','189.925','202.032','343.823','236.705','228.802','589.592','460.733']
    rows=['3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22']
    n=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
    lines= csv.reader(open('%s'%path,"rt",encoding='utf-8'))
    lists=[]
    for l in lines:
        lists.append(l) 
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./SVHC.xlsx'))
    for i in range(len(elements)):
        list=[]
        m=0
        for line in lists:
            if ('%s'%lab_number in line[0]) and ('%s'%elements[i] in line[2]) and ('(ref)' not in line[2]):                         
                list.append(line) 
            m+=1
        excel.Sheets("1").Cells(rows[i],2).Value=elements[i]
        if list!=[]:
            if str(list[0][4])=='未校正':
                excel.Sheets("1").Cells(rows[i],3).Value='未校正'
            else:
                excel.Sheets("1").Cells(rows[i],3).Value=float(list[0][4])*volume/quality
        else:
            excel.Sheets("1").Cells(rows[i],3).Value='未走标准曲线'
        if n[i]==n[len(elements)-1]:  
            address=os.path.abspath('.')
            workbook.SaveAs('%s\\SVHC %s.xlsx'%(address,name))
            excel.Application.Quit()
path=tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/66-01-2018-012 5110 ICP-OES/'%now_1,filetypes=[("csvfile", "*.csv")])
get_list()
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while  action==1:
    get_list()
    action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))       
else:    
    os.startfile('.\\')

