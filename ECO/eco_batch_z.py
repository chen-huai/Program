#!/usr/biEn/env python
# -*- coding:utf-8 -*-
import easygui as g
import os
import os.path
from win32com.client import Dispatch
import win32com.client as win32com
import time
import re
t_1 = int(time.strftime('%Y'))
t_2 = t_1-1
now = str(t_1)
last_time = str(t_2)
name='ECO Batch '+time.strftime("%Y%m%d %H_%M")
def eco_batch():
    path=g.fileopenbox(msg=None,title=None,default='Z:/Inorganic_batch/Microwave/batch/*.doc',filetypes=None)
    w=Dispatch('Word.Application')
    w.Visible=0
    doc=w.Documents.Open(r'%s'%path)
    a=doc.Content.Text
    b=a.split('\r')
    sample_list=[]
    analyte_list=[]
    weight=[]
    volume=[]
    batch_no=re.findall('Inorganic.\d{6}.\d{1,3}',a)
    for r in b:
        if re.findall('\d+\.+\d{1,4}',r)!=[]:
            if ('CHM' not in r):
                weight.append(r)
                volume.append(int(float('%.2f'%float(r))*50))       
    for i in b:
        if ('9E' in i) or ('8E' in i) or ('Extractable' in i) or ('HeavyMetals' in i) or ('Al.Cr.Ti,Zr' in i) or ('ISO17072-1' in i) or ('SolubleHM' in i):
                analyte_list.append(i)      
    for each in b:
        if ('/' in each) and (len(each)>5) and ('/'+last_time not in each) and ('/'+now not in each)  and ('D' not in each) and ('GB/T' not in each):
                sample_list.append(each) 
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = 1  
    wb=excel.Workbooks.Add()
    ws=wb.Worksheets('Sheet1')
    ws.Cells(1,1).Value='No.'
    ws.Cells(1,2).Value='Sample No.'
    ws.Cells(1,3).Value='Analyte'
    ws.Cells(1,4).Value='Weight'            
    ws.Cells(1,5).Value='Volume'
    ws.Cells(1,6).Value='DF'
    ws.Cells(1,7).Value='Batch No'
    ws.Cells(2,1).Value=1
    ws.Cells(2,2).Value='BLK'
    ws.Cells(2,6).Value=5
    r=3
    n=0
    for each in sample_list:
        ws.Cells(r,1).Value=r-1
        ws.Cells(r,2).Value=each
        ws.Cells(r,3).Value=analyte_list[n]
        ws.Cells(r,4).Value=weight[n]
        ws.Cells(r,5).Value=volume[n]
        ws.Cells(r,6).Value=5
        ws.Cells(r,7).Value=batch_no[0]
        r+=1
        n+=1
    r+=1
    ws.Cells(r,2).Value='Analyte'
    ws.Cells(r,3).Value='RL'
    ws.Cells(r,4).Value='DL'
    ws.Cells(r,5).Value='UV'
    ws.Cells(r,6).Value='Unit'
    ws.Cells(r,7).Value='Unit (Raw Data)'
    r+=1
    ws.Cells(r,2).Value='Sb'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1
    ws.Cells(r,2).Value='As'
    ws.Cells(r,3).Value=0.2
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1
    ws.Cells(r,2).Value='Cd'
    ws.Cells(r,3).Value=0.1
    ws.Cells(r,4).Value=0.2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1   
    ws.Cells(r,2).Value='Cr'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Co'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Cu'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Pb'
    ws.Cells(r,3).Value=0.2
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Hg'
    ws.Cells(r,3).Value=0.02
    ws.Cells(r,4).Value=0.2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Ni'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Ba'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Se'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Mn'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1    
    ws.Cells(r,2).Value='Zn'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1 
    ws.Cells(r,2).Value='Al'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1   
    ws.Cells(r,2).Value='Ti'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1   
    ws.Cells(r,2).Value='Zr'
    ws.Cells(r,3).Value=0.5
    ws.Cells(r,4).Value=2
    ws.Cells(r,5).Value='10%'
    ws.Cells(r,6).Value='mg/kg'
    ws.Cells(r,7).Value='ug/L'
    r+=1  
    wb.SaveAs('Z:/Inorganic_batch/Microwave/result/%s.xlsx'%name)         
eco_batch()   
g.msgbox("the programme has been finished", ok_button="OK!")             
            