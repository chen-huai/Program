#!/usr/biEn/env python
# -*- coding:utf-8 -*-

import easygui as g
import os
import os.path
from win32com.client import Dispatch
import time
import csv
import re
file = __file__
file_2= file.split('\\')
file_2.pop()
file_address = '\\'.join(file_2)
t_1 = int(time.strftime('%Y'))
t_2 = t_1-1
now = str(t_1)
last_time = str(t_2)
name='ECO Batch '+time.strftime("%Y%m%d")
def eco_batch():
    path=g.fileopenbox(msg=None,title=None,default='Z:/Inorganic_batch/Microwave/batch/*.doc',filetypes=None)
    w=Dispatch('Word.Application')
    w.Visible=0
    doc=w.Documents.Open(r'%s'%path)
    a=doc.Content.Text
    b=a.split('\r')
    sequence_mm = []
    rw=int(input('input number:'))
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
        if ('9E' in i) or ('8E' in i) or ('Extractable' in i) or ('HeavyMetals' in i) or ('Al.Cr.Ti,Zr' in i) or ('ISO17072-1' in i)  or ('SolubleHM' in i):
                analyte_list.append(i)     
    for each in b:
        if ('/' in each) and (len(each)>5) and ('/'+last_time not in each) and ('/'+now not in each)and ('GB/T' not in each)and ('D' not in each):
                sample_list.append(each) 
    top=['No.','Sample No.','Analyte','Weight','Volume','DF','Batch No.']
    blk=['1','BLK','','','','5','']    
    output_file = open('Z:/Inorganic_batch/Microwave/result/%s.csv'%name,'at',newline='',encoding='utf-8')
    output_writer = csv.writer(output_file)
    r=rw
    if r==1:
        output_writer.writerow(top)
        output_writer.writerow(blk)
        r+=1
    n=0
    for each in sample_list:
        batch=['5']
        batch.insert(0,r)
        batch.insert(1,each)
        batch.insert(2,analyte_list[n])
        batch.insert(3,weight[n])
        batch.insert(4,volume[n])
        batch.insert(6,batch_no[0])
        output_writer.writerow(batch)
        r+=1 
        n+=1
def limit():
    kon=[]
    top_2=['','Analyte','RL','DL','UV','Unit','Unit (Raw Data)']
    sb=['','Sb',0.5,2,'10%','mg/kg','ug/L']
    as_1=['','As',0.2,2,'10%','mg/kg','ug/L']
    cd=['','Cd',0.1,0.2,'10%','mg/kg','ug/L']
    hg=['','Hg',0.02,0.2,'10%','mg/kg','ug/L']
    output_file = open('Z:/Inorganic_batch/Microwave/result/%s.csv'%name,'at',newline='',encoding='utf-8')
    output_writer = csv.writer(output_file)
    output_writer.writerow(kon)
    output_writer.writerow(top_2)
    output_writer.writerow(sb)
    output_writer.writerow(as_1)
    output_writer.writerow(cd)
    sb[1]='Cr'
    output_writer.writerow(sb)
    sb[1]='Co'
    output_writer.writerow(sb)
    sb[1]='Cu'
    output_writer.writerow(sb)
    as_1[1]='Pb'
    output_writer.writerow(as_1)
    output_writer.writerow(hg)
    sb[1]='Ni'
    output_writer.writerow(sb)
    sb[1]='Ba'
    output_writer.writerow(sb)
    sb[1]='Se'
    output_writer.writerow(sb)
    sb[1]='Mn'
    output_writer.writerow(sb)
    sb[1]='Zn'
    output_writer.writerow(sb)
    sb[1]='Al'
    output_writer.writerow(sb)
    sb[1]='Ti'
    output_writer.writerow(sb)
    sb[1]='Zr'
    output_writer.writerow(sb)    
eco_batch() 
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while  action==1:
    eco_batch()
    action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))       
else:
    limit()
    os.startfile('Z:/Inorganic_batch/Microwave/result')                
            