#!/usr/bin/env python
# -*- coding: cp936 -*-
import easygui as g
import datetime
import os
import os.path
from win32com.client import Dispatch
import time
import csv
import re
t_1 = int(time.strftime('%Y'))
t_2 = t_1-1
now = str(t_1)
last_time = str(t_2)
file = __file__
file2 = file.split('\\')
file2.pop()
file_address = '\\'.join(file2)
def pH_batch():     
        path=g.fileopenbox(msg=None,title=None,default=file_address+'/*.xlsx',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence_mm = []
        name = 'ph batch'+time.strftime(" %Y-%m-%d")
        rw=int(raw_input('input number:'))
        for each in b:
            if ('/' in each) and (len(each)>5) and ('/'+last_time not in each) and ('/'+now not in each)and ('GB/T' not in each)and ('D' not in each):
                    sequence_mm.append(each)
                    
        standard_1 = ['pH cal','1','Standard','','','','','','','','1']
        cc = ['pH Measure','1','CC','','','','','','','','1']
        blk_1 = ['pH Measure','1','BLK','before','','','','','','','1']
        blk_2 = ['pH Measure','1','BLK','after','','','','','','','1']              
        output_file = open('Z:/Inorganic_batch/Formaldehyde/batch/%s.csv'%name,'ab')        
        output_writer = csv.writer(output_file)  
        m=rw 
        ph_4045=re.findall('ISO4045',a)
        if m==1: 
            if ph_4045!=[]:
                blk_2[2]='4045 BLK'
                output_writer.writerow(blk_2)  
            else:                          
                output_writer.writerow(standard_1)
                output_writer.writerow(cc)
                output_writer.writerow(blk_1)
                output_writer.writerow(blk_2)
                m+=1            
        r=rw
        qc=0                    
        for each in sequence_mm:
            if ph_4045!=[]:
                batch = ['pH Measure','1','','','','','','','1']
                batch.insert(2,'4045 '+str(r))
                batch.insert(3,each+'A') 
                output_writer.writerow(batch)
                batch[3]=each+'B'  
                output_writer.writerow(batch)
                batch[3]=each+'C'
                output_writer.writerow(batch)
                batch[3]=each+'D'
                output_writer.writerow(batch)
                r+=1   
            else:                 
                batch = ['pH Measure','1','','','','','','','1']
                batch.insert(2,r)
                batch.insert(3,each)
                output_writer.writerow(batch)
                batch[3]=each+'A'  
                output_writer.writerow(batch)
                batch[3]=each+'B'
                output_writer.writerow(batch) 
                if (qc+1)%20==0:
                    output_writer.writerow(cc)
                qc+=1            
                r+=1
        output_writer.writerow(cc)  
name = 'ph batch'+time.strftime(" %Y-%m-%d")
def di_water():
    output_file = open('Z:/Inorganic_batch/Formaldehyde/batch/%s.csv'%name,'ab')        
    output_writer = csv.writer(output_file) 
    water = ['pH Measure','1','DI Water','','','','','','','','1'] 
    output_writer.writerow(water)           
pH_batch()
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while  action==1:
    pH_batch()
    action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))       
else:
    di_water()
    os.startfile('Z:/Inorganic_batch/Formaldehyde/batch') 