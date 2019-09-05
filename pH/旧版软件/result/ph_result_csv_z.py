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
t_3 = time.strftime('%Y-%m-%d %H:%M:%S ')
file = __file__
file2 = file.split('\\')
file2.pop()
file_address = '\\'.join(file2)
def pH_result():
        path=g.fileopenbox(msg=None,title=None,default=file_address+'/*.xlsx',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence_mm = []
        name = 'ph result'+time.strftime(" %Y-%m-%d")   
        for each in b:
            if ('/' in each) and (len(each)>5) and ('/'+last_time not in each) and ('/'+now not in each)and ('GB/T' not in each)and ('D' not in each):
                    sequence_mm.append(each)
            top= ['Determination start','Method name','ID1.Value','ID2.Value','RS01.Name','RS01.Value','RS02.Name','RS02.Value','Lab TEMP','[DELTA]ph']
            kon = ['','','','','','','','','','']   
            standard_1 = ['%s UTC+8'%t_3,'pH Cal','Standard','','','','','','','0']
            cc = ['%s UTC+8'%t_3,'pH Measure','CC','','pH','','T','']
            blk_1 = ['%s UTC+8'%t_3,'pH Measure','BLK','before','pH','','T','']
            blk_2 = ['%s UTC+8'%t_3,'pH Measure','BLK','after','pH','','T','']              
            output_file = open('Z:/Inorganic_batch/Formaldehyde/result/%s.csv'%name,'wb')
            output_writer = csv.writer(output_file) 
            output_writer.writerow(top)
            output_writer.writerow(kon)                                    
            output_writer.writerow(standard_1)
            output_writer.writerow(cc)
            output_writer.writerow(blk_1)
            output_writer.writerow(blk_2)  
            r=1
            qc=0                    
            for each in sequence_mm:   
                batch = ['%s UTC+8'%t_3,'pH Measure','pH','','T','']
                batch.insert(2,r)
                batch.insert(3,each)
                output_writer.writerow(batch)
                batch[3]=each+'A'  
                output_writer.writerow(batch)
                batch[3]=each+'B'
                output_writer.writerow(batch) 
                if (qc+1)%10==0:
                    output_writer.writerow(cc)
                qc+=1
                r+=1
        output_writer.writerow(cc)   
        os.startfile('Z:/Inorganic_batch/Formaldehyde/result')                        
                
                                                                
pH_result()


        
        
        
