# -*- coding:utf-8 -*-


import easygui as g
import time
import datetime
import os
import os.path
from win32com.client import Dispatch
import win32com.client as win32com
now = int(time.strftime('%Y'))
last_time = now-1
def Cr():
        path=g.fileopenbox(msg=None,title=None,default='Z:/Inorganic_batch/Formaldehyde/batch/*.doc',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence=[]
        count=0
        doc.Close()
        w.Quit()       
        for each in b:               
            if ('/' in each) and (len(each)>5) and ('/'+str(last_time) not in each) and ('/'+str(now) not in each) :
                    if 'agingB' in b[count+5]:
                        sequence.append('%sB'%each)
                    elif 'aging' in b[count+5]:   
                        sequence.append('%sA'%each)
                    else:
                            sequence.append('%s'%each)                      
            count+=1
        name=time.strftime("CR %m-%d-%H-%M-%S")       
        file=open('Z:/Inorganic_batch/Formaldehyde/batch/%s.txt'%name,'a')
        file.write('BLK\n')
        file.write('BLK+DPC\n')
        file.write('BLK SPIKE\n')
        file.write('BLK SPIKE+DPC\n')
        file.write('SAMPLE SPIKE\n')
        file.write('SAMPLE SPIKE+DPC\n')
        QC=6
        for each in sequence:
            QC+=1
            file.write('%s\n'%each)
            file.write('%s+D\n'%each)
            if QC%20==0:
                file.write('QC\n')                       
        file.write('QC\n')
def HCHO():
        path=g.fileopenbox(msg=None,title=None,default='Z:/Inorganic_batch/Formaldehyde/batch/*.doc',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence=[]
        doc.Close()
        w.Quit()
        #判断是否日本方法
        JISL=0       
        for each in b:              
                if ('/' in each) and (len(each)>5) and ('/'+str(last_time) not in each) and ('/'+str(now) not in each) and ('Formal+GB/T' not in each) and ('5mg/kg' not in each):
                        sequence.append('%s'%each)                        
                if ('Formal+JISL1041A') in each:
                        JISL=1
                if ('Formal+JISL1041B') in each:
                        JISL=2      
        name=time.strftime("HCHO %m-%d-%H-%M-%S")       
        file=open('Z:/Inorganic_batch/Formaldehyde/batch/%s.txt'%name,'a')       
        file.write('BLK\n')
        file.write('BLK SPIKE\n')
        file.write('SAMPLE SPIKE\n')
        if JISL==0:
            QC=3
            for each in sequence:
                QC+=1
                file.write('%s\n'%each)
                if QC%20==0:
                        file.write('QC\n')

            file.write('QC\n')
        if JISL==1:
            for each in sequence:
                    file.write('A0-%s\n'%each)
            for each in sequence:
                    file.write('A1-%s\n'%each)
            file.write('AS-4.5\n')
        if JISL==2:
            for each in sequence:
                    file.write('A0-%s\n'%each)
            for each in sequence:
                    file.write('A2-%s\n'%each)

            file.write('AS-4.5\n') 
option=g.ccbox('which test item', choices=('Cr','HCHO'))
if option==1:
        Cr()
else:
        HCHO()
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while action ==1:
        option=g.ccbox('which test item', choices=('Cr','HCHO'))
        if option==1:
                Cr()
        else:
                HCHO()
        action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))        
os.startfile('Z:/Inorganic_batch/Formaldehyde/batch')
