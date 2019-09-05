#!/usr/bin/env python
# -*- coding: cp936 -*-
import easygui as g
import openpyxl
import datetime
import os
import os.path
import codecs,sys
from win32com.client import Dispatch
import win32com.client as win32com
import tkinter
import time
t_1 = int(time.strftime('%Y'))
t_2 = t_1-1
now = str(t_1)
last_time = str(t_2)
def heavy_metal():
        path=g.fileopenbox(msg=None,title=None,default='./*.xlsx',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence=[]
        sequence_mm=[]
        for each in b:
                if ('/' in each) and (len(each)>5) and ('/'+last_time not in each) and ('/'+now not in each) and ('D' not in each):                        
                            sequence_mm.append(each)
        excel = win32com.gencache.EnsureDispatch('Excel.Application')        
        excel.Visible = 1
        wb=excel.Workbooks.Add()
        ws=wb.Worksheets('Sheet1')
        row=2
        for each in sequence_mm:
                ws.Cells(row,1).Value=each+'A'
                row+=1                        
                ws.Cells(row,1).Value=each+'B'
                row+=1       
                ws.Cells(row,1).Value=each+'C'
                row+=1       
heavy_metal()
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while action ==1:
        heavy_metal()
        action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
        
        
        
