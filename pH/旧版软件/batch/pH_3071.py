#_*_ coding: utf-8 _*_
#revise date:0912 add Sequency No.

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
file = __file__
file_2= file.split('\\')
file_2.pop()
file_address = '\\'.join(file_2)

now = int(time.strftime('%Y'))
last_time = now-1

now_list = list(str(now))
del now_list[0]
del now_list[0]
now_2 = ''.join(now_list)

name=time.strftime("%Y_%m_%d_%H_%M")


#鑾峰彇sample no
path=g.fileopenbox(msg=None,title=None,default=file_address+'/*.doc',filetypes=None)
w=Dispatch('Word.Application')
w.Visible=0
doc=w.Documents.Open(r'%s'%path)
a=doc.Content.Text
b=a.split('\r')
sample_list=[]
        
for each in b:               
    if ('/' in each) and (len(each)>5) and ('/'+str(last_time) not in each) and ('/'+str(now) not in each) and ('D' not in each) and ('GB' not in each) and ('(1)' not in each) :
        sample_list.append('%s'%each)

print(sample_list)


#濉埌excel涓�
excel = win32com.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.Application.DisplayAlerts = False
workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'c:\ph\ph.xlsx'))

length=len(sample_list)


Project_No='66.'+now_2+'.'
sequency_No=1

if length<8:
    row_number=22
    for each in sample_list:
        excel.Sheets('1').Cells(row_number,1).Value=each
        excel.Sheets('1').Cells(row_number,2).Value=sequency_No
        excel.Sheets('1').Cells(row_number+1,1).Value='QC'
        row_number+=1
        sequency_No+=1
        if each.split('/')[0] not in Project_No:
            Project_No=Project_No+each.split('/')[0]+'/'
            excel.Sheets('1').Cells(4,3).Value=Project_No


    
if (length>8) or (length==8):
    
    row_number=22
    for each in sample_list[0:7]:
        excel.Sheets('1').Cells(row_number,1).Value=each
        excel.Sheets('1').Cells(row_number,2).Value=sequency_No
        row_number+=1
        sequency_No+=1
        if each.split('/')[0] not in Project_No:
            Project_No=Project_No+each.split('/')[0]+'/'
            excel.Sheets('1').Cells(4,3).Value=Project_No

    excel.Sheets('1').Cells(row_number,1).Value='QC'
    row_number=20
    count=1
    sheetname=2

    Project_No='66.'+now_2+'.'

    for each in sample_list[7:]:

        if count%9!=0:
            excel.Sheets(sheetname).Cells(row_number,1).Value=each
            excel.Sheets(sheetname).Cells(row_number,2).Value=sequency_No
            excel.Sheets(sheetname).Cells(row_number+1,1).Value='QC'
            row_number+=1
            sequency_No+=1
            if each.split('/')[0] not in Project_No:
                Project_No=Project_No+each.split('/')[0]+'/'
                excel.Sheets(sheetname).Cells(4,3).Value=Project_No
       
            
        if count%9==0:
            excel.Sheets(sheetname).Cells(row_number,1).Value=each
            excel.Sheets(sheetname).Cells(row_number,2).Value=sequency_No
            excel.Sheets(sheetname).Cells(row_number+1,1).Value='QC'
            if each.split('/')[0] not in Project_No:
                Project_No=Project_No+each.split('/')[0]+'/'
                excel.Sheets(sheetname).Cells(4,3).Value=Project_No
            sequency_No+=1

            
            excel.Sheets(sheetname).Copy(None, excel.Sheets(sheetname))
            old_name=sheetname
            sheetname+=1
            excel.Sheets('%s (2)'%old_name).Name=sheetname
            row_number=20
            excel.Sheets(sheetname).Cells(20,1).Value=''
            excel.Sheets(sheetname).Cells(21,1).Value=''
            excel.Sheets(sheetname).Cells(22,1).Value=''
            excel.Sheets(sheetname).Cells(23,1).Value=''
            excel.Sheets(sheetname).Cells(24,1).Value=''
            excel.Sheets(sheetname).Cells(25,1).Value=''
            excel.Sheets(sheetname).Cells(26,1).Value=''
            excel.Sheets(sheetname).Cells(27,1).Value=''
            excel.Sheets(sheetname).Cells(28,1).Value=''
            excel.Sheets(sheetname).Cells(29,1).Value=''
            excel.Sheets(sheetname).Cells(20,2).Value=''
            excel.Sheets(sheetname).Cells(21,2).Value=''
            excel.Sheets(sheetname).Cells(22,2).Value=''
            excel.Sheets(sheetname).Cells(23,2).Value=''
            excel.Sheets(sheetname).Cells(24,2).Value=''
            excel.Sheets(sheetname).Cells(25,2).Value=''
            excel.Sheets(sheetname).Cells(26,2).Value=''
            excel.Sheets(sheetname).Cells(27,2).Value=''
            excel.Sheets(sheetname).Cells(28,2).Value=''
            excel.Sheets(sheetname).Cells(29,2).Value=''
            Project_No='66.'+now_2+'.'
       
        count+=1

if excel.Sheets(sheetname).Cells(20,1).Value=='':
    excel.Sheets(sheetname).Delete()

    
workbook.SaveAs('c:\ph\%s.xlsx'%name)   
excel.Application.Quit()


g.msgbox("the programme has been finished", ok_button="OK!")

os.startfile('c:/ph/')

excel = win32com.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'c:\ph\%s.xlsx'%name))

