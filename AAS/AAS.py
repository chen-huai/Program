#_*_ coding: utf-8 _*_
import easygui as g

import os
from win32com.client import Dispatch
import win32com.client as win32com

import time
file = __file__
file_2= file.split('\\')
file_2.pop()
file_address = '\\'.join(file_2)

now = int(time.strftime('%Y'))
last_time = now-1


def heavy_metal():
        path=g.fileopenbox(msg=None,title=None,default=file_address+'/*.xlsx',filetypes=None)
        w=Dispatch('Word.Application')
        w.Visible=0
        doc=w.Documents.Open(r'%s'%path)
        a=doc.Content.Text
        b=a.split('\r')
        sequence=[]
        sequence_mm=[]
        for each in b:
                #print(each)
                if ('/' in each) and (len(each)>5) and ('/'+str(last_time) not in each) and ('/'+str(now) not in each):
                        
                        if ' ' in each :
                                
                                if each.split(' ')[1] not in sequence:
                                        #鏃犳硶璇嗗埆CPSC-AM, 闇�瑕佹墜鍔ㄨ浆鎹�
                                        if 'CPSC\x1eAM' in each:
                                                
                                                sequence.append(each.split(' ')[1])
                                                sequence_mm.append('CPSC-AM %s'%each.split(' ')[1])

                                        else:
                                                sequence.append(each.split(' ')[1])
                                                sequence_mm.append(each)
                                else:
                                        # 123/1 鍜孧M 123/1锛屽氨淇濈暀MM 123/1,
                                        # MM 123/1 鍜� CPSCMM 123/1 淇濈暀绗竴涓�
                                        if each.split(' ')[1] in sequence_mm:
                                                sequence_mm.remove(each.split(' ')[1])
                                                sequence_mm.append(each)

                                                
                        else:
                                if each not in sequence:
                                        sequence.append(each)
                                        sequence_mm.append(each)

        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        
        excel.Visible = 1
        wb=excel.Workbooks.Add()
        ws=wb.Worksheets('Sheet1')
        
        row=2
        for each in sequence_mm:
                
                ws.Cells(row,1).Value=each.replace('+','-')
                row+=1                        
        
heavy_metal()
action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while action ==1:
        heavy_metal()
        action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
        
