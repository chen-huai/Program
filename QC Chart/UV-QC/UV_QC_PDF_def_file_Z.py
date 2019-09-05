#_*_ coding: utf-8 _*_

import easygui as g
import datetime
import os
import os.path
import codecs,sys
import win32com.client as win32com
import tkinter
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import win32timezone
import time
import re
now = time.strftime('%Y')
t = str(now)
print("Begin")
def UV_QC(name,matter,value,address):
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/%s/%s'%(t,address,name), title="Select files",filetypes=[("pdf file", "*.pdf")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./%s'%matter))        
        maxcolumn=8
        while excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(3,maxcolumn).Value is not None:
                maxcolumn+=1        
        k=0
        now_2=[]
        for each1 in path:
                p = r'%s'%each1
                pdf = PdfFileReader(open(p, 'rb'))                
                total_page_number=pdf.getNumPages()
                page_number=0                
                while  page_number <total_page_number:
                        pageobj=pdf.getPage(page_number)
                        text=pageobj.extractText()
                        a=text.split()
                        now_2.append(re.findall("(\d{1,2}/\d{1,2}/\d{4})",text))
                        lenth=len(a)
                        count=0
                        b='\n'.join(a)
                        if re.findall("(QC\d.\d{3,4})",b) !=[]:
                            for each in a:
                                if (('QC' in each) or ('qc' in each)) and ('*' not in each) and ('/' not in each)and ('QCQ' not in each):
                                    content1=each.split(' ') 
                                    content2=[]
                                    for each in content1:
                                        if each != '':
                                            content2.append(each)
                                            content1=''.join(content2)          
                                            content2 = content1.split('QC')
                                            qc_value = re.findall("(\d.\d{2,3})",content2[1])
                                            content2 += qc_value                                           
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).Value=now_2[0]
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).NumberFormat= "m/d/yyyy"
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(3,maxcolumn+k).Value='QC'
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).Value=float(content2[2])
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).NumberFormat="0.000"
                                            k+=1                                           
                                    count+=1
                            page_number+=1
                        else:
                            while count <lenth:
                                if ('QC' in a[count]) and ('*' not in a[count])and ('QCQ' not in a[count]):                             
                                    if ('/' not in a[count+1]) and (float(a[count+1])>float(value)):
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).Value=now_2[0]
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(2,maxcolumn+k).NumberFormat= "m/d/yyyy"
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(3,maxcolumn+k).Value='QC'
                                            excel.Sheets("TC_XMN_CHM_F_Q.67E").Cells(4,maxcolumn+k).Value=float(a[count+1])
                                            k+=1
                                count+=1
                            page_number+=1
option=g.ccbox('which test item',title='UV QC data input', choices=('Cr','HCHO'))
if option==1:
        print ('Cr')
        UV_QC('Cr-VI','QC Chart_Cr_66-01-2013-011 CARY100.xlsx',0.035,'66-01-2013-011 UV-Vis (100)')
else:
        print ('Formal')
        UV_QC('formal','QC Chart_HCHO_66-01-2016-051 CARY60.xlsx',1,'66-01-2016-051 UV-Vis (60)')
g.msgbox('Finish ')

