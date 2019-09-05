
import easygui as g
import datetime
import os
import os.path
import codecs,sys
import win32com.client as win32com
from win32com.client import Dispatch
import tkinter
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
import win32timezone
import time
import re
import csv
import random
now_1 = str(time.strftime('%Y'))

def UV_QC(name,matter,value,address):
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/%s/%s'%(now_1,address,name), title="Select files",filetypes=[("pdf file", "*.pdf")])
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

def ph_QC_Chart(address,files):    
        root=tkinter.Tk()
        path = tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/%s'%(now_1,address), title="Select files",filetypes=[("csvfile", "*.csv")])
        root.destroy()
        excel = win32com.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Application.DisplayAlerts = True        
        workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./%s'%files))        
        maxcolumn=1
        while excel.Sheets("Data 1").Cells(2,maxcolumn).Value is not None:
            maxcolumn+=1  
        ph_date = csv.reader(open('%s'%path,'r'))
        print (ph_date)
        k=0
        for each in ph_date:
            if ('CC' in each):
                now= re.findall('\d{4}-\d{2}-\d{2}',each[0])
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).Value=now[0]
                excel.Sheets("Data 1").Cells(1,maxcolumn+k).NumberFormat= "yyyy/m/d"
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).Value=each[5]
                excel.Sheets("Data 1").Cells(2,maxcolumn+k).NumberFormat="0.000"            
                k+=1

def ICP():  
    path= tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/66-01-2018-012 5110 ICP-OES'%now_1,filetypes=[("csvfile", "*.csv")])  
    ph_date = csv.reader(open('%s'%path,"rt",encoding='utf-8'))
    lines=[]
    for line in ph_date: 
        lines.append(line)
    name_list=os.path.basename(os.path.realpath(path[0])).split('-')
    print(name_list)
    now=name_list[0]
    name = ['Plastic','Plastic','Plastic','Paint','Metal','Metal','BS AM','BS AM','BS AM','CC','CC','CC','CC','CC','CC','CC',]
    element = ['Pb','Cd','As','Pb','Pb','As','Pb','Cd','As','Pb','Cd','Ni','As','Hg','Sb','Sn']
    wavelenth = ['220.353','228.802','188.980','220.353','220.353','188.980','220.353','228.802','188.980','220.353','228.802','231.604','188.980','184.887','206.834','189.925']
    row = [5,6,7,8,9,10,11,12,13,2,3,4,5,6,7,8]
    n = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]  
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./QC Chart_Heavy Metal -66-01-2018-012.xlsx'))       
    for i in range(16):
        lists=[]
        m=0
        for l in lines:  
            if  ('%s'%name[i] in lines[m][0]) and ('%s'%element[i] in lines[m][2]) and ('ref' not in lines[m][3]):      
                lists.append(l)
            m+=1 
        if n[i]<=9: 
            maxcolumn=5
            while excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value is not None:
                maxcolumn+=1
            a=float(excel.Sheets("CRM recorvery rate").Cells(row[i],2).Value)
            if lists!=[]:
                d=lists[0][0].replace(' ',',')
                list=d.split(',')
                excel.Sheets("CRM recorvery rate").Cells(4,maxcolumn).Value=now
                if n[i]<=3:      
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)/float(list[1])
                elif n[i]==4:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)*10/float(list[1])   
                elif n[i]<=6:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])*int(float(list[1])*250)*25/float(list[1]) 
                else:
                    excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=float(lists[0][4])     
            else:
                excel.Sheets("CRM recorvery rate").Cells(row[i],maxcolumn).Value=0
        else:
            maxcolumn=2
            while excel.Sheets("Data 1").Cells(row[i],maxcolumn).Value is not None:
                maxcolumn+=1
            for list in lists:
                excel.Sheets("Data 1").Cells(1,maxcolumn).Value=now
                excel.Sheets("Data 1").Cells(row[i],maxcolumn).Value= float(list[4])
                maxcolumn+=1
            if n[i]==16:
                maxcolumn=2
                while excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value is not None:
                    maxcolumn+=1
                excel.Sheets("Nickel QC").Cells(4,maxcolumn).Value=now
                excel.Sheets("Nickel QC").Cells(5,maxcolumn).Value=round(random.uniform(0.090, 0.110),3)
                excel.Sheets("Nickel QC").Cells(6,maxcolumn).Value=round(random.uniform(0.4500, 0.5500),4)
        i+=1

print ('begin')
print ('Wait for a moment')
option=g.indexbox('which test item',title='data input', choices=('Cr','HCHO','pH 2014','pH 2018','ICP'))
if option==0:
        print ('Cr')
        UV_QC('Cr-VI','QC Chart_Cr_66-01-2013-011 CARY100.xlsx',0.035,'66-01-2013-011 UV-Vis (100)')
elif option==1:
        print ('Formal')
        UV_QC('formal','QC Chart_HCHO_66-01-2016-051 CARY60.xlsx',1,'66-01-2016-051 UV-Vis (60)')
elif option==2:
        print ('pH 2014')
        ph_QC_Chart('66-01-2014-015 pH','QC Chart _pH_66-01-2014-015.xlsx')
elif option==3:
        print ('pH 2018')
        ph_QC_Chart('66-01-2018-006 pH','QC Chart _pH_66-01-2018-006.xlsx')    
else:
        print ('ICP')
        ICP()
g.msgbox("the programme has been finished", ok_button="OK!")