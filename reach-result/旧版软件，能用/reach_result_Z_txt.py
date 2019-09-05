#!/usr/bin/env python
# -*- coding: cp936 -*-
import os
import os.path
import time
import easygui as g
from win32com.client import Dispatch
import win32com.client as win32com
import random
now_1 = str(time.strftime('%Y'))
def get_list(lab_number,element,wavelenth,row,n):  
    name = lab_number.replace('/', '_') 
    f2 = open(path,"r")
    lines = f2.readlines()
    list=[]
    for line in lines:
        if ('%s'%lab_number in line) and ('%s'%element in line) and ('(ref)' not in line):
            line=line.replace(',','\t')
            list.append(line.split('\t'))   
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True    
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'./SVHC.xls'))
    excel.Sheets("1").Cells(row,15).Value=element
    if str(list[0][4])=='    \xce\xb4\xd0\xa3\xd5\xfd':
        excel.Sheets("1").Cells(row,16).Value='未校正'
    else:
        excel.Sheets("1").Cells(row,16).Value=float(list[0][4])*250
    if n=='14':
        workbook.SaveAs('Z:\\Inorganic_batch\\Microwave\\result\\%s.xls'%name)
        excel.Application.Quit()
def main():    
    lab_number = raw_input('please fill in the lab number:')
    print 'begin'
    print 'Wait for a moment'
    get_list('%s'%lab_number,'B','249.772','4','1')
    get_list('%s'%lab_number,'As','188.980','5','2')
    get_list('%s'%lab_number,'Sb','206.834','6','3')
    get_list('%s'%lab_number,'Cr','267.716','7','4')
    get_list('%s'%lab_number,'Pb','220.353','8','5')
    get_list('%s'%lab_number,'Co','228.615','9','6')
    get_list('%s'%lab_number,'Ba','230.424','11','7')
    get_list('%s'%lab_number,'Sn','189.925','12','8')
    get_list('%s'%lab_number,'Mo','202.032','13','9')
    get_list('%s'%lab_number,'Zr','343.823','14','10')
    get_list('%s'%lab_number,'Al','236.705','15','11')    
    get_list('%s'%lab_number,'Cd','228.802','17','12')
    get_list('%s'%lab_number,'Na','589.592','18','13')
    get_list('%s'%lab_number,'Sr','460.733','20','14')
    
path=g.fileopenbox(msg=None,title=None,default='Z:/Data/%s/66-01-2013-012 ICP-OES'%now_1,filetypes=None)
main()

action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))
while  action==1:
    main()
    action=g.ccbox('whether need to run the program again', choices=('continue','finsh'))       
else:    
    os.startfile('Z:\\Inorganic_batch\\Microwave\\result') 



