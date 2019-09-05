#!/usr/bin/env python
# -*- coding: utf-8 -*-
import time
import pyautogui
import os
import os.path
import openpyxl
import easygui as g
from win32com.client import Dispatch
import win32com.client as win32com
import tkinter
import tkinter as tk  # 使用Tkinter前需要先导入
from tkinter import scrolledtext
now_1 = str(time.strftime('%Y'))


# 第1步，实例化object，建立窗口window
window = tk.Tk()
 
# 第2步，给窗口的可视化起名字
window.title('操作界面，作者：Frank Chen')
 
# 第3步，设定窗口的大小(长 * 宽)
window.geometry('535x320')  # 这里的乘是小x
window.resizable(0, 0)
#第一行
l_1 = tk.Label(window, text='reach 自动填写结果', font=('Arial', 16,'bold'), width=30, height=2)
l_1.place(x=0, y=0) 

#第一列
t_1 = tk.scrolledtext.ScrolledText(window,width=30,height=15)
t_1.insert('end','显示输入的编号信息：\n')
t_1.place(x=5, y=50)

#第二列，第一行
l_2 = tk.Label(window, text='请输入起始编号：', font=('Arial',14), height=2)
l_2.place(x=300, y=50)


var_usr_name = tk.StringVar()
var_usr_name.set('2')
entry_usr_name = tk.Entry(window, textvariable=var_usr_name,  width=17,font=('Arial', 14))
entry_usr_name.place(x=300, y=100)
entry_usr_name.focus()

#第二列，第二行
def select_file():
    global path,list,id_number
    path=tkinter.filedialog.askopenfilenames(initialdir='Z:/Data/%s/66-01-2018-012 5110 ICP-OES/SVHC'%now_1,filetypes=[("file", "*.xls*")]) 
    list=[]
    id_number=[]
    pyautogui.PAUSE=1
    pyautogui.FAILSAFE= True 
    excel = win32com.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.Application.DisplayAlerts = True   
    workbook = excel.Workbooks.Open(os.path.join(os.getcwd(),r'%s'%path))
    row=5
    x=117
    #t_1.delete(1.0, tkinter.END)
    #t_1.insert('end','显示输入的编号信息：\n')
    t_1.insert('end','%s\n'%os.path.basename(os.path.realpath(path[0])))
    t_1.see(tk.END)
    window.update()
    
    while excel.Sheets('SVHC').Cells(x,11).Value is not None:
        x+=1  
    for i in range(x-5):
        result=round(excel.Sheets("SVHC").Cells(row,11).Value,3)
        id_no=excel.Sheets("SVHC").Cells(row,1).Value
        list.append(result)
        id_number.append(id_no)
        row+=1    
    window.update()

b_1= tk.Button(window, text='选择结果文件', font=('Arial', 12), width=20, height=1, command=select_file)
b_1.place(x=300, y=145)


#第二列，第三行

def get_data():
    v.set(2)
    t_1.delete(1.0, tkinter.END)
    t_1.insert('end','显示输入的编号信息：\n')
    t_1.insert('end','%s\n'%os.path.basename(os.path.realpath(path[0])))
    t_1.insert('end','即将输入结果的编号：%s \n'%float(var_usr_name.get()))
    t_1.see(tk.END)
    window.update() 
    fill_sample_number=float(var_usr_name.get())
    m=id_number.index(fill_sample_number)
    time.sleep(10)
    len_id_number=set(id_number)
    window.update() 
    for n in range(len(len_id_number)-1):
        window.update() 
        if v.get()==1:
            t_1.insert('end','============================\n')
            break
        else:
            print(m+1,len(list),len(id_number),n,len(len_id_number),range(len(len_id_number)-1))
            if m+1<len(list):
                while id_number[m]==id_number[m+1]:
                    if str(list[m])=='-2146826273':
                        m+=1                   
                    else:
                        if str(list[m+1])=='-2146826273':
                            list[m+1]=list[m]
                        else:
                            list[m+1]=min(list[m],list[m+1])
                    m+=1 
                if list[m]<0:
                    list[m]=0 
                pyautogui.typewrite('%s'%list[m],0.0000000001)
                #t_1.insert('end', '编号%s结果：%s\n'%(id_number[m],list[m]))
                t_1.insert('end', '结果：%s\n'%list[m])
                window.update()
                pyautogui.typewrite(['Enter']) 
                t_1.insert('end', '即将输入结果的编号：%s\n'%id_number[m+1])
                t_1.see(tk.END)
                window.update()        
                #time.sleep(3)
            else:
                pyautogui.typewrite(['right']*6)
                pyautogui.typewrite('%s'%id_number[m],0.0000000001)
                t_1.insert('end','============================\n')
            m+=1
b_2= tk.Button(window, text='确认开始', font=('Arial', 12), width=20, height=1, command=get_data)
b_2.place(x=300, y=195)

#第二列，第四行
v = tk.IntVar()
v.set(2)
rb1 = tk.Radiobutton(window, text='停止', variable=v, value=1, indicatoron=False, width=10, height=2)
rb1.place(x=300, y=245)
rb2 = tk.Radiobutton(window, text='继续', variable=v, value=2, indicatoron=False, width=10, height=2)
rb2.place(x=430, y=245)
window.update
window.mainloop()
