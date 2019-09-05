#!/usr/bin/env python
# -*- coding: utf-8 -*-


 
import tkinter as tk  # 使用Tkinter前需要先导入
import tkinter.messagebox
import pickle
import openpyxl
import easygui as g
import os.path
import win32com.client as win32com
import time
import datetime
import calendar
import pyautogui
time_1 = int(time.strftime('%m'))
time_2=calendar.month_abbr[time_1]
address=os.path.abspath('.')
# 第1步，实例化object，建立窗口window
window = tk.Tk()
 
# 第2步，给窗口的可视化起名字
window.title('作者：Frank Chen')
 
# 第3步，设定窗口的大小(长 * 宽)
window.geometry('421x265')  # 这里的乘是小x

#第一行
l_1 = tk.Label(window, text='pH填写', font=(16), height=1)
l_1.grid(column=2,row=1) 
#第二行
method='4045'
def select_method():
    global method
    if (var1.get() == 1) :     # 如果选中第一个选项，未选中第二个选项
        method='3071'
        print(method)
    elif (var1.get() == 2):
        method='4045'
        print(method)
var1 = tk.IntVar()  # 定义var1和var2整型变量用来存放选择行为返回值
var1.set(2)

r1= tk.Radiobutton(text='3071', variable=var1, value=1,command=select_method)
r1.grid(column=1, row=2)
r2= tk.Radiobutton(text='4045', variable=var1, value=2,command=select_method)
r2.grid(column=3, row=2)

#第三行

l_2= tk.Label(window, text='结果1', font=(14))
l_2.grid(column=1, row=3) 
l_3= tk.Label(window, text='结果2', font=(14))
l_3.grid(column=2, row=3) 

l_4= tk.Label(window, text='稀释1', font=(14))
l_4.grid(column=3, row=3) 
l_5= tk.Label(window, text='稀释2', font=(14))
l_5.grid(column=4, row=3) 

#第四行
e_1 = tk.StringVar()
e_1_name = tk.Entry(window, textvariable=e_1, width=10, font=(14))
e_1_name.grid(column=1, row=4)
e_1_name.focus()
e_2 = tk.StringVar()
e_2_name = tk.Entry(window, textvariable=e_2, width=10, font=(14))
e_2_name.grid(column=2, row=4)
e_3 = tk.StringVar()
e_3_name = tk.Entry(window, textvariable=e_3, width=10, font=(14))
e_3_name.grid(column=3, row=4)
e_4 = tk.StringVar()
e_4_name = tk.Entry(window, textvariable=e_4, width=10, font=(14))
e_4_name.grid(column=4, row=4)
#第五行
blk='No BLK'
def select_blk():
    global blk,result_5
    if (var2.get() == 1) : 
        blk='No BLK'
        e_6.set('9')
    elif (var2.get() == 2): 
        blk='BLK'
        e_6.set('10')

var2 = tk.IntVar()  # 定义var1和var2整型变量用来存放选择行为返回值
var2.set(1)
r1= tk.Radiobutton(text='No BLK', variable=var2, value=1,command=select_blk)
r1.grid(column=1, row=5)
r2= tk.Radiobutton(text='BLK', variable=var2, value=2,command=select_blk)
r2.grid(column=3, row=5)
e_5 = tk.StringVar()
e_5_name = tk.Entry(window, textvariable=e_5, width=10, font=(14))
e_5_name.grid(column=4, row=5)



#第六行
l_6= tk.Label(window, text='←', font=(14))
l_6.grid(column=1, row=6) 
l_7= tk.Label(window, text='向左移动', font=(1))
l_7.grid(column=2, row=6) 
e_6 = tk.StringVar()
e_6_name = tk.Entry(window, textvariable=e_6, width=10, font=(14))
e_6_name.grid(column=3, row=6)
e_6.set('10')
#result_5=e_6.get()

#第七行
l_8= tk.Label(window, text='实验室温度', font=(14))
l_8.grid(column=1, row=7) 
l_9= tk.Label(window, text='溶液温度', font=(1))
l_9.grid(column=3, row=7) 
e_7 = tk.StringVar()
e_7_name = tk.Entry(window, textvariable=e_7, width=10, font=(14))
e_7_name.grid(column=2, row=7)
e_8 = tk.StringVar()
e_8_name = tk.Entry(window, textvariable=e_8, width=10, font=(14))
e_8_name.grid(column=4, row=7)


#第八行
l_10= tk.Label(window, text='请再次\n确认信息\n是否填全', font=(1))
l_10.grid(column=2, row=8) 
def auto_result():
    time.sleep(3)
    if blk=='BLK':
            pyautogui.typewrite('%s'%e_5.get()) 
            pyautogui.typewrite(['right']*7)
    pyautogui.typewrite('%s'%e_7.get())
    pyautogui.typewrite(['right'])
    pyautogui.typewrite('%s'%e_8.get())
    pyautogui.typewrite(['up']*2)
    time.sleep(1)
    pyautogui.typewrite(['left']*int(e_6.get()))
    pyautogui.typewrite('%s'%e_1.get())
    pyautogui.typewrite(['down'])
    pyautogui.typewrite('%s'%e_2.get()) 
    pyautogui.typewrite(['down']*3)
    print(method)
    if method=='3071':
        pyautogui.typewrite('0')
    elif method=='4045':
        a=((float(e_3.get())+float(e_4.get()))-(float(e_1.get())+float(e_2.get())))/2
        pyautogui.typewrite('%s'%float('%.2f'%a))
b_1= tk.Button(window, text='自动\n填写', font=('Arial', 12), height=2, command=auto_result)
b_1.grid(column=3,row=8)
window.mainloop()