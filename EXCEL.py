import tkinter as tk
from tkinter import*
import  pandas
import tkinter.filedialog
import os
import xlrd
import numpy
dict_1={'路径':' ','年':'0','月':'0','日':'0'
        ,'存储路径':' '}
global k,j,i,h
class interfaceclass:
    window_1=tk.Tk()
    window_1.title('考勤自动读取')
    window_1.geometry('500x300')
    
    global var100
    global var101
    def slecpath1():
         path_1=tkinter.filedialog.askopenfilename() #选择文件
         #path_=tkinter.filedialog.askdirectory()#选择路径
         print(path_1)
         path_1=path_1.replace("/","\\\\")
         dict_1['路径']=path_1
         var100.set(path_1) 

    def slecpath2():
        path_2=tkinter.filedialog.askopenfilename()
        path_2=path_2.replace("/","\\\\")
        dict_1['存储路径']=path_2
        var101.set(path_2)
        print(dict_1['存储路径'])

    def read_excel():
        print(dict_1['路径'])
        DB_value=pandas.DataFrame(columns=('姓名', '日期', '时间'))#创建一个EXCEL表
        DB_1=pandas.DataFrame(columns=('姓名', '日期', '时间'))
        DB_excel=pandas.read_excel(dict_1['路径'],sheet_name=0)
        DB_excel.loc[:,'日期']=0#增加一个空列
        DB_excel.loc[:,'时间']=0#增加一个空列
        DB_value['日期']=DB_excel['日期时间'].dt.date
        DB_value['时间']=DB_excel['日期时间'].dt.time
        DB_value['姓名']=DB_excel['姓名']
        list_name=DB_value['姓名'].unique()#统计人名
        list_time=DB_value['日期'].unique()#统计天数
        k=len(list_name)#统计人数
        j=len(list_time)#统计天数
        i=0#天数
        h=0#人数
        print(k)
        print(j)
        while h<k:
            df1=df1=DB_value[(DB_value['姓名']==list_name[h])]
            while i<j:
                df2=df1[df1['日期']==list_time[i]]
                a=df2['时间'].min()
                b=df2['时间'].max()
                df10=df2[(df2['时间']==a)|(df2['时间']==b)]
                DB_1=pandas.concat([DB_1,df10])
                i=i+1
            if i == 26 :
               h=h+1
               i=0
        print(DB_1)
        DB_1.to_excel(dict_1['存储路径'],sheet_name='sheet1')#写入EXCEL 
           #df1=DB_value[(DB_value['姓名']==list_name[0])&(DB_value['日期']==list_time[i])]
           #a=df1['时间'].min()
           #b=df1['时间'].max()
           #df100=df1[(df1['时间']==a)|(df1['时间']==b)]
           #DB_1=pandas.concat([DB_1,df100])
           #i=i+1
          
        print(DB_1)
        
        #DB_excel_list_name[0]=DB_excel[DB_excel['时间']<'8:30:00']
        #print(DB_excel_1)
        #print(list_time)
       
    label_1=tk.Label(window_1,text='目标路径',width=10,height=1,anchor='w',justify='right')#row行，colum列
    label_1.place(x = 1,y = 10,anchor='nw')

    label_2=tk.Label(window_1,text='存储路径',width=10,height=1,anchor='w',justify='right')#row行，colum列
    label_2.place(x = 1,y = 100,anchor='nw')

    var100 = tk.StringVar()#地址显示
    label_100=tk.Label(window_1,textvariable=var100,width=40,height=1,bg='white',anchor='w',justify='right')#row行，colum列
    label_100.place(x = 60,y = 10,anchor='nw')

    var101 = tk.StringVar()#存储地址显示
    label_100=tk.Label(window_1,textvariable=var101,width=40,height=1,bg='white',anchor='w',justify='right')#row行，colum列
    label_100.place(x = 60,y = 100,anchor='nw')


    button1= tk.Button(window_1,text='路径选择',height=1,width=10,command=slecpath1)
    button1.place(x = 370,y = 10,anchor='nw')

    button2= tk.Button(window_1,text='转换',height=1,width=10,command=read_excel)
    button2.place(x = 370,y = 50,anchor='nw')

    button3= tk.Button(window_1,text='存储选择',height=1,width=10,command=slecpath2)
    button3.place(x = 370,y = 100,anchor='nw')

    






    window_1.mainloop()