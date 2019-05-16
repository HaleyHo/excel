#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog 
from project3 import BCT
from project2 import wd

class UI_dataanalysis_calss:
    def __init__(self, window):
        self.window = window
		#界面的标题及尺寸不可变
        self.bct=BCT()
        self.wd=wd()
        self.window.title('数据处理软件')
        self.window.geometry('+700+500')
        self.window.resizable(False, False)
        self.Create_window()

    def Create_window(self):
        self.election_file_path_1 = tk.StringVar()
        self.election_file_path_2 = tk.StringVar()
        self.save_file_path = tk.StringVar()
        self.search = tk.StringVar()
		#创建一个一级窗口
        self.labelframe_1 = tk.LabelFrame(self.window,  text='文件处理：',padx=20,pady=10)
        self.labelframe_1.pack()#pady=10
        self.labelframe_2 = tk.LabelFrame(self.window,  text='选择的选项：',padx=10,pady=10)
        self.labelframe_2.pack()
        self.labelframe_3 = tk.LabelFrame(self.window,  text='请选择需要生成的表格：',padx=10,pady=10)
        self.labelframe_3.pack()
        
		#原始数据文件路径选取
        tk.Label(self.labelframe_1,  text = 'DRP',  background='PowderBlue').grid(row = 0,  column = 0)
        tk.Entry(self.labelframe_1,  textvariable = self.election_file_path_1).grid(row = 0,  column = 1)
        self.open_file_path_button = tk.Button(self.labelframe_1,  text='路径选择',rcommand=self.election_file_directory_1)
        self.open_file_path_button.grid(row=0,  column=2)

        #原始数据文件路径选取
        tk.Label(self.labelframe_1,  text = '请上传key in的文件',  background='PowderBlue').grid(row = 1,  column = 0)
        tk.Entry(self.labelframe_1,  textvariable = self.election_file_path_2).grid(row = 1,  column = 1)
        self.open_file_path_button = tk.Button(self.labelframe_1,  text='路径选择',command=self.election_file_directory_2)
        self.open_file_path_button.grid(row=1,  column=2)
       
		#文件保存
        tk.Label(self.labelframe_1,  text='   输入文件名  ',  background='PowderBlue').grid(row=2, column=0)
        tk.Entry(self.labelframe_1,  textvariable = self.save_file_path).grid(row=2,  column=1)
        self.save_file_path_button = tk.Button(self.labelframe_1,  text='文件保存', command=self.save_file_directory)
        self.save_file_path_button.grid(row = 2,  column=2)
        #搜索
        tk.Label(self.labelframe_1,  text='输入关键字',  background='PowderBlue').grid(row=3, column=0)
        search_value = tk.Entry(self.labelframe_1,textvariable = self.search).grid(row=3,  column=1)
        self.search_button = tk.Button(self.labelframe_1,  text='搜索', state='disable',command=self.search_action)
        self.search_button.grid(row = 3,  column=2)
   
		#点击分析处理文件
        self.bit_analysis = tk.Button(self.labelframe_1,  text='点击进行分析',state='disable', command=self.progressbar_1)
        self.bit_analysis.grid(row = 4, column = 2)
        #点击生成表格
        self.button_2 = tk.Button(self.labelframe_1,  text='生成表格',state='disable', command=self.sheet)        
        self.button_2.grid(row=4,  column=0)
       
		#判断文件路径是否都已备录      
        self.judge_1 = 0 #打开文件路径
        self.judge_2 = 0 #保存文件路径
        self.judge_3 = 0 #上面两者是否都已选择
        self.judge_4 = 0 #条件全选或全不选按键
        self.confirm_list = [] 
    #文件路径的变量保存
    def election_file_directory_1(self):		
        self.election_file_path_1.set(filedialog.askopenfilename(filetypes=[('excel file', '.xlsx')]))
        if self.election_file_path_1.get() != '':
            self.judge_1 = 1
    #key in文件路径的变量保存
    def election_file_directory_2(self):        
        self.election_file_path_2.set(filedialog.askopenfilename(filetypes=[('excel file', '.xlsx')]))

    #文件生成路径的保存同时将图表生成按键激活
    def save_file_directory(self):
        self.save_file_path.set(filedialog.asksaveasfilename(filetypes=[('excel file', '.xlsx')]))
        if self.save_file_path.get() != '':			
            self.judge_2 = 1		
        self.judge_3 = self.judge_1+self.judge_2		
        if self.judge_3 == 2:			
            self.bit_analysis.config(state='normal')
            self.search_button.config(state='normal')
        else:			
            self.judge_3 = 0

    def search_action(self):
        values = self.search.get()       
        if values == '':
            tk.messagebox.showerror(title='error', message='请输入需要搜索的内容')
            return           
        else:
            list_component_search = []
            #进行数据分析，生成list_component
            if self.election_file_path_2.get()=='':
                data_analysis_judge_1 = self.wd.ddlist(self.election_file_path_1.get())
            else:
                data_analysis_judge_1 = self.bct.ddlist(self.election_file_path_1.get(),self.election_file_path_2.get())
            for i in data_analysis_judge_1:
                if str.lower(values) in str.lower(i):
                    list_component_search.append(i)
            if list_component_search == []:
                tk.messagebox.showerror(title='error', message='没有搜索到该值,请重新输入')
            else:
                self.labelframe_3.destroy()
                self.labelframe_3 = tk.LabelFrame(self.window,  text='请选择需要生成的表格：',padx=10,pady=10)
                self.labelframe_3.pack()  
                #创建多选项变量储存列表和多选项选键的储存列表
                self.var_1 = []
                self.checkbutton = []
                #从第四行开始，每生成5个可选项换下一行
                for i in range(len(list_component_search)):         
                    if i == 0:              
                        row_1 = 3               
                        column_1 = 0
                    else:                               
                        row_1 = int(i/5)+3              
                        column_1 = i%5          
                    self.var_1_1 = tk.StringVar()           
                    self.var_1.append(self.var_1_1)         
                    elect_1 = tk.Checkbutton(self.labelframe_3,  text=list_component_search[i], fg='MidnightBlue', variable=self.var_1[-1],  onvalue=list_component_search[i],  offvalue='off')         
                    self.checkbutton.append(elect_1)            
                    elect_1.grid(row=row_1, column=column_1)          
                    elect_1.deselect()    
                #创建条件全选快捷按键及全部不选的变更同时生成下一步数据可视化化的按键,同时激活生成表格按键      
                self.button_1 = tk.Button(self.labelframe_3,  text='条件全选',command = self.selectall)      
                self.button_1.grid(row=row_1+1, column=0)
                self.button_confirm = tk.Button(self.labelframe_3,  text='确定', command=self.confirm)        
                self.button_confirm.grid(row=row_1+1,  column=3)
                # tk.Label(self.labelframe_2,  text='选择的选项：',  background='PowderBlue').grid(row=5, column=0,sticky=W)

    def progressbar_1(self):
        self.labelframe_3.destroy()
        self.labelframe_3 = tk.LabelFrame(self.window,  text='请选择需要生成的表格：',padx=10,pady=10)
        self.labelframe_3.pack()   
        #进行数据分析，生成list_component
        if self.election_file_path_2.get()=='':

            data_analysis_judge_1 = self.wd.ddlist(self.election_file_path_1.get())
        else:
            data_analysis_judge_1 = self.bct.ddlist(self.election_file_path_1.get(),self.election_file_path_2.get())
        #创建多选项变量储存列表和多选项选键的储存列表
        self.var_1 = []
        self.checkbutton = []
		#从第四行开始，每生成5个可选项换下一行
        for i in range(len(data_analysis_judge_1)):			
            if i == 0:				
                row_1 = 3				
                column_1 = 0
            else:								
                row_1 = int(i/5)+3				
                column_1 = i%5			
            self.var_1_1 = tk.StringVar()			
            self.var_1.append(self.var_1_1)			
            elect_1 = tk.Checkbutton(self.labelframe_3,  text=data_analysis_judge_1[i], fg='MidnightBlue', variable=self.var_1[-1],  onvalue=data_analysis_judge_1[i],  offvalue='off')			
            self.checkbutton.append(elect_1)			
            elect_1.grid(row=row_1, column=column_1)			
            elect_1.deselect()    
		#创建条件全选快捷按键及全部不选的变更同时生成下一步数据可视化化的按键,同时激活生成表格按键		
        self.button_1 = tk.Button(self.labelframe_3,  text='条件全选',command = self.selectall)		
        self.button_1.grid(row=row_1+1, column=0)
        self.button_confirm = tk.Button(self.labelframe_3,  text='确定', command=self.confirm)        
        self.button_confirm.grid(row=row_1+1,  column=3)
        # tk.Label(self.labelframe_1,  text='选择的选项：',  background='PowderBlue').place(x=20, y=10, anchor='sw')
	#条件全选按键或全部不选按键 
    def selectall(self):		
        if self.judge_4 == 0:			
            for i in range(len(self.checkbutton)):				
                self.checkbutton[i].select()				
                self.button_1.config(text='全部不选')				
                self.judge_4 = 1		
        else:			
            for i in range(len(self.checkbutton)):				
                self.checkbutton[i].deselect()				
                self.button_1.config(text='条件全选')				
                self.judge_4 = 0 
    def confirm(self): 
        self.button_2.config(state='normal')      
        for i in self.var_1:
            if i.get()!='off':              
                if i.get() not in self.confirm_list:
                    self.confirm_list.append(i.get())
        for i in range(len(self.confirm_list)):         
            if i == 0:              
                row_1 = 3               
                column_1 = 0
            else:                               
                row_1 = int(i/5)+3              
                column_1 = i%5  
            tk.Label(self.labelframe_2,  text='['+self.confirm_list[i]+']').grid(row=row_1, column=column_1)
                
    def sheet(self): 
        self.button_2.config(state='disable')     
        #如果未选中选项，警告提示未选
        if len(self.confirm_list) == 0:
            tk.messagebox.showerror(title='error', message='请选择选项')
        else:
            if self.election_file_path_2.get()=='':
                self.wd.bb(self.election_file_path_1.get(),self.save_file_path.get(),self.confirm_list)
            else:
                self.bct.wexcel(self.election_file_path_1.get(),self.election_file_path_2.get(),self.save_file_path.get(),self.confirm_list)
            tk.messagebox.showinfo(title='successful', message='数据分析完成，表格已生成')
            #生成选择再次运行和退出程序选项
            self.labelframe_4 = tk.LabelFrame(self.window,  text='表格已生成，选择再次运行程序或是退出程序：',padx=0,pady=10)
            self.labelframe_4.pack()#pady=10
            self.button_3 = tk.Button(self.labelframe_4, text='再次运行', state='disable', command=self.again_run)
            self.button_3.grid(row=0, column=0, padx=50)
            self.button_4 = tk.Button(self.labelframe_4, text='退出程序', state='disable', command=lambda:exit())
            self.button_4.grid(row=0, column=2, padx=70)
            self.button_3.config(state='normal')           
            self.button_4.config(state='normal')  
	
    def again_run(self):		
        self.labelframe_1.destroy()		
        self.labelframe_2.destroy()		
        self.labelframe_3.destroy()	
        self.labelframe_4.destroy() 	
        self.Create_window()
        
if __name__=='__main__':
	window = tk.Tk()
	ui_1 = UI_dataanalysis_calss(window)
	window.mainloop()