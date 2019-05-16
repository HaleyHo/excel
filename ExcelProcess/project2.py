#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 12 08:43:02 2018

@author: f7689594
"""

import tkinter as tk
import tkinter.filedialog as tf
import pandas as pd
import datetime
import xlsxwriter

class wd:
    def __init__(self):
        self.path=tk.StringVar()
        self.path_=''
        self.path_save=''
        self.save_var=tk.StringVar()
#     def selectPath(self):
#         self.path_ = tf.askopenfilename()
#         self.path.set(self.path_)
# #        print(self.path_)
#     def save(self):
#         self.path_save = tf.asksaveasfilename()
#         self.save_var.set(self.path_save)
    
#     def wn(self):
#         root.title('表格轉換')
#         root.geometry('310x85')
#         tk.Label(root,text = "選擇文件:").grid(row = 0, column = 0)
#         tk.Entry(root, textvariable = self.path).grid(row = 0, column = 1)
#         tk.Button(root, text = "...", command = self.selectPath).grid(row = 0, column = 2)
        
#         self.e_gname=tk.Entry(root,textvariable = self.save_var)
#         self.e_gname.grid(row = 1, column = 1)
#         tk.Label(root,text = "選擇保存路徑:").grid(row = 1, column = 0)
#         tk.Button(root, text = "...", command = self.save).grid(row = 1, column = 2)
#         tk.Button(root, text = "確認保存", command = self.bb).grid(row = 2, column = 1)
    def ddlist(self,url):
        # url=self.path_
        
        # self.generatename=self.e_gname.get()
        d = pd.read_excel(url, sheetname='FATP')
        d=d.dropna(axis=1,how='all')
        d=d.dropna(axis=0,how='all')
        d.index=range(len(d))
        d.columns=d.loc[0]
        str_1=''
        for j in d.loc[0]:
                str_1=j
                break   
        for i in range(len(d)):
            if d.loc[i,str_1]!='Component':
                d= d.drop(i)
            else:
                break
        self.str_1=str_1
        d.index=range(len(d))
        d.columns=d.loc[0]
        d=d.drop(0)#找到Component所在行，并删除之前到行，在重新遍历索引后，将第一行设为columns，然后删除第一行
        d.index=range(len(d))#重新遍历索引，完成导入表格
        d=d[['Component','Wifi Per','Cell Per','Color','Config','Vendor','Ship Qty','ETA','Alloc','NG','Fail']]#筛选有用到信息
        d[['Component','Color','Config','Vendor','ETA','Alloc','NG','Fail']]=d[['Component','Color','Config','Vendor','ETA','Alloc','NG','Fail']].fillna('')#将nan填充为‘’
        d=d[d['ETA']!='']#过滤掉日期为空到行
        d.index=range(len(d))
        df=pd.DataFrame()
        for i in range(len(d)):
            list_1=[]
            list_2=[]
            list_Alloc = []
            list_NG = []
            str_2=d.loc[i,'Component']
            if d.loc[i,'Cell Per']>0 and d.loc[i,'Wifi Per']>0:
                str_2=d.loc[i,'Component']
            elif d.loc[i,'Wifi Per']>0:
                str_2=str_2+'  (WF)'
            elif d.loc[i,'Cell Per']>0:
                str_2 = str_2 + '  (Cell)'
            if d.loc[i,'Color']!='':
                str_2 = str_2 + '  (' +  str(d.loc[i,'Color'])  + ')'#字符串拼接
            a=datetime.datetime.strptime(d.loc[i,'ETA'],'%m/%d/%Y')
            a=a.strftime('%Y-%m-%d')
            list_1.append(a)
            list_2.append(d.loc[i,'Ship Qty'])
            if d.loc[i,'Alloc']:
                list_Alloc.append(d.loc[i,'Alloc'])
            else:
                list_Alloc.append(0)
            num= 0
            if d.loc[i,'NG']:
                num += int(d.loc[i,'NG'])
            if d.loc[i,'Fail']:
                num += int(d.loc[i, 'Fail'])
            list_NG.append(num)
            for j in range(i+1,len(d)):
                if d.loc[j,'Config']=='' and d.loc[j,'Vendor']=='':
                    a=datetime.datetime.strptime(d.loc[j,'ETA'],'%m/%d/%Y')
                    a=a.strftime('%Y-%m-%d')
                    list_1.append(a)
                    list_2.append(d.loc[j,'Ship Qty'])
                    if d.loc[j, 'Alloc']!='':
                        list_Alloc.append(d.loc[j, 'Alloc'])
                    else:
                        list_Alloc.append(0)
                    num = 0
                    if d.loc[j, 'NG']!='':
                        num += int(d.loc[j, 'NG'])
                    if d.loc[j, 'Fail']!='':
                        num += int(d.loc[j, 'Fail'])
                    list_NG.append(num)
                else:
                    break
            df=df.append(pd.DataFrame([[str_2,d.loc[i,'Vendor'],list_1,list_2,list_Alloc,list_NG]],columns=['A','B','C','D','E','F']))
        df=df[df['A']!='']
        
        df.index=range(len(df))
        self.df=df
        dic_dd={}
        for dd in range(len(df)):
            dic_dd[df.at[dd,'A']]=0
        self.list_dd=list(dic_dd.keys())#这是下拉框列表
        self.d=d
        return self.list_dd

        #拼接字段并保存到df中，零件名，厂商，日期列表，数量列表
        
    def bb(self,url,save_name,list_search=[]):
        self.ddlist(url)
        
        # list_search=['ANT1','Battery','C3 Transfer Flex']#这是勾选项列表
        
        df=self.df
        d=self.d
        
        df=df.T
        df.columns=df.loc['A']
        df=df[list_search]
        df=df.T
        df.index=range(len(df))
        
        df_grouped = df.groupby('A')#对零件名字进行分组
        list_5=[]
        for A,df_g in df_grouped:
            list_8=[]
            list_3_1=[]
            list_4_1=[]
            list_NG_1_1=[]
            list_Alloc_1_1=[]
            df_g_grouped = df_g.groupby('B')#对厂商名字进行分组
            for B,df_g_g in df_g_grouped:
                df_g_g.index=range(len(df_g_g))
                list_3=df_g_g.at[0,'C']
                list_4=df_g_g.at[0,'D']
                list_Alloc_1 = df_g_g.at[0,'E']
                list_NG_1 = df_g_g.at[0,'F']
                for i in range(1,len(df_g_g)):
                    list_3=list_3+df_g_g.at[i,'C']
                    list_4=list_4+df_g_g.at[i,'D']
                    list_Alloc_1 = list_Alloc_1 + df_g_g.at[i,'E']
                    list_NG_1 = list_NG_1 + df_g_g.at[i,'F']
                list_8.append(B)
                list_3_1.append(list_3)
                list_4_1.append(list_4)
                list_Alloc_1_1.append(list_Alloc_1)
                list_NG_1_1.append(list_NG_1)
            list_5.append([A, list_8, list_3_1, list_4_1, list_Alloc_1_1, list_NG_1_1])#将相同零件名下对相同厂商合并在一行并保存到列表list_5中
        list_6=[]
        for l in list_5:
            for i in range(len(l[2])):
                n = 0
                while n<len(l[2][i]):
                    j=n+1
                    while j<len(l[2][i]):
                        if l[2][i][n]==l[2][i][j]:
                            l[3][i][n]=l[3][i][n]+l[3][i][j]
                            l[4][i][n]=l[4][i][n]+l[4][i][j]
                            l[5][i][n]=l[5][i][n]+l[5][i][j]
                            del l[2][i][j]
                            del l[3][i][j]
                            del l[4][i][j]
                            del l[5][i][j]
                            j-=1
                        j+=1
                    n+=1
            list_6.append(l)#将相同日期进行合并相应到数量进行相加
        list_9=[]
        for i in range(len(list_6)):
            list_9_1 = []
            dic_Alloc = {}
            dic_NG = {}
            for j in range(len(list_6[i][2])):
                dic_1 = {}
                for aa in range(len(list_6[i][2][j])):
                    if dic_Alloc.get(list_6[i][2][j][aa]):
                        dic_Alloc[list_6[i][2][j][aa]] += list_6[i][4][j][aa]
                    else:
                        dic_Alloc[list_6[i][2][j][aa]] = list_6[i][4][j][aa]
                    if dic_NG.get(list_6[i][2][j][aa]):
                        dic_NG[list_6[i][2][j][aa]] += list_6[i][5][j][aa]
                    else:
                        dic_NG[list_6[i][2][j][aa]] = list_6[i][5][j][aa]

                    dic_1[list_6[i][2][j][aa]] = list_6[i][3][j][aa]
                list_9_1.append(dic_1)
            list_9.append([list_6[i][0],list_6[i][1],list_9_1,dic_Alloc,dic_NG])
            #将日期和数量相对应放到字典里面，以便于后面到填写操作
        list_7=[]
        for date in d['ETA']:
            list_7.append(datetime.datetime.strptime(date,'%m/%d/%Y'))
        date_max=max(list_7)
        date_min=min(list_7)#找到起始日期和最后一天到日期
        dates = pd.date_range(date_min, date_max)#生成日期数列
        dates=dates.strftime('%Y-%m-%d')#将数列中到日期格式转化为str，
        list_1 = []
        list_2 = []
        for i in range (len(dates)+4):
              a = ord('A') + i
              list_1.append(chr(a))
              if len(list_1) > 26 :             
                      if len(list_1)%26==0:
                              b = len(list_1)//26 
                              c = ord('A') + b - 2
                      else:             
                              b = len(list_1)//26               
                              c = ord('A') + b -1
                      if len(list_1)%26==0: 
                              e = ord('A') + 25                     
                      else:       
                              d = len(list_1) % 26                                     
                              e = ord('A') + d -1               
                      f = chr(c) + chr(e)
                      list_2.append(f)
        list_A=list_1[0:26]+list_2#生成表格表头英文序列（A，B，C...)
        workbook = xlsxwriter.Workbook(save_name+'.xlsx')#开始写表
        ws = workbook.add_worksheet('CTB')
        ws.set_column('A:A',30)
        ws.set_column('B:B',20)
        style_1 = workbook.add_format({'align':'center','valign':'vcenter','bold':True,'font_size':18,'font_name':'Cambria'}) #CTB   居中
        style_2 = workbook.add_format({'font_size':12,'font_name':'Cambria'})#12号
        style_3 = workbook.add_format({'bg_color':'#FFFF00','font_size':12,'font_name':'Cambria','border':7})#黄色12号
        style_4 = workbook.add_format({'bg_color':'#000000','font_size':12,'font_color':'#FFFFFF','font_name':'Cambria','border':7})#黑底白字12号
        style_5 = workbook.add_format({'bg_color':'#C6E2FF','bold':True,'font_size':12,'font_name':'Cambria','border':7})#浅蓝色底12号
        style_6 = workbook.add_format({'bold':True,'font_size':12,'font_name':'Cambria'})#加粗 12号
        style_7 = workbook.add_format({'bg_color':'#FFE1FF','font_size':12,'bold':True,'font_name':'Cambria','border':7})#红色底12号
        style_8 = workbook.add_format({'bg_color':'#FFFF00','align':'center','bold':True,'font_size':12,'font_name':'Cambria'})#加粗黄色居中8
        style_9 = workbook.add_format({'bg_color':'#000000','align':'center','bold':True,'font_size':12,'font_color':'#FFFFFF','font_name':'Cambria'})#加粗黑色居中9
        style_10 = workbook.add_format({'bg_color':'#FFE1FF','align':'center','bold':True,'font_size':12,'font_name':'Cambria'})#加粗红色居中10
        style_11 = workbook.add_format({'align':'center','valign':'vcenter','bold':True,'font_size':12,'font_name':'Cambria'})#白色加粗居中11
        format2 = workbook.add_format({'font_color': '#EE0000'})
        style_12 = workbook.add_format({'bg_color':'#D4D4D4','bold':True,'font_size':12,'font_name':'Cambria','border':7})#加粗 12号
        
        ws.merge_range('A2:B3','CTB',style_1)
        ws.write(1,2,'Vendor',style_6)
        ws.write(0,0,self.str_1)
        for i in range(len(dates)):
            date = datetime.datetime.strptime(dates[i],'%Y-%m-%d') 
            day = date.weekday()
            if day == 6:
                ws.write(1,3+i,dates[i],style_12)
            else:
                ws.write(1,3+i,dates[i],style_6)#这个for循环写入日期
        ws.write(3,1,'Total input',style_6)
        ws.write(4,1,'Cell Black',style_4)
        ws.write(5,1,'Cell White',style_2)
        ws.write(6,1,'Cell Gold',style_3)
        ws.write(7,1,'Cell Pink',style_7)
        ws.write(8,1,'Cell cum',style_6)
        
        ws.write(9,1,'Wifi Black',style_4)
        ws.write(10,1,'Wifi White',style_2)
        ws.write(11,1,'Wifi Gold',style_3)
        ws.write(12,1,'Wifi Pink',style_7)
        ws.write(13,1,'Wifi cum',style_6)
        ws.merge_range('A4:A14','Plan',style_1)
        n=14
        for d in range(len(dates)):
            coord=[]
            x=list_A[3+d]
            ws.write_formula(8,3+d,'=SUM('+x+'5:'+x+'8)',style_6)
            ws.write_formula(13,3+d,'=SUM('+x+'10:'+x+'13)',style_6)
            ws.write_formula(3,3+d,'='+x+'9'+'+'+x+'14',style_6)#写入Total input行到公式
            ws.write(4,3+d,'',style_4)
            ws.write(5,3+d,'',style_2)
            ws.write(6,3+d,'',style_3)
            ws.write(7,3+d,'',style_7)
            ws.write(9,3+d,'',style_4)
            ws.write(10,3+d,'',style_2)
            ws.write(11,3+d,'',style_3)
            ws.write(12,3+d,'',style_7)
        for i in range(len(list_9)):
            AA = 'A'+str(n+1)+':A'+str(n+len(list_9[i][1])+5)
            if 'black' in list_9[i][0].lower():
                ws.merge_range(AA,list_9[i][0],style_9)
            elif 'white' in list_9[i][0].lower():
                ws.merge_range(AA,list_9[i][0],style_11)
            elif 'gold' in list_9[i][0].lower():
                ws.merge_range(AA,list_9[i][0],style_8)        
            elif 'pink' in list_9[i][0].lower():
                ws.merge_range(AA,list_9[i][0],style_10)
            else:
                ws.merge_range(AA,list_9[i][0],style_11)#填写物料名+wf/cell+color

            for j in range(len(list_9[i][1])):
                ws.write(n+j,2,list_9[i][1][j],style_2)#厂商名
                ws.write(n+j,1,'Daily Supply',style_2)
                for d in range(len(dates)):
                    if dates[d] in list_9[i][2][j]:
                        ws.write(n+j,3+d,list_9[i][2][j][dates[d]],style_2)#写入对应的数量
                    else:
                        ws.write(n+j,3+d,'',style_2)

            for d in range(len(dates)):
                if dates[d] in list_9[i][3]:
                    if int(list_9[i][3][dates[d]])!=0:
                        ws.write(n + len(list_9[i][1]) + 1, 3 + d, list_9[i][3][dates[d]], style_2)
                if dates[d] in list_9[i][4]:
                    if int(list_9[i][4][dates[d]])!=0:
                        ws.write(n + len(list_9[i][1]) + 2, 3 + d, list_9[i][4][dates[d]], style_2)
                x=list_A[3+d]
                coord=[] 
                for j in range(len(list_9[i][1])):
                    y=n+j+1
                    coord.append(x+str(y))
                c='+'.join(coord)
                ws.write_formula(n+len(list_9[i][1]),3+d,'='+c,style_2)#Sum Supply 
                coord=[]
                x=list_A[3+d]
                y1=n+len(list_9[i][1])+2
                y2=n+len(list_9[i][1])+3
                coord.append(x+str(y1))
                coord.append(x+str(y2))
                c='+'.join(coord)
                ws.write_formula(n+len(list_9[i][1])+3,3+d,'='+c,style_2) #Cum Transfer+NG
                x1=list_A[2+d]
                if '(wf)' in list_9[i][0].lower():
                    if 'black' in list_9[i][0].lower():
                        p=x+'10'
                    elif 'white' in list_9[i][0].lower():
                        p=x+'11'
                    elif 'gold' in list_9[i][0].lower():
                        p=x+'12'
                    elif 'pink' in list_9[i][0].lower():
                        p=x+'13'
                    else:
                        p=x+'10-'+x+'11-'+x+'12-'+x+'13'
                elif '(cell)' in list_9[i][0].lower():
                    if 'black' in list_9[i][0].lower():
                        p=x+'5'
                    elif 'white' in list_9[i][0].lower():
                        p=x+'6'
                    elif 'gold' in list_9[i][0].lower():
                        p=x+'7'
                    elif 'pink' in list_9[i][0].lower():
                        p=x+'8'
                    else:
                        p=x+'5-'+x+'6-'+x+'7-'+x+'8'
                else:
                    if 'black' in list_9[i][0].lower():
                        p=x+'5-'+x+'10'
                    elif 'white' in list_9[i][0].lower():
                        p=x+'6-'+x+'11'
                    elif 'gold' in list_9[i][0].lower():
                        p=x+'7-'+x+'12'
                    elif 'pink' in list_9[i][0].lower():
                        p=x+'8-'+x+'13'
                    else:
                        p=x+'4' 
                if d==0:
                    ws.write_formula(n+len(list_9[i][1])+4,3,'='+'0'+'-D'+str(n+len(list_9[i][1])+4)+'-'+p,style_5)
                else:
                    ws.write_formula(n+len(list_9[i][1])+4,3+d,'='+x1+str(n+len(list_9[i][1])+5)+'+'+x1+str(n+len(list_9[i][1])+1)+'-'+p+'-'+x+str(n+len(list_9[i][1])+4),style_5) 
                    #填写Delta的公式
                ws.conditional_format('D'+str(n+len(list_9[i][1])+5)+':'+list_A[len(dates)+3]+str(n+len(list_9[i][1])+5), {'type': 'cell',
                                                 'criteria': '<','value': 0,'format': format2}) 
            ws.write(n+len(list_9[i][1]),1,'Sum Supply',style_2)
            ws.write(n+len(list_9[i][1])+1,1,'Transfer',style_2)
            ws.write(n+len(list_9[i][1])+2,1,'NG',style_2)
            ws.write(n+len(list_9[i][1])+3,1,'Cum Transfer+NG',style_2)
            ws.write(n+len(list_9[i][1])+4,1,'Delta',style_5)
                       
            n=n+len(list_9[i][1])+5
        workbook.close()
        

# root = tk.Tk()
# wd().wn()
# root.mainloop()