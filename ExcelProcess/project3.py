#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 11 15:23:16 2018

@author: f7689594
"""

import pandas as pd
import numpy as np
import datetime
import xlsxwriter

class BCT:
    def __init__(self):
        self.list_color=[]
        self.list_1=[]
        self.list_2=[]
        self.list_read=[]
        self.list_search_1=[]
        self.df_d=pd.DataFrame()
        self.str_1=''
    def read_G(self,url):
        d = pd.read_excel(url, sheet_name='CTB')
        d=d.dropna(axis=1,how='all')
        d=d.dropna(axis=0,how='all')
        d=d.fillna('')
        d.columns=d.loc[0]
        d.index= range(len(d))
        d_1=d.loc[0:11]
        for i in range(2,len(d_1)):
            dic_color={}
            list_color=[]
            list_color.append(d_1.at[i,d_1.columns[1]])
            for date in range(3,len(d_1.columns)):
                if d_1.at[i,d_1.columns[date]]!='':
                    dic_color[d_1.columns[date]]=d_1.at[i,d_1.columns[date]]
            list_color.append(dic_color)
            self.list_color.append(list_color)
            
        d_2=d.loc[12:]
        d_2.index=range(len(d_2))
        n=0
        self.list_1 = []
        self.list_read = []
        while n<len(d_2):
            list_2=[]
            dic_2={}
            dic_3={}
            list_2.append(d_2.at[n,'CTB'])
            for j in range(n,len(d_2)):
                if d_2.at[j,'Vendor']!='':
                    n+=1
                else:
                    break
            for date in range(3,len(d_2.columns)):
                if d_2.at[n+1,d_2.columns[date]]!='':
                    dic_2[d_2.columns[date]]=d_2.at[n+1,d_2.columns[date]]
                if d_2.at[n+2,d_2.columns[date]]!='':
                    dic_3[d_2.columns[date]]=d_2.at[n+2,d_2.columns[date]]
            list_2.append([dic_2,dic_3])
            self.list_1.append(list_2)
            n+=5
        for i in self.list_1:
            self.list_read.append(i[0])
        
    def read_x(self,url,list_search_A=[]):
        self.list_search_1=list_search_A
        list_search=self.list_read+self.list_search_1
        list_search = list(set(list_search))
        d = pd.read_excel(url, sheet_name='FATP')
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
        d[['Component', 'Color', 'Config', 'Vendor', 'ETA', 'Alloc', 'NG', 'Fail']] = d[['Component', 'Color', 'Config', 'Vendor', 'ETA', 'Alloc', 'NG', 'Fail']].fillna('')  #将nan填充为‘’
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
        #拼接字段并保存到df中，零件名，厂商，日期列表，数量列表
        df.index = range(len(df))
        self.df_df=df

        df = df.T
        df.columns = df.loc['A']
        df = df[list_search]
        df = df.T
        df.index = range(len(df))
        
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
            list_5.append([A,list_8,list_3_1,list_4_1,list_Alloc_1_1,list_NG_1_1])#将相同零件名下对相同厂商合并在一行并保存到列表list_5中
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
        self.list_2=list_9
        self.df_d=d

    def wexcel(self,url1,url2,sheet_name,list_search):
        self.read_G(url2)
        self.read_x(url1,list_search)
        list_date=[]
        d=self.df_d
        for date in d['ETA']:
            list_date.append(datetime.datetime.strptime(date,'%m/%d/%Y'))
        for i in range(len(self.list_1)):
            for date in list(self.list_1[i][1][0].keys()):
                list_date.append(datetime.datetime.strptime(date,'%Y-%m-%d'))
            for date in list(self.list_1[i][1][1].keys()):
                list_date.append(datetime.datetime.strptime(date,'%Y-%m-%d'))
        for i in range(len(self.list_color)):
            for date in list(self.list_color[i][1].keys()):
                list_date.append(datetime.datetime.strptime(date,'%Y-%m-%d'))
        date_max=max(list_date)
        date_min=min(list_date)#找到起始日期和最后一天到日期

        dates = pd.date_range(date_min, date_max)#生成日期数列
        dates=dates.strftime('%Y-%m-%d')#将数列中到日期格式转化为str，
        list_9=[]

        for l_2 in self.list_2:
            list_add=[]
            list_add.append(l_2[0])
            list_add.append(l_2[1])
            list_add.append(l_2[2])
            list_add.append(l_2[3])
            list_add.append(l_2[4])
            for l_1 in self.list_1:
                if l_1[0]==l_2[0]:
                    list_add.append(l_1[1])
            try:
                list_add[5]
            except:
                list_add.append([{},{}])
            list_9.append(list_add)

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
        workbook = xlsxwriter.Workbook(sheet_name+'.xlsx')#开始写表
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
        str_1=self.str_1
        ws.write(0,0,str_1)
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
            if dates[d] in self.list_color[0][1]:
                ws.write(4,3+d,self.list_color[0][1][dates[d]],style_4)
            else:
                ws.write(4,3+d,'',style_4)

            if dates[d] in self.list_color[1][1]:
                ws.write(5,3+d,self.list_color[1][1][dates[d]],style_2)
            else:
                ws.write(5,3+d,'',style_2)
                
            if dates[d] in self.list_color[2][1]:
                ws.write(6,3+d,self.list_color[2][1][dates[d]],style_3)
            else:
                ws.write(6,3+d,'',style_3)
                
            if dates[d] in self.list_color[3][1]:
                ws.write(7,3+d,self.list_color[3][1][dates[d]],style_7)
            else:
                ws.write(7,3+d,'',style_7)
                
            if dates[d] in self.list_color[5][1]:
                ws.write(9,3+d,self.list_color[5][1][dates[d]],style_4)
            else:
                ws.write(9,3+d,'',style_4)
                
            if dates[d] in self.list_color[6][1]:
                ws.write(10,3+d,self.list_color[6][1][dates[d]],style_2)
            else:
                ws.write(10,3+d,'',style_2)
            if dates[d] in self.list_color[7][1]:
                ws.write(11,3+d,self.list_color[7][1][dates[d]],style_3)
            else:
                ws.write(11,3+d,'',style_3)
                
            if dates[d] in self.list_color[8][1]:
                ws.write(12,3+d,self.list_color[8][1][dates[d]],style_7)
            else:
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
                if dates[d] in list_9[i][5][0]:
                    ws.write(n+len(list_9[i][1])+1,3+d,list_9[i][5][0][dates[d]],style_2)
                else:
                    ws.write(n+len(list_9[i][1])+1,3+d,'',style_2)
                if dates[d] in list_9[i][5][1]:
                    ws.write(n+len(list_9[i][1])+2,3+d,list_9[i][5][1][dates[d]],style_2)
                else:
                    ws.write(n+len(list_9[i][1])+2,3+d,'',style_2)
                        
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
        
        
        return self.list_color,self.list_1,self.list_2,list_9,self.df_df
    def ddlist(self,url1,url2):
        self.read_G(url2)
        self.read_x(url1)
        dic_dd={}
        for dd in range(len(self.df_df)):
            dic_dd[self.df_df.at[dd,'A']]=0
        for dd in self.list_1:
            dic_dd[dd[0]]=0
        return list(dic_dd.keys())
            
        


#list_color,list_1,list_2,list_9,df=BCT().wexcel()
#list_dd=BCT().ddlist()







































