#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 29 09:40:24 2018

@author: zt
"""

import pandas as pd
import numpy as np
import datetime
import xlsxwriter
from pandas.core.frame import DataFrame

def component(excel_name):
    d = pd.read_excel(excel_name, sheet_name='FATP')
    d=d.dropna(axis=1,how='all')
    d=d.dropna(axis=0,how='all')
    d.index=range(len(d))
    d.columns=range(len(d.loc[0]))

    for i in range(len(d)):
        if d.loc[i,0]!='Component':
            d= d.drop(i)
        else:
            break

    #原表中取出所需數據
    d.index=range(len(d))
    d.columns=d.loc[0]
    d=d.drop(0)
    d.index=range(len(d))
    d=d[['Component','OEM PN','Description','Part Detail','Vendor','Config','Req','ETA','Ship Qty']]
    d[['Component','OEM PN','Vendor','Config','Req','ETA','Ship Qty','Description','Part Detail']]=d[['Component','OEM PN','Vendor','Config','Req','ETA','Ship Qty','Description','Part Detail']].fillna('')
    d=d[d['ETA']!='']
    d.index=range(len(d))

    #所有component列表
    list_component = []
    for i in range(len(d)):
        if (d.loc[i,'Component'])!='' and len(d.loc[i,'Component']) <=31 and '\n' not in d.loc[i,'Component'] :
            list_component.append(d.loc[i,'Component'])
    list_component = list(set(list_component))
    return list_component,d


def write_sheet(read_sheet,d,component_return,save_name):
    list_component_return = component_return
    if read_sheet=='':
        url2 = False
    else:
        url2 = True
    

    #寫表
    workbook = xlsxwriter.Workbook(save_name+'.xlsx')

    data = np.array(d)
    list_dataframe = data.tolist()
    for i in range(len(list_dataframe)):
        if list_dataframe[i][0] == '':
            list_dataframe[i][0] = list_dataframe[i-1][0]
            list_dataframe[i][1] = list_dataframe[i-1][1]
            list_dataframe[i][2] = list_dataframe[i-1][2]
            list_dataframe[i][3] = list_dataframe[i-1][3]
            list_dataframe[i][4] = list_dataframe[i-1][4]
            list_dataframe[i][5] = list_dataframe[i-1][5]
            list_dataframe[i][6] = list_dataframe[i-1][6]
            
        a=datetime.datetime.strptime(list_dataframe[i][7],'%m/%d/%Y')
        list_dataframe[i][7]=a.strftime('%Y/%m/%d')
        
    d = DataFrame(list_dataframe)
    d.columns = ['Component','OEM PN','Description','Part Detail','Vendor','Config','Req','ETA','Ship Qty'] 
                
    for component in list_component_return:

        dd = d[d['Component']==component]    
        dd.index = range(len(dd))
                
        dd = dd.sort_values(by=['Component','Vendor','Config'])
        dd.index = range(len(dd)) 
        #將重複的合併
        i=1
        while i<len(dd) :
            if dd.loc[i,'Component'] == dd.loc[i-1,'Component'] and dd.loc[i,'Vendor'] == dd.loc[i-1,'Vendor'] and dd.loc[i,'Config'] == dd.loc[i-1,'Config'] and dd.loc[i,'Req'] == dd.loc[i-1,'Req'] and dd.loc[i,'ETA'] == dd.loc[i-1,'ETA']:
                 dd.loc[i-1,'Ship Qty'] +=  dd.loc[i,'Ship Qty']
                 dd = dd.drop(i)
                 dd.index = range(len(dd))
            else:
                i += 1

        #创建行索引列表
        list_alphabet = []
        for j in range(26):
            a = chr(65+j)
            list_alphabet.append(a)
        for j in range(26):
            a = chr(65+j)
            list_alphabet.append('A'+a)
        for j in range(26):
            a = chr(65+j)
            list_alphabet.append('B'+a)
        for j in range(26):
            a = chr(65+j)
            list_alphabet.append('C'+a)
        for j in range(26):
            a = chr(65+j)
            list_alphabet.append('D'+a)
        
        data = np.array(dd)
        list_delta = data.tolist()
        
        condition = False
        if url2 :
            db = pd.read_excel(read_sheet,sheet_name=None)
            keys = db.keys()
            if component in keys:
                condition = True

        if condition:
            df = pd.read_excel(read_sheet, sheet_name=component)
            df = df.drop(0)
            df.columns = df.loc[1]
            df.index = range(len(df))
            df.columns = range(len(df.loc[0]))
            df = df.fillna('')
            
        sheet = workbook.add_worksheet(component)
            
        list_sheet_date = []
        for i in range(len(list_delta)):
            list_sheet_date.append(list_delta[i][7])
        if condition :
            for x in range(8,len(df.loc[1])):
                list_sheet_date.append(df.loc[0,x])

        if condition :
            list_input = []
            for x in range(2,len(df),4):
                for y in range(8,len(df.loc[x])):
                    list_input1 = []
                    if df.loc[x,y] != '' :
                        list_input1.append(df.loc[x-1,0])
                        list_input1.append(df.loc[x-1,1])
                        list_input1.append(df.loc[x-1,2])
                        list_input1.append(df.loc[x-1,3])
                        list_input1.append(df.loc[x-1,4])
                        list_input1.append(df.loc[x-1,5])
                        list_input1.append(df.loc[x-1,6])
                        list_input1.append(df.loc[0,y])
                        list_input1.append(df.loc[x,y])
                        list_input.append(list_input1)
            list_fail = []
            for x in range(3,len(df),4):
                for y in range(8,len(df.loc[x])):
                    list_fail1 = []
                    if df.loc[x,y] != '' :
                        list_fail1.append(df.loc[x-2,0])
                        list_fail1.append(df.loc[x-2,1])
                        list_fail1.append(df.loc[x-2,2])
                        list_fail1.append(df.loc[x-2,3])
                        list_fail1.append(df.loc[x-2,4])
                        list_fail1.append(df.loc[x-2,5])
                        list_fail1.append(df.loc[x-2,6])
                        list_fail1.append(df.loc[0,y])
                        list_fail1.append(df.loc[x,y])
                        list_fail.append(list_fail1)

        list_sheet_col = []
        firstday = min(list_sheet_date)
        lastday = max(list_sheet_date)
        firstday_3 = datetime.datetime.strptime(firstday,'%Y/%m/%d')
        lastday_3 = datetime.datetime.strptime(lastday,'%Y/%m/%d')
        lastday_3=lastday_3+datetime.timedelta(days=1)
        datelist = pd.date_range(firstday_3,lastday_3)
        datelist = list(datelist)
        for j in range(len(datelist)):
            datelist[j] = datelist[j].strftime('%Y/%m/%d')
            list_sheet_col.append(datelist[j])

        #合并相同项,作為每個表中的行
        list_sheet_row = []
        list_sheet_row.append(list_delta[0])
        for k in range(1,len(list_delta)):
            if (list_delta[k][0] != list_delta[k-1][0]) or (list_delta[k][1] != list_delta[k-1][1]) or (list_delta[k][2] != list_delta[k-1][2]) or (list_delta[k][3] != list_delta[k-1][3]):
                list_sheet_row.append(list_delta[k])


        def transfer(x,step):
            x = int(x)
            x = x + step
            x = '%d' %x
            return x
        
        def conditional(x1,y1,x2,y2,criteria,value,my_format,step):
            for x in range(len(list_sheet_row)):
                sheet.conditional_format(y1+x1+':'+y2+x2, {'type':'cell',
                                                        'criteria': criteria,
                                                        'value':value,
                                                        'format':my_format})
                x1 = transfer(x1,step)
                x2 = transfer(x2,step)

        #樣式設置
        format_conditional_red = workbook.add_format({'bg_color':'#EEA2AD','border':1})
        format_conditional_green = workbook.add_format({'bg_color':'#7CCD7C','border':1,'align':'center'})
        format_title = workbook.add_format({'bold':True,'align':'center'})
        format_title_row = workbook.add_format({'font_name': 'Calibri','font_size': 12,'valign':'center','bg_color':'#969696','bold':True,'border':1})
        format_week = workbook.add_format({'bg_color':'#D3D3D3','border':1})
        format_qty = workbook.add_format({'bg_color':'#FFEC8B','border':1,'align':'center'})
        format_vendor = workbook.add_format({'bg_color':'#87CEEB','align':'center'})
        format_config = workbook.add_format({'bg_color':'#FFD39B','align':'center'})
        format_center = workbook.add_format({'align':'center'})
                            
        #列寬設置
        sheet.set_column('A:A',12)
        sheet.set_column('B:C',22)
        sheet.set_column('D:E',12)
        sheet.set_column('F:F',22)
        sheet.set_column(6,len(list_sheet_col)+7,12)

        sheet.set_row(1,20)
        sheet.set_row(1,20)
        sheet.set_row(2,24)
        
        conditional('5','H','5',list_alphabet[len(list_sheet_col)+6],'>',0,format_conditional_green,4)
        conditional('7','H','7',list_alphabet[len(list_sheet_col)+6],'<',0,format_conditional_red,4)

        sheet.merge_range(0,0,1,1,'SUM OF TOTAL INPUT'+ '\n' + component,format_title)
        sheet.write(2,0,'Component',format_title_row)
        sheet.write(2,1,'OEM PN',format_title_row)
        sheet.write(2,2,'Description',format_title_row)
        sheet.write(2,3,'Part Detail',format_title_row)
        sheet.write(2,4,'Vendor',format_title_row)
        sheet.write(2,5,'Config',format_title_row)
        sheet.write(2,6,'Req Qty',format_title_row)
        sheet.write(2,7,'Cum',format_title_row)

        #写row
        y = 0
        for x in range(3,4*len(list_sheet_row)+3,4):
            sheet.write(x,0,list_sheet_row[y][0],format_title)
            sheet.write(x,1,list_sheet_row[y][1],format_center)
            sheet.write(x,2,list_sheet_row[y][2],format_center)
            sheet.write(x,3,list_sheet_row[y][3],format_center)
            sheet.write(x,4,list_sheet_row[y][4],format_vendor)
            sheet.write(x,5,list_sheet_row[y][5],format_config)
            sheet.write(x,6,list_sheet_row[y][6],format_center)
            sheet.write(x+1,4,'Input',format_center)
            sheet.write(x+2,4,'Fail/Transfer',format_center)
            sheet.write(x+3,4,'Delta',format_center)
            y = y + 1

    #    #写日期
        list_week = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
        y = 0
        for x in range(8,len(list_sheet_col)+8):
            sheet.write(2,x,list_sheet_col[y],format_title_row)

            date = datetime.datetime.strptime(list_sheet_col[y],'%Y/%m/%d')
            day = date.weekday()
            sheet.write(1,x,list_week[day],format_title)
            if day == 6:
                x1 = '%d' %(len(list_sheet_row)*4+3)
                sheet.conditional_format(list_alphabet[y+8]+'2'+':'+list_alphabet[y+8]+x1, {'type':'no_blanks',
                                                        'criteria': '!=',
                                                        'value':0,
                                                        'format':format_week})
                sheet.conditional_format(list_alphabet[y+8]+'2'+':'+list_alphabet[y+8]+x1, {'type':'blanks',
                                                        'criteria': '!=',
                                                        'value':0,
                                                        'format':format_week})

            y += 1

    #    写公式
        x1,x2='5','6'
        for x in range(6,4*len(list_sheet_row)+6,4):
            y1 = list_alphabet[8]
            sheet.write(x,8,'='+'-'+y1+x1+'-'+y1+x2)
            x1 = transfer(x1,4)
            x2 = transfer(x2,4)

        x1,x2,x3,x4 ='4','5','6','7'
        for x in range(6,4*len(list_sheet_row)+6,4):
            for y in range(9,len(list_sheet_col)+8):
                y1 = list_alphabet[y-1]
                y2 = list_alphabet[y]
                sheet.write(x,y,'='+y1+x1+'+'+y1+x4+'-'+y2+x2+'-'+y2+x3)
            x1 = transfer(x1,4)
            x2 = transfer(x2,4)
            x3 = transfer(x3,4)
            x4 = transfer(x4,4)

        #将对应值填入表中
        row = []
        col = []
        value = []
        for i in range(len(list_delta)):
            for q in range(len(list_sheet_row)):
                if (list_delta[i][0] == list_sheet_row[q][0]) and (list_delta[i][1] == list_sheet_row[q][1]) and (list_delta[i][2] == list_sheet_row[q][2]) and (list_delta[i][3] == list_sheet_row[q][3]):
                    row.append(q*4+3)
        for i in range(len(list_delta)):
            for q in range(len(list_sheet_col)):
                if (list_delta[i][7] == list_sheet_col[q]):
                    col.append(q+8)
        for i in range(len(list_delta)):
            a = list_delta[i][8]
            value.append(a)

        for i in range(len(list_delta)):
            sheet.write(row[i],col[i],value[i],format_qty)

        #Cum
        z = '4'
        for x in range(len(list_sheet_row)):
            y = len(list_sheet_col) + 7
            y1 = list_alphabet[y]
            A = 'I'+z
            B = y1+z
            sheet.write(4*x+3,7,'='+'SUM('+A+':'+B+')',format_center)
            z = transfer(z,4)

    #添加key in值
        if condition:
            row = []
            col = []
            value = []
            for i in range(len(list_input)):
                for q in range(len(list_sheet_row)):
                    if (list_input[i][0] == list_sheet_row[q][0]) and (list_input[i][1] == list_sheet_row[q][1]) and (list_input[i][2] == list_sheet_row[q][2]) and (list_input[i][3] == list_sheet_row[q][3]):
                        row.append(q*4+4)
            for i in range(len(list_input)):
                for q in range(len(list_sheet_col)):
                    if (list_input[i][7] == list_sheet_col[q]):
                        col.append(q+8)
            for i in range(len(list_input)):
                a = list_input[i][8]
                value.append(a)

            for i in range(len(list_input)):
                sheet.write(row[i],col[i],value[i])

            row = []
            col = []
            value = []
            for i in range(len(list_fail)):
                for q in range(len(list_sheet_row)):
                    if (list_fail[i][0] == list_sheet_row[q][0]) and (list_fail[i][1] == list_sheet_row[q][1]) and (list_fail[i][2] == list_sheet_row[q][2]) and (list_fail[i][3] == list_sheet_row[q][3]):
                        row.append(q*4+5)
            for i in range(len(list_fail)):
                for q in range(len(list_sheet_col)):
                    if (list_fail[i][7] == list_sheet_col[q]):
                        col.append(q+8)
            for i in range(len(list_fail)):
                a = list_fail[i][8]
                value.append(a)

            for i in range(len(list_fail)):
                sheet.write(row[i],col[i],value[i])

    workbook.close()