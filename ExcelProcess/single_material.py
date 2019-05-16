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
import sys


def error_warn(error_msg):
    print(error_msg)

    sys.exit()
    return error_msg


def alphabet(number):
    _ascii = []
    _over_ascii = []
    for i in range(number):
        a = ord('A') + i
        _ascii.append(chr(a))
        if len(_ascii) > 26:
            if len(_ascii) % 26 == 0:
                b = len(_ascii) // 26
                c = ord('A') + b - 2
            else:
                b = len(_ascii) // 26
                c = ord('A') + b - 1
            if len(_ascii) % 26 == 0:
                e = ord('A') + 25
            else:
                d = len(_ascii) % 26
                e = ord('A') + d - 1
            f = chr(c) + chr(e)
            _over_ascii.append(f)
    alphabet_result = _ascii[0:26] + _over_ascii
    return alphabet_result


def component(excel_name, sheet_name='FATP'):
    raw_data = pd.read_excel(excel_name, sheet_name=sheet_name, header=None)
    raw_data.dropna(axis=1, how='all', inplace=True)
    raw_data.dropna(axis=0, how='all', inplace=True)
    raw_data.reset_index(drop=True, inplace=True)

    # 找出起始位置
    start_loc = raw_data[raw_data[0] == 'Component'].index.tolist()
    if len(start_loc) != 1:
        error_warn('第一列中没有Component或有多个Component')

    start_row = start_loc[0]
    raw_data.columns = raw_data.iloc[start_row]
    raw_data = raw_data.iloc[start_row + 1:, :]

    try:
        raw_data = raw_data[
            ['Component', 'OEM PN', 'Description', 'Part Detail', 'Vendor', 'Config', 'Req', 'ETA', 'Ship Qty']
        ]
    except KeyError as error:
        error_warn(error)

    raw_data = raw_data[~(pd.isnull(raw_data['ETA']))]
    raw_data = raw_data[~(pd.isnull(raw_data['Ship Qty']))]
    raw_data.fillna(method='pad', inplace=True)
    raw_data.reset_index(drop=True, inplace=True)

    # 补全合并单元格，转换日期格式
    try:
        _date = pd.to_datetime(raw_data.ETA, format="%m/%d/%Y")
        raw_data.ETA = _date.apply(lambda x: datetime.datetime.strftime(x, "%Y-%m-%d"))
    except ValueError as date_error:
        error_warn(date_error)

    # 所有component列表
    raw_data = raw_data[raw_data['Component'].str.len() >= 31 & ~(raw_data['Component'].str.contains('\n'))]
    user_select = list(set(raw_data['Component']))
    return user_select, raw_data


def write_sheet(read_sheet, raw_data, user_return, save_name):
    # 定义表名
    workbook = xlsxwriter.Workbook(save_name)
    for component_name in user_return:
        # 判定是否有key-in表
        if read_sheet == '':
            key_in = False
        else:
            key_in = True

        # 判断是否有相同sheet
        exist_same_sheet = False
        if key_in:
            key_in_excel = pd.ExcelFile(read_sheet)
            key_in_names = key_in_excel.sheet_names
            if component_name in key_in_names:
                exist_same_sheet = True

        data_process = raw_data[raw_data['Component'] == component_name]
        data_process = data_process.sort_values(by=['Component', 'Vendor', 'Config'])
        data_process.reset_index(drop=True, inplace=True)

        # 將重複的合併
        i = 1
        while i < len(data_process):
            if (data_process.loc[i, 'Component'] == data_process.loc[i-1, 'Component'] and
                    data_process.loc[i, 'Vendor'] == data_process.loc[i-1, 'Vendor'] and
                    data_process.loc[i, 'Config'] == data_process.loc[i-1, 'Config']and
                    data_process.loc[i, 'Req'] == data_process.loc[i-1, 'Req'] and
                    data_process.loc[i, 'ETA'] == data_process.loc[i-1, 'ETA']):
                data_process.loc[i-1, 'Ship Qty'] += data_process.loc[i, 'Ship Qty']
                data_process = data_process.drop(i)
                data_process.reset_index(drop=True, inplace=True)
            else:
                i += 1

        data = np.array(data_process)
        data_result = data.tolist()

        if exist_same_sheet:
            key_in_excel = pd.read_excel(read_sheet, sheet_name=component_name)
            key_in_excel = key_in_excel.drop(0)
            key_in_excel.columns = key_in_excel.loc[1]
            key_in_excel.reset_index(drop=True, inplace=True)
            key_in_excel.columns = range(len(key_in_excel.loc[0]))
            key_in_excel = key_in_excel.fillna('')

        # 写sheet
        sheet = workbook.add_worksheet(component_name)
        sheet_date = []
        for i in range(len(data_result)):
            sheet_date.append(data_result[i][7])
        if exist_same_sheet:
            for x in range(8, len(key_in_excel.loc[1])):
                sheet_date.append(key_in_excel.loc[0, x])

        if exist_same_sheet:
            key_in_input = []
            for x in range(2, len(key_in_excel), 4):
                for y in range(8, len(key_in_excel.loc[x])):
                    _input = []
                    if key_in_excel.loc[x, y] != '':
                        _input.append(key_in_excel.loc[x-1, 0])
                        _input.append(key_in_excel.loc[x-1, 1])
                        _input.append(key_in_excel.loc[x-1, 2])
                        _input.append(key_in_excel.loc[x-1, 3])
                        _input.append(key_in_excel.loc[x-1, 4])
                        _input.append(key_in_excel.loc[x-1, 5])
                        _input.append(key_in_excel.loc[x-1, 6])
                        _input.append(key_in_excel.loc[0, y])
                        _input.append(key_in_excel.loc[x, y])
                        key_in_input.append(_input)
            key_in_fail = []
            for x in range(3, len(key_in_excel), 4):
                for y in range(8, len(key_in_excel.loc[x])):
                    _fail = []
                    if key_in_excel.loc[x, y] != '':
                        _fail.append(key_in_excel.loc[x-2, 0])
                        _fail.append(key_in_excel.loc[x-2, 1])
                        _fail.append(key_in_excel.loc[x-2, 2])
                        _fail.append(key_in_excel.loc[x-2, 3])
                        _fail.append(key_in_excel.loc[x-2, 4])
                        _fail.append(key_in_excel.loc[x-2, 5])
                        _fail.append(key_in_excel.loc[x-2, 6])
                        _fail.append(key_in_excel.loc[0, y])
                        _fail.append(key_in_excel.loc[x, y])
                        key_in_fail.append(_fail)

        # 找出时间起始和结束范围
        sheet_col = []
        start_date = min(sheet_date)
        end_date = max(sheet_date)
        _start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d')
        _last_date = datetime.datetime.strptime(end_date, '%Y-%m-%d')
        _last_date = _last_date + datetime.timedelta(days=1)
        date_range = pd.date_range(_start_date, _last_date)
        date_range = list(date_range)
        for date in date_range:
            date = date.strftime('%Y-%m-%d')
            sheet_col.append(date)

        list_alphabet = alphabet(len(sheet_col) + 8)

        # 合并相同项,作為每個表中的行
        sheet_rows = []
        sheet_rows.append(data_result[0])
        for k in range(1, len(data_result)):
            if (data_result[k][0] != data_result[k-1][0] or data_result[k-1][1] != data_result[k-1][1] or
                    data_result[k][2] != data_result[k-1][2] or data_result[k][3] != data_result[k-1][3] or
                    data_result[k][4] != data_result[k-1][4] or data_result[k][5] != data_result[k-1][5]):
                sheet_rows.append(data_result[k])

        def transfer(space, step):
            space = int(space)
            space = space + step
            space = '%d' % space
            return space

        def conditional(start_row, start_col, end_row, end_col, criteria, decision_value, my_format, step):
            for _every_row in range(len(sheet_rows)):
                sheet.conditional_format(start_col + start_row + ':' + end_col + end_row,
                                         {'type': 'cell', 'criteria': criteria,
                                          'value': decision_value, 'format': my_format})
                start_row = transfer(start_row, step)
                end_row = transfer(end_row, step)

        # 樣式設置
        format_conditional_red = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'bg_color': '#EEA2AD'})
        format_conditional_green = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'bg_color': '#7CCD7C', 'align': 'center'})
        format_title = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 12, 'bold': True, 'align': 'center', 'valign': 'vcenter'})
        format_title_row = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 12, 'align': 'center',
             'bg_color': '#969696', 'bold': True, 'border': 1})
        format_week = workbook.add_format(
            {'bg_color': '#D3D3D3'})
        format_qty = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'bg_color': '#FFEC8B', 'border': 1, 'align': 'center'})
        format_vendor = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'bg_color': '#87CEEB', 'align': 'center'})
        format_config = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'bg_color': '#FFD39B', 'align': 'center'})
        format_center = workbook.add_format(
            {'font_name': 'Cambria', 'font_size': 11, 'align': 'center'})

        # 列寬設置
        sheet.set_column('A:A', 12)
        sheet.set_column('B:C', 22)
        sheet.set_column('D:E', 12)
        sheet.set_column('F:F', 22)
        sheet.set_column(6,len(sheet_col)+7, 12)

        sheet.set_row(1, 20)
        sheet.set_row(1, 20)
        sheet.set_row(2, 24)

        conditional('5', 'H', '5', list_alphabet[len(sheet_col)+7], '>', 0, format_conditional_green, 4)
        conditional('7', 'H', '7', list_alphabet[len(sheet_col)+7], '<', 0, format_conditional_red, 4)

        sheet.merge_range(0, 0, 1, 1, 'SUM OF TOTAL INPUT' + '\n' + component_name, format_title)
        sheet.write(2, 0, 'Component', format_title_row)
        sheet.write(2, 1, 'OEM PN', format_title_row)
        sheet.write(2, 2, 'Description', format_title_row)
        sheet.write(2, 3, 'Part Detail', format_title_row)
        sheet.write(2, 4, 'Vendor', format_title_row)
        sheet.write(2, 5, 'Config', format_title_row)
        sheet.write(2, 6, 'Req Qty', format_title_row)
        sheet.write(2, 7, 'Cum', format_title_row)

        # 写row
        y = 0
        for x in range(3, 4*len(sheet_rows)+3, 4):
            sheet.write(x, 0, sheet_rows[y][0], format_title)
            sheet.write(x, 1, sheet_rows[y][1], format_center)
            sheet.write(x, 2, sheet_rows[y][2], format_center)
            sheet.write(x, 3, sheet_rows[y][3], format_center)
            sheet.write(x, 4, sheet_rows[y][4], format_vendor)
            sheet.write(x, 5, sheet_rows[y][5], format_config)
            sheet.write(x, 6, sheet_rows[y][6], format_center)
            sheet.write(x+1, 4, 'Input', format_center)
            sheet.write(x+2, 4, 'Fail/Transfer', format_center)
            sheet.write(x+3, 4, 'Delta', format_center)
            y = y + 1

        # 写日期
        week_day = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        y = 0
        for x in range(8, len(sheet_col)+8):
            sheet.write(2, x, sheet_col[y], format_title_row)
            date = datetime.datetime.strptime(sheet_col[y], '%Y-%m-%d')
            day = date.weekday()
            sheet.write(1, x, week_day[day], format_title)
            if day == 6:
                x1 = '%d' % (len(sheet_rows)*4+3)
                sheet.conditional_format(list_alphabet[y + 8] + '2' + ':' + list_alphabet[y + 8] + x1,
                                         {'type': 'no_blanks', 'format': format_week})
                sheet.conditional_format(list_alphabet[y+8]+'2'+':'+list_alphabet[y+8] + x1,
                                         {'type': 'blanks', 'format': format_week})
            y += 1

        # 写公式
        x1, x2 = '5', '6'
        for x in range(6, 4*len(sheet_rows)+6, 4):
            y1 = list_alphabet[8]
            sheet.write(x, 8, '=' + '-' + y1 + x1 + '-' + y1 + x2)
            x1 = transfer(x1, 4)
            x2 = transfer(x2, 4)

        x1, x2, x3, x4 = '4', '5', '6', '7'
        for x in range(6, 4*len(sheet_rows)+6, 4):
            for y in range(9, len(sheet_col)+8):
                y1 = list_alphabet[y-1]
                y2 = list_alphabet[y]
                sheet.write(x, y, '=' + y1 + x1 + '+' + y1 + x4 + '-' + y2 + x2 + '-' + y2 + x3)
            x1 = transfer(x1, 4)
            x2 = transfer(x2, 4)
            x3 = transfer(x3, 4)
            x4 = transfer(x4, 4)

        # 将对应值填入表中
        row = []
        col = []
        value = []
        for i in range(len(data_result)):
            for q in range(len(sheet_rows)):
                if (data_result[i][0] == sheet_rows[q][0] and data_result[i][1] == sheet_rows[q][1] and
                        data_result[i][2] == sheet_rows[q][2] and data_result[i][3] == sheet_rows[q][3]):
                    row.append(q*4+3)
        for i in range(len(data_result)):
            for q in range(len(sheet_col)):
                if data_result[i][7] == sheet_col[q]:
                    col.append(q+8)
        for i in range(len(data_result)):
            a = data_result[i][8]
            value.append(a)
        for i in range(len(data_result)):
            sheet.write(row[i], col[i], value[i], format_qty)

        # Cum
        z = '4'
        for x in range(len(sheet_rows)):
            y = len(sheet_col) + 7
            y1 = list_alphabet[y]
            A = 'I'+z
            B = y1+z
            sheet.write(4*x+3, 7, '=' + 'SUM(' + A + ':' + B + ')', format_center)
            z = transfer(z, 4)

        # 添加key in值
        if exist_same_sheet:
            row = []
            col = []
            value = []
            for i in range(len(key_in_input)):
                for q in range(len(sheet_rows)):
                    if (key_in_input[i][0] == sheet_rows[q][0] and key_in_input[i][1] == sheet_rows[q][1] and
                            key_in_input[i][2] == sheet_rows[q][2] and key_in_input[i][3] == sheet_rows[q][3]):
                        row.append(q*4+4)
            for i in range(len(key_in_input)):
                for q in range(len(sheet_col)):
                    if key_in_input[i][7] == sheet_col[q]:
                        col.append(q+8)
            for i in range(len(key_in_input)):
                a = key_in_input[i][8]
                value.append(a)
            for i in range(len(key_in_input)):
                sheet.write(row[i], col[i], value[i])

            row = []
            col = []
            value = []
            for i in range(len(key_in_fail)):
                for q in range(len(sheet_rows)):
                    if (key_in_fail[i][0] == sheet_rows[q][0] and key_in_fail[i][1] == sheet_rows[q][1] and
                            key_in_fail[i][2] == sheet_rows[q][2] and key_in_fail[i][3] == sheet_rows[q][3]):
                        row.append(q*4+5)
            for i in range(len(key_in_fail)):
                for q in range(len(sheet_col)):
                    if key_in_fail[i][7] == sheet_col[q]:
                        col.append(q+8)
            for i in range(len(key_in_fail)):
                a = key_in_fail[i][8]
                value.append(a)
            for i in range(len(key_in_fail)):
                sheet.write(row[i], col[i], value[i])

    workbook.close()


if __name__ == '__main__':
    list_component, d = component('DRP.xlsx')
    component_return = ['ANT1']
    write_sheet('', d, component_return, 'result.xlsx')