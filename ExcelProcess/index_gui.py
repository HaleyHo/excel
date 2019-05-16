from __future__ import absolute_import
from Data_analysis import *
from read_excel_edited import make_sheet
import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from TkinterDnD2 import TkinterDnD

from threading import Thread


class Gui:
    # 初始化Gui布局
    def __init__(self, master):
        master.title("IQC日報自動生成軟體--beta1.0")
        self.master = master
        self.cd = []  # 模拟的条件
        self.selced_cd = []

        # 进度条设置
        self.style = ttk.Style()
        self.style.theme_use("alt")
        self.style.configure("RGB(0,102,208).Horizontal.TProgressbar")

        self.data_state = False

        # 设置主窗体大小
        master.geometry("750x500+600+300")
        master.resizable(False, False)

        # 文件操作面板
        self.frame_content = tk.LabelFrame(master, text="文件操作:", height=40, relief=tk.RIDGE)
        self.frame_content.pack(fill=tk.X, padx=10)

        # 设置打开路径和保存路径的类型
        self.openpath = tk.StringVar()
        self.savepath = tk.StringVar()
        self.xfile = tk.StringVar()
        self.result_list = []

        # 打开文件框（可实现拖拽效果）
        default_value = '選擇文件路徑，拖入文件即可'
        tk.entry = self.build_label_value_block(default_value)
        # 给 widget 添加拖放方法
        self.add_drop_handle(tk.entry, self.handle_drop_on_local_path_entry)

        # 保存文件框
        self.save_name = tk.Entry(self.frame_content, width=55, font=('Verdana', 15), state="normal", textvariable=self.savepath)
        self.save_name.grid(row=3, column=0)

        # 打开文件按钮
        ttk.Button(self.frame_content, text="打開文件", width=15, command=self.openfile).grid(row=1, column=1, padx=5)

        # 数据分析按钮
        self.data_analyze_btn = ttk.Button(self.frame_content, text="分析數據", width=15,
                                           command=lambda: self.thread_it(self.analyze_data))
        self.data_analyze_btn.grid(row=2, column=1, padx=5)

        # 保存文件按钮
        ttk.Button(self.frame_content, text="保存路徑", width=15, command=self.save_file).grid(row=3, column=1, padx=5)

        # 这里添加进度条：
        self.pbr = ttk.Progressbar(self.frame_content, length=560, maximum=100, mode='indeterminate',
                                   orient=tk.HORIZONTAL)
        self.pbr.grid(row=2, column=0)
        self.pbr.pack_forget()

        # 条件面板
        self.apply_frame = tk.LabelFrame(master, text='條件查詢:', height=30, relief=tk.RIDGE)
        self.apply_frame.pack(fill=tk.X, padx=10, pady=10)

        self.lab = tk.Label(self.apply_frame, width=23)
        self.lab.grid(row=0, column=0, sticky=tk.E, pady=1, padx=5)

        self.build_report = ttk.Button(self.apply_frame, text='生成日報', width=15, command=self.get_results)
        self.build_report.grid(row=0, column=2, sticky=tk.E, pady=1)

        # 创建一个开始日期下拉列表
        self.start_date = tk.StringVar()
        self.stChosen = ttk.Combobox(self.apply_frame, width=29, textvariable=self.start_date)
        self.stChosen.grid(row=0, column=0, padx=5)  # 设置其在界面中出现的位置  column代表列   row 代表行

        # 创建一个结束日期下拉列表
        self.end_date = tk.StringVar()
        self.edChosen = ttk.Combobox(self.apply_frame, width=28, textvariable=self.end_date)
        self.edChosen.grid(row=0, column=1, padx=5)  # 设置其在界面中出现的位置  column代表列   row 代表行

        # 提示面板
        self.remind_frame = tk.LabelFrame(master, text='溫馨提示:', height=30, relief=tk.RIDGE)
        self.remind_frame.pack(fill=tk.X, padx=10, pady=10)

        self.lb = ttk.Label(self.remind_frame, text='數據越多處理時間越長，請耐心等待！！')
        self.lb.pack(fill=tk.X, padx=10, pady=10)

        self.lb1 = ttk.Label(self.remind_frame, text='')
        self.lb1.pack(fill=tk.X, padx=10, pady=10)

    # 打开文件函数
    def openfile(self):
        # 获取到文件路径
        name = filedialog.askopenfilename(parent=self.frame_content, filetypes=[('xlsx file', '.xlsx')])
        oldopenfilepath = self.openpath.get()
        if name != oldopenfilepath:
            self.cd.clear()
            self.selced_cd.clear()
            self.database = pd.DataFrame()
            self.openpath.set(name)
            self.data_analyze_btn.state(statespec=('!disabled',))

    # 文件拖拽lable的值
    def build_label_value_block(self, default_value, entry_size=55):
        entry = tk.Entry(self.frame_content, width=entry_size, font=('Verdana', 15), textvariable=self.openpath)
        entry.grid(row=1, column=0, padx=3, pady=1)
        entry.insert(0, default_value)
        return entry

    # 文件推拽处理
    def add_drop_handle(self, widget, handle):
        widget.drop_target_register('DND_Files')

        def drop_enter(event):
            event.widget.focus_force()
            return event.action

        def drop_position(event):
            return event.action

        def drop_leave(event):
            # leaving 应该清除掉之前 drop_enter 的 focus 状态, 怎么清?
            return event.action

        widget.dnd_bind('<<DropEnter>>', drop_enter)
        widget.dnd_bind('<<DropPosition>>', drop_position)
        widget.dnd_bind('<<DropLeave>>', drop_leave)
        widget.dnd_bind('<<Drop>>', handle)

    def handle_drop_on_local_path_entry(self, event):
        if event.data:
            files = event.widget.tk.splitlist(event.data)
            for f in files:
                if os.path.exists(f):
                    event.widget.delete(0, 'end')
                    event.widget.insert('end', f)
                else:
                    print('Not dropping file "%s": file does not exist.' % f)
        return event.action

    def thread_it(self, func):
        # 将函数打包进线程
        # 创建
        t = Thread(target=func)
        # 守护 !!!
        # t.setDaemon(True)
        # 启动
        t.start()
        # 阻塞--卡死界面！
        # t.join()

    # 分析数据
    def analyze_data(self):
        # 判断是否打开文件
        if len(self.openpath.get()) == 0:
            messagebox.showerror(title="錯誤", message="請輸入文件路徑", parent=self.frame_content)
        # elif self.openpath.get().lower().find('report') < 0:
        #     self.lb["text"] = '請選擇正確的文件進行數據處理！'
            return

        # 获取时间
        t1 = time.time()

        # 获取打开文件的路径
        path = self.openpath.get()
        self.data_analyze_btn.state(statespec=('!disabled',))
        self.pbr.start(interval=10)
        try:
            # 判断是什么部件的表格
            if path.find('Top Module') >= 0:
                self.lb["text"] = '數據正在分析中，請耐心等待'
                self.lb1["text"] = ''
                self.xfile = 'Top_module.xlsx'
                # try:
                self.result_list, date_list, pn_list, vendor_list, config_list = data_analysis_top(path)
                print('111', date_list)
                # 去除重复日期
                for date in np.unique(date_list):
                    self.cd.append(date)

                print(self.cd)

            elif path.find('HSG') >= 0:
                self.lb["text"] = '數據正在分析中，請耐心等待'
                self.lb1["text"] = ''
                self.xfile = 'other_module.xlsx'

                self.result_list, date_list, pn_list, vendor_list, config_list = data_analysis_hsg(path)
                # 去除重复日期
                for date in np.unique(date_list):
                    self.cd.append(date)

            else:
                self.lb["text"] = '數據正在分析中，請耐心等待'
                self.lb1["text"] = ''
                self.xfile = 'other_module.xlsx'

                self.result_list, date_list, pn_list, vendor_list, config_list = data_analysis_eeparts(path)
                # 去除重复日期
                for date in np.unique(date_list):
                    self.cd.append(date)

            # 运行时间
            self.speed1 = time.time() - t1

            self.cd = sorted(self.cd)
            # 设置日期下拉列表的值
            self.stChosen["values"] = self.cd
            self.edChosen["values"] = self.cd

            self.data_state = True
            self.pbr.stop()

            # 设置下拉列表默认值
            self.stChosen.current(0)
            self.edChosen.current(len(self.cd)-1)
            self.lb["text"] = '數據處理完成，條件選擇生成報表'

        except:
            messagebox.showerror(title="錯誤", message="請選擇正確的文件進行數據分析！", parent=self.frame_content)
            self.pbr.stop()

    # 文件保存路徑
    def save_file(self):
        name = filedialog.asksaveasfilename(parent=self.frame_content, filetypes=[('excel file', '.xlsx')],
                                            initialfile='_Daily_Reports'+'('+self.stChosen.get()+'至'+self.edChosen.get()+')')
        self.savepath.set(name)

    # 導出文件
    def get_results(self):
        result = []
        if self.openpath.get() and self.data_state:
            # self.pbr.start(interval=10)
            t = time.time()

            # 是否有保存路径
            if self.savepath.get():
                for i in self.result_list:
                    if (i[0] <= self.edChosen.get() and i[0] >= self.stChosen.get()) or (i[0] <= self.stChosen.get() and i[0] >= self.edChosen.get()):
                        result.append(i)

                fname = gui.savepath.get()
                sheet = make_sheet(self.xfile, fname, result)

                if self.openpath.get().find('Top Module') >= 0:
                    self.lb["text"] = '報表正在生成中，請耐心等待'
                    sheet.top_module_sum()
                    sheet.ttl_sum_top()
                else:
                    self.lb["text"] = '報表正在生成中，請耐心等待'
                    sheet.Other_module_sum()
                    sheet.ttl_sum_other()

                # 运行耗时
                self.speed = time.time() - t + self.speed1
                # 设置提示框运行耗时的值
                self.lb1["text"] = '本次數據處理一共耗時:' + str(self.speed) + 's'
                self.lb["text"] = '報表生成成功！'
            # self.pbr.stop()
        else:
            messagebox.showerror(title="錯誤", message="請先打開文件，並進行數據分析！", parent=self.frame_content)


if __name__ == '__main__':
    root = TkinterDnD.Tk()

    # 添加logo画布
    can = tk.Canvas(root, width=150, height=180, bg='white')
    im = tk.PhotoImage(file='Logo.png')
    can.create_image(355, 80, image=im)
    can.pack(fill='both', expand=1, side='bottom')

    gui = Gui(root)
    root.mainloop()

