import tkinter as tk  # 使用Tkinter前需要先导入
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename


# 实例化object，建立窗口window

#tex's type is str,is table name, row and column is number
def Big_Table(text, row, column, origin_path, save_path):
    t1 = tk.Label(window, text = text, bg='green')
    t1.grid(row = row, column = column)
    gt1 = tk.Button(window, text='更多设置', command=LittleWin, width=12, height=1)
    gt1.grid(row = row + 1, column = column + 2)
    # 选择要处理的excel
    sf = tk.Button(window, text='报表组原始数据', command=SourceFile, width=12, height=1)
    sf.grid(row = row + 2, column = column + 2)
    # 将选择的Excel的地址写出
    sf1 = tk.Entry(window, textvariable = origin_path, state='readonly', width=30)
    sf1.grid(row = row + 2, column = column + 3)
    # 生保存目标excel
    tf = tk.Button(window, text='保存的目标文件', command=TargetFile, width=12, height=1)
    tf.grid(row = row + 3, column = column + 2)
    # 保存目标excel的路径
    tf1 = tk.Entry(window, textvariable = save_path, state='readonly', width=30)
    tf1.grid(row = row + 3, column = column + 3)
    # 生成Excel的按键
    """
    lb = tk.Label(window, text='点击下面的按键生成目标文件', bg='green')
    lb.grid(row = row + 5, column = column + 3)
    gf = tk.Button(window, text='生成目标文件', command=GenerateFile, width=12, height=1)
    gf.grid(row = row + 6, column = column + 3)
    """
    # --------------------------------------------------------------

def LittleWin():
    #生成小弹框
    list = ['输入如BG，BV', '输入如Mini，Main', '输入如MLB，Scorpirs,AARM-1', '以逗号隔开', '以逗号隔开']
    window1 = tk.Tk()
    window1.title('获取专案名称')
    window1.geometry('600x300')
    window1.resizable(False, False)
    for row, str in enumerate(list):
        value = str
        enter = Entry(window1)
        enter.insert(1, value)
        enter.grid(row=row, column=1, padx=30, pady=5)
        case_name = tk.Label(window1, text='专案名称', width=12, height=1)
        case_name.grid(row=0, column=0)
        build_stage = tk.Label(window1, text='Build stage', width=12, height=1)
        build_stage.grid(row=1, column=0)
        station_type = tk.Label(window1, text='工站类型', width=20, height=1)
        station_type.grid(row=2, column=0)
        build_stage = tk.Label(window1, text='工站名称', width=12, height=1)
        build_stage.grid(row=3, column=0)
        station_type = tk.Label(window1, text='工站ID', width=12, height=1)
        station_type.grid(row=4, column=0)
        gf = tk.Button(window1, text='保存', command=GetStaName, width=12, height=1)
        gf.grid(row=6, column=0)
        gf = tk.Button(window1, text='退出', command=window1.destroy, width=12, height=1)
        gf.grid(row=6, column=2)


# 显示出源文件路径
def SourceFile():
    path_ = askopenfilename(filetypes=[('Excel', 'xlsx')])
    origin_path.set(path_)
def Reference():
    path_ = askopenfilename(filetypes=[('Excel', 'xlsx')])
    _path.set(path_)
def TargetFile():
    path_ = asksaveasfilename(filetypes=[('Excel', 'xlsx')])
    save_path.set(path_)
def GenerateFile():
    print('OK')
def GetStaName():
    enter = Entry(window1)
    get_str = enter.get()
    print(get_str)


if __name__ == '__main__':
    window = tk.Tk()
    window.title('Excel报表生成软件')
    window.geometry('800x600')
    window.resizable(False, False)  # 是否可以設置大小
    # 初始化需要的变量
    origin_path = tk.StringVar()  # 原始文件路径
    res_path = tk.StringVar()  # 生成文件路径
    save_path = tk.StringVar()  # 保存文件路径

    for i in range(0, 38, 8):
        Big_Table('table' + str(int(i/8 + 1)), i, 0, origin_path, save_path)

    lb = tk.Label(window, text='点击下面的按键生成目标文件', bg='green')
    lb.grid(row=38, column= 3)
    gf = tk.Button(window, text='生成目标文件', command=GenerateFile, width=12, height=1)
    gf.grid(row=40, column =3)


    # 大窗口退出
    gf = tk.Button(window, text='退出', command=window.quit, width=12, height=1)
    gf.grid(row=40, column=4)

    window.mainloop()

