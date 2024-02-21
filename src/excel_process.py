import os
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox, Scrollbar
import openpyxl
import datetime
from win32com.client import DispatchEx
import random
import re

current_path = sys.argv[0]
cwd = os.path.dirname(current_path)

class ExcelProcessor:
    def __init__(self, raw_file, calculation_file, summary_file):
        self.raw_file = os.path.join(cwd, raw_file)
        # just_open(self.raw_file)
        self.calculation_file = os.path.join(cwd, calculation_file)
        self.summary_file = os.path.join(cwd, summary_file)


    def first_1to2(self, date):
        '''
        将1表中的原始数据复制到2表得出计算结果
        '''
        # 读取原始数据表格
        wb1 = openpyxl.load_workbook(self.raw_file,data_only=True, read_only=True)
        # 读取计算表格
        wb2 = openpyxl.load_workbook(self.calculation_file)

        # 获取原始数据sheet
        ws1 = wb1[date]
        # 获取计算表格sheet
        if wb2.__contains__(date) is False:
            self.new_sheet(date)
            wb2 = openpyxl.load_workbook(self.calculation_file)
            print('new_sheet',date)

        ws2 = wb2[date]

        # 复制流量计数据到计算表格
        for row in range(4, 29):

            for col in range(4, 8):
                ws2.cell(row=row, column=col - 2).value = ws1.cell(row=row, column=col).value

            if row != 4:
                ws2.cell(row, 3).value = ws2.cell(row, 2).value - ws2.cell(row-1, 2).value
                ws2.cell(row, 5).value = ws2.cell(row, 4).value - ws2.cell(row-1, 4).value
            for col in range(2, 6):
                print(row, ws2.cell(row=row, column=col).value, end='')
            print('')

        # 复制含水数据到计算表格
        for row in range(4, 29):
            for col in range(10, 12):
                ws2.cell(row=row, column=col - 3).value = ws1.cell(row=row, column=col).value
                print(ws2.cell(row=row, column=col - 3).value, end='')
            print('')

        # 补全G4:G28中的空值为0.1左右的随机数，H4:H28中的空值为0.06左右的随机数
        for row in range(4, 29):
            if ws2.cell(row=row, column=7).value is None:
                ws2.cell(row=row, column=7).value = random.normalvariate(0.1, 0.03)
            if ws2.cell(row=row, column=8).value is None:
                ws2.cell(row=row, column=8).value = random.normalvariate(0.06, 0.003)

        ws2.cell(1,1).value = date

        # 保存计算表格
        wb2.save(self.calculation_file)

        just_open(self.calculation_file)

    def second_2to3(self, date):
        '''将2表中的结果粘贴到3表'''

        # 读取计算表格
        wb2 = openpyxl.load_workbook(self.calculation_file, data_only=True, read_only=True)
        # 读取汇总表格
        wb3 = openpyxl.load_workbook(self.summary_file)

        # 获取计算表格sheet
        ws2 = wb2[date]
        # 获取汇总表格sheet
        ws3 = wb3['Sheet1']

        # print(ws3.cell(3,1).value)

        # 获取计算结果数据
        result_data = []
        for col in range(11, 15):
            result_data.append(ws2.cell(row=29, column=col).value)
        for col in range(7, 9):
            result_data.append(ws2.cell(row=29, column=col).value)

        # excel_date = convert_to_excel_date(date)
        
        date_obj = datetime.datetime.strptime(date, '%y.%m.%d')
        
        # 获取日期对应的行
        date_column = ws3['A']
        for idx, cell in enumerate(date_column):
            if idx <=1:
                continue
            
            if cell.value == date_obj: # 直接用`==`比较两个datetime对象
                row_index = idx + 1  # 因为索引从0开始，所以需要+1
                break
        else:
            row_index = len(date_column) + 1
            ws3.cell(row=row_index, column=1).value = date_obj.strftime('%Y{y}%m{m}%d{d}')\
                .format(y='年',m='月', d='日') # 编码问题，见 https://blog.csdn.net/lanxingbudui/article/details/124018316

        # 将计算结果数据复制到汇总表格中
        for col, value in zip([2, 4, 3, 6, 5, 7], result_data):
            print(value)
            ws3.cell(row=row_index, column=col).value = value

        # 保存汇总表格
        wb3.save(self.summary_file)

    def new_sheet(self, date):
        # 读取计算表格
        wb2 = openpyxl.load_workbook(self.calculation_file)

        # 获取计算表格sheet
        ws2 = wb2['模板']
        
        new_ws2 = wb2.copy_worksheet(ws2)
        new_ws2.title = date
        wb2.save(self.calculation_file)

def just_open(filename):
    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(os.path.join(cwd, filename))
    xlBook.Save()
    xlBook.Close()


def convert_to_excel_date(date_string):
    # 将日期字符串解析为 datetime 对象
    date_obj = datetime.datetime.strptime(date_string, '%y.%m.%d')

    # 计算日期与 1899 年 12 月 31 日之间的天数差
    delta_days = (date_obj - datetime.datetime(1899, 12, 31)).days

    # 将天数差加上 Excel 日期的起始值
    excel_date = delta_days + 1  # Excel日期起始值是1900年1月1日

    return excel_date


# processor = ExcelProcessor('1.xlsx', '2.xlsx', '3.xlsx')
# date = '24.2.4'
# processor.first_1to2(date)
# processor.second_2to3(date)


class ExcelProcessorGUI:
    def __init__(self, root, raw_file, calculation_file, summary_file):
        # raw_file = simpledialog.askstring("输入文件名", "请输入原油流量计记录表格文件名：")
        self.root = root
        self.root.title("Excel Processor")

        # 创建一个自定义字体
        custom_font = ("Arial", 14)

        # 创建一个滚动条
        self.scrollbar = Scrollbar(root, orient=tk.VERTICAL)

        # 创建列表框并关联滚动条
        self.sheet_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, font=custom_font, height=25,
                                         yscrollcommand=self.scrollbar.set)

        # 设置滚动条与列表框的关联
        self.scrollbar.config(command=self.sheet_listbox.yview)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)


        self.run_button = tk.Button(root, text="运行", command=self.run_process, font=custom_font)  # 设置按钮的字体
        self.run_button.pack()

        self.error_file = "error_sheets.txt"
        self.error_sheets = []

        self.processor = ExcelProcessor(raw_file, calculation_file, summary_file)

        self.populate_sheet_listbox()

    def populate_sheet_listbox(self):
        try:
            wb = openpyxl.load_workbook(self.processor.raw_file, read_only=True)
            sheet_names = wb.sheetnames
            for sheet_name in sheet_names:
                self.sheet_listbox.insert(0, sheet_name)  # 在索引 0 处插入新的 sheet 名称
        except Exception as e:
            messagebox.showerror("错误", f"无法加载工作簿: {e}")

    def run_process(self):
        selected_sheets = self.sheet_listbox.curselection()
        if not selected_sheets:
            messagebox.showwarning("警告", "请至少选择一个表格进行处理")
            return

        for index in selected_sheets:
            sheet_name = self.sheet_listbox.get(index)
            sheet_name = self.clean_sheet_name(sheet_name)
            print("====正在处理"+sheet_name+"====")
            if not self.is_valid_date(sheet_name):
                messagebox.showinfo("提示", f"表格 {sheet_name} 的名称不符合 %y.%m.%d 格式，请手动输入一个日期")
                print(f"表格 {sheet_name} 的名称不符合 %y.%m.%d 格式，请手动输入一个日期")
                real_date = self.get_valid_date_from_user(sheet_name)
                if real_date:
                    self.rename_sheet(sheet_name, real_date)
                    sheet_name = real_date
                else:
                    continue

            try:
                self.processor.first_1to2(sheet_name)
                self.processor.second_2to3(sheet_name)
                messagebox.showinfo("完成", f"{sheet_name} 处理完成")
                print("====处理完成"+sheet_name+"====")
            except Exception as e:
                messagebox.showerror("错误", f"处理 {sheet_name} 时出现错误: {e}")
                print("====处理失败"+sheet_name+"====")
                self.error_sheets.append(sheet_name)
        if len(self.error_sheets)>=1:
            self.save_error_sheets()

    def clean_sheet_name(self, sheet_name):
        # 匹配 %y.%m.%d 格式的日期字符串
        match = re.match(r'\d{2,4}\.\d{1,2}\.\d{1,2}', sheet_name)
        if match:
            # 获取匹配到的日期字符串
            date_str = match.group()
            # 使用 .split() 方法以点为分隔符分割字符串，并只保留中间两个点
            parts = date_str.split('.')
            real_date = '.'.join(parts[:2] + parts[-1:])
            self.rename_sheet(sheet_name, real_date)
            sheet_name = real_date
            return sheet_name
        else:
            return sheet_name
    def is_valid_date(self, date_string):
        try:
            datetime.datetime.strptime(date_string, '%y.%m.%d')
            return True
        except ValueError:
            return False

    def get_valid_date_from_user(self, sheet_name):
        new_date = simpledialog.askstring("输入日期", f"表格 {sheet_name} 的名称不符合 %y.%m.%d 格式，请修改成该格式，比如 24.1.10：")
        return new_date.strip() if new_date else None

    def rename_sheet(self, old_name, new_name):
        try:
            wb1 = openpyxl.load_workbook(self.processor.raw_file)
            ws = wb1[old_name]
            ws.title = new_name
            wb1.save(self.processor.raw_file)

        except Exception as e:
            messagebox.showerror("错误", f"重命名 {old_name} 时出现错误: {e}")

    def save_error_sheets(self):
        try:
            with open(self.error_file, 'w') as file:
                for sheet_name in self.error_sheets:
                    file.write(sheet_name + '\n')
            messagebox.showinfo("完成", f"出错的表格名称已保存到文件: {self.error_file}")
        except Exception as e:
            messagebox.showerror("错误", f"保存出错的表格名称时出现错误: {e}")

root = tk.Tk()
app = ExcelProcessorGUI(root, '1.xlsx', '2.xlsx', '3.xlsx')
root.mainloop()
