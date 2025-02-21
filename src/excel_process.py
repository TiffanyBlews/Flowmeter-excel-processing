import os
import sys
import re
import random
import datetime
import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog
import openpyxl

class ExcelProcessor:
    def __init__(self, raw_file, output_file):
        assert os.path.exists(raw_file), "原始数据文件不存在"
        self.raw_file = raw_file

        if not os.path.exists(output_file):
            df = pd.DataFrame(columns=['日期', '人工化验液量（刮板）（t/d）', '人工化验液量（质量）（t/d）', '人工化验油量（t/d）', '人工含水（%）', '含水仪油量（t/d）', '含水仪含水（%）'])
            df.to_excel(output_file, index=False)
        self.output_file = output_file
        

    def process_sheet(self, sheet_name, date_obj):
        df_raw = pd.read_excel(self.raw_file, sheet_name=sheet_name)
        result = self.calculate_values(df_raw)
        df_result = pd.read_excel(self.output_file, sheet_name='Sheet1')
        
        if date_obj in df_result['时间'].values:
            df_result.loc[df_result['时间'] == date_obj, df_result.columns[1:7]] = result
        else:
            df_result.loc[len(df_result), df_result.columns[:7]] = [date_obj] + result

        df_result.to_excel(self.output_file, index=False)

    def calculate_values(self, df):
        data = df.iloc[2:28].copy()
        processed = pd.DataFrame()
        result = [None] * 6

        processed['VolumeFlowReading'] = data.iloc[:, 3]
        processed['VolumeFlowDiff'] = processed['VolumeFlowReading'].diff()
        processed['VolumeFlowDiffCalculated'] = processed['VolumeFlowDiff'] * 0.889
        result[0] = processed['VolumeFlowDiffCalculated'].sum()

        processed['MassFlowReading'] = data.iloc[:, 5]
        processed['MassFlowDiff'] = processed['MassFlowReading'].diff()
        processed['MassFlowDiffCalculated'] = processed['MassFlowDiff'] * 0.889
        result[1] = processed['MassFlowDiffCalculated'].sum()

        processed['WaterManual'] = data.iloc[:, 9].apply(lambda x: x if pd.notna(x) else random.normalvariate(0.1, 0.05))
        processed['WaterOnline'] = data.iloc[:, 10].apply(lambda x: x if pd.notna(x) else random.normalvariate(0.06, 0.02))
        result[3] = float(processed['WaterManual'].mean())
        result[5] = float(processed['WaterOnline'].mean())

        result[2] = float(result[0] * (1 - 0.01 * result[3]))
        result[4] = (processed['MassFlowDiffCalculated'] * (1 - 0.01 * processed['WaterOnline'])).sum()
        print(result)
        return result

class ExcelProcessorGUI:
    def __init__(self, root, processor):
        self.root = root
        self.processor = processor
        self.root.title("Excel 数据处理器")

        self.sheet_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, height=20)
        self.sheet_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=self.sheet_listbox.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheet_listbox.config(yscrollcommand=self.scrollbar.set)

        self.run_button = tk.Button(root, text="运行", command=self.run_process)
        self.run_button.pack()

        self.clear_button = tk.Button(root, text="清空选择", command=self.clear_list)
        self.clear_button.pack()

        self.populate_sheet_listbox()

    def populate_sheet_listbox(self):
        try:
            wb = openpyxl.load_workbook(self.processor.raw_file, read_only=True)
            for sheet_name in wb.sheetnames:
                self.sheet_listbox.insert(tk.END, sheet_name)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载工作簿: {e}")

    def clear_list(self):
        self.sheet_listbox.selection_clear(0, tk.END)

    def run_process(self):
        selected_sheets = [self.sheet_listbox.get(i) for i in self.sheet_listbox.curselection()]
        if not selected_sheets:
            messagebox.showwarning("警告", "请至少选择一个表格进行处理")
            return

        for sheet in selected_sheets:
            date_str = self.normalize_date(sheet)
            print(date_str)
            try:
                date_obj = datetime.datetime.strptime(date_str, "%Y.%m.%d")
            except ValueError:
                try:
                    date_obj = datetime.datetime.strptime(date_str, "%y.%m.%d")
                except ValueError:
                    date_str = simpledialog.askstring("输入日期", f"{sheet} 不是有效的日期格式 %Y.%m.%d，请输入正确的日期格式 (如 2024.12.20)：")
                    try:
                        date_obj = datetime.datetime.strptime(date_str, "%Y.%m.%d")
                    except ValueError:
                        messagebox.showerror("错误", "无效的日期格式，跳过此表格")
                        continue
            
            self.processor.process_sheet(sheet, date_obj)
            print(f"====={sheet} 处理完成=====")

    def normalize_date(self, date_str):
        return re.sub(r'\.+', '.', date_str).strip('. ')

if __name__ == "__main__":
    current_path = sys.argv[0]
    cwd = os.path.dirname(current_path)    
    raw_file = os.path.join(cwd, '1.xlsx')
    output_file = os.path.join(cwd, '3.xlsx')

    processor = ExcelProcessor(raw_file, output_file)

    root = tk.Tk()
    app = ExcelProcessorGUI(root, processor)
    root.mainloop()
