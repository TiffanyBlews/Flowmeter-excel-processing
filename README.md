# 某企业原油流量计数据处理
帮家里人写的工作用小程序，通过使用`openpyxl`操作Excel表格来代替人工复制粘贴处理数据，使用`tkinter`制作了简单GUI。

## 构建方法
1. `pip install -r .\src\requirements.txt`
2. `pip install pyinstaller`
3. `pyinstaller .\src\excel_process.spec`

## 使用方法
在程序相同目录下放置`1.xlsx`, `2.xlsx`, `3.xlsx`，其中`1.xlsx`为原始原油流量计导出数据、`2.xlsx`为计算表格、`3.xlsx`为最终数据汇总表格。之后运行程序即可选择日期进行处理。