# 某企业原油流量计数据处理
帮家里人写的工作用小程序，使用 `pandas` 与 `openpyxl` 处理表格数据，使用`tkinter`制作了简单GUI。

## 构建方法
1. `pip install -r .\src\requirements.txt`
2. `pip install pyinstaller`
3. `pyinstaller .\src\excel_process.spec`

## 使用方法
在程序相同目录下放置原始原油流量计导出数据`1.xlsx`，之后运行程序即可选择指定日期的sheet进行处理，并导出到最终数据汇总表格`3.xlsx`。若没有`3.xlsx`则自动新建一个。
