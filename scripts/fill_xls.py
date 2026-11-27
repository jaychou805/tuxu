# -*- coding: utf-8 -*-
"""
图需 .xls 填写脚本模板：用 WPS COM 保留图片回写单元格。

用法：
1. 改 SRC / FINAL 路径
2. 改 SHEET_NAME（目标 sheet 名）
3. 改 CELLS：{行号(0-based): {列号(1-based): 文本}}
   常见列：B列=2(具体文案说明), C列=3(卖点说明), D列=4(画面内容), H列=8(需求细则)
4. 运行：<venv>/Scripts/python fill_xls.py

依赖：xlrd pywin32（先 pip install）
"""
import win32com.client
import os, shutil, time

SRC = r'C:\path\to\源文件.xls'
FINAL = r'C:\path\to\输出文件.xls'
SHEET_NAME = '主图做图需求文档'

# 行号(0-based) -> {列号(1-based) -> 文本}
CELLS = {
    # 6: {2: '尺寸（1600x1600）\n\n主标题文案：...\n副标题文案：...\n其他文案：...',
    #     3: '布局：...\n标题位置：...\n色调：...'},
}

TMP = os.path.join(os.path.dirname(FINAL), '_tmp_fill.xls')
shutil.copy2(SRC, TMP)

app = win32com.client.DispatchEx('KET.Application')  # 强制新实例
for attr in ('Visible', 'DisplayAlerts'):
    try:
        setattr(app, attr, False)
    except Exception:
        pass

try:
    wb = app.Workbooks.Open(TMP)
    ws = None
    for s in wb.Worksheets:
        if s.Name == SHEET_NAME:
            ws = s
            break
    if ws is None:
        raise RuntimeError('sheet 未找到: ' + SHEET_NAME)

    for r, cols in CELLS.items():
        for c, text in cols.items():
            ws.Cells(r + 1, c).Value = text

    wb.Save()
    try:
        wb.Close(False)
    except Exception:
        pass
    print('tmp SAVED')
finally:
    app.Quit()

for attempt in range(3):
    try:
        shutil.copy2(TMP, FINAL)
        print('FINAL UPDATED:', FINAL)
        os.remove(TMP)
        break
    except PermissionError:
        print(f'原文件被占用(第{attempt+1}次)，2秒后重试...')
        time.sleep(2)
else:
    print('原文件被占用，已保留临时文件:', TMP)
