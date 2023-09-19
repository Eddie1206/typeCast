# MIT License
# Copyright (c) 2023 Kinco Inc.
#Created by Eddie 包 at Kinco on September 18, 2023

import pandas as pd
from pyexcel import Sheet
import pyexcel_io.writers
import pyexcel_xls
import os
#import xlrd
#import部分可根据实际使用环境调整，部分引入为pyinstaller打包时所需，在直接执行.py时不必要

print("------------------------------------------------")
print("Welcome to Kinco Format Converter!")
print("Copyright © 2023 Kinco, Inc.")
print("------------------------------------------------")

#使用while循环读取参数，直至参数格式正确break
while True:
    input_path = input("请输入源 Excel 文件路径：")
    _, ext = os.path.splitext(input_path)
    if ext.lower() == '.xls':
        break
    elif ext.lower() == '.xlsx':
        print("目前仅支持 .xls 格式文件。")
    else:
        print("错误：无效的文件格式。")

while True:
    output_path = input("请输入终 Excel 文件路径：")
    _, ext = os.path.splitext(output_path)
    if ext.lower() == '.xls':
        break
    elif ext.lower() == '.xlsx':
        print("目前仅支持输出 .xls 格式文件。")
    else:
        print("错误：无效的文件格式。")

#最大语言和状态的输入可缺省，读入空字符串时直接采用默认值
while True:
    num_languages = input("请输入语言数量（默认为 8）：")
    if num_languages == '':
        num_languages = 8
        break
    elif num_languages.isdigit() and int(num_languages) > 0:
        num_languages = int(num_languages)
        break
    else:
        print("错误：请输入一个大于0的整数。")

while True:
    num_states = input("请输入状态数量（默认为 6）：")
    if num_states == '':
        num_states = 6
        break
    elif num_states.isdigit() and int(num_states) > 0:
        num_states = int(num_states)
        break
    else:
        print("错误：请输入一个大于0的整数。")

#文件读入df, 转换后数据首先以数组data保存，language_non_empty标记已使用的语言列
df = pd.read_excel(input_path, engine="xlrd", keep_default_na=False, na_values=[''])
data = []
language_non_empty = [False] * num_languages

#格式重排版
for _, row in df.iterrows():
    label = row['文字标签名称']
    for state in range(num_states):
        state_data = [label, state]
        non_empty_texts = False
        for lang in range(1, num_languages + 1):
            col_name = f"Language {lang}, 状态 {state}"
            text = row[col_name]
            state_data.append(text)
            if pd.notna(text):
                non_empty_texts = True
                language_non_empty[lang - 1] = True
        if non_empty_texts:
            data.append(state_data)
        label = ''

#添加Kinco表头，随后保存为dataFrame
data.insert(0, ['文本库标签名', '状态'] + [f'Language {i}' for i in range(1, num_languages + 1)])
data.insert(0, ['Text Lib', 'V200'] + ['' for _ in range(1, num_languages + 1)])
columns = ['文本库标签名', '状态'] + [f'Language{i}' for i in range(1, num_languages + 1)]
new_df = pd.DataFrame(data, columns=columns)

#丢弃空语言列
for i, is_non_empty in enumerate(language_non_empty, start=1):
    if not is_non_empty:
        col_to_drop = f"Language{i}"
        new_df.drop(columns=[col_to_drop], inplace=True)

#空字符串填充NaN随后存储，使用pyexcel写入文件
#由于pandas在最新版本中取消了对xlwt的支持，因此在写入方式上采用了与读取方式不同的技术
new_df.fillna('', inplace=True)
sheet = Sheet(new_df.values.tolist())
sheet.save_as(output_path)

print("------------------------------------------------")
print(f"转换完成！输出文件已保存到: {output_path}")
print("------------------------------------------------")

#暂锁界面以便用户看到成功提示
os.system("pause")
