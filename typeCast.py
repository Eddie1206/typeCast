#Kinco Inc. All rights reserved.
#Created by Eddie 包 at Kinco on September 18, 2023

import pandas as pd
from pyexcel import Sheet
import pyexcel_io.writers
import pyexcel_xls
import os
#import xlrd

#C:\Users\EDDIE\Desktop\typeCast\origin.xls
#C:\Users\EDDIE\Desktop\typeCast\output.xls

print("------------------------------------------------")
print("Welcome to Kinco Format Converter!")
print("Copyright © 2023 Kinco, Inc.")
print("------------------------------------------------")

while True:
    input_path = input("请输入源 Excel 文件路径：")
    _, ext = os.path.splitext(input_path)
    if ext.lower() == '.xls':
        try:
            df = pd.read_excel(input_path, engine="xlrd", keep_default_na=False, na_values=[''])
            # 检查前三列的列名是否与期望的文本匹配
            headerChecker = ["文字标签名称", "Language 1, 状态 0", "Language 2, 状态 0"]
            if list(df.columns[:3]) == headerChecker:
                # 如果匹配，跳出循环
                break
            else:
                print("输入文件非标准weinview文本库格式，无法转换")
        except FileNotFoundError:
            print("文件不存在，请重新输入")
        except PermissionError:
            print("无权限访问文件，请重新输入")
        except IOError as e:
            print(f"读取文件时发生错误: {e}")
        except Exception as e:
            print(f"读取文件时发生异常: {e}")
    elif ext.lower() == '.xlsx':
        print("目前仅支持 .xls 格式文件。")
    else:
        print("错误：无效的文件格式。")

while True:
    output_path = input("请输入终 Excel 文件路径：")
    _, ext = os.path.splitext(output_path)
    if os.path.isabs(output_path):
        confirm = input("警告：识别到绝对路径，确认要使用此路径吗？(yes/no): ")
        if confirm.lower() != 'yes':
            continue
    if ext.lower() == '.xls':
        break
    elif ext.lower() == '.xlsx':
        print("目前仅支持输出 .xls 格式文件。")
    else:
        print("错误：无效的文件格式。")

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

data = []
language_non_empty = [False] * num_languages

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

data.insert(0, ['文本库标签名', '状态'] + [f'Language {i}' for i in range(1, num_languages + 1)])
data.insert(0, ['Text Lib', 'V200'] + ['' for _ in range(1, num_languages + 1)])
columns = ['文本库标签名', '状态'] + [f'Language{i}' for i in range(1, num_languages + 1)]
new_df = pd.DataFrame(data, columns=columns)

for i, is_non_empty in enumerate(language_non_empty, start=1):
    if not is_non_empty:
        col_to_drop = f"Language{i}"
        new_df.drop(columns=[col_to_drop], inplace=True)

new_df.fillna('', inplace=True)
sheet = Sheet(new_df.values.tolist())
sheet.save_as(output_path)

print("------------------------------------------------")
print(f"转换完成！输出文件已保存到: {output_path}")
print("------------------------------------------------")

os.system("pause")