Copyright (c) 2023 Kinco Inc. 

typeCast是一个文本库迁移脚本，旨在解决weinview工程中文本标签转换至Dtools的繁琐问题。
使用typeCast进行格式转换接受四项参数：
```
	输入(源)文件地址
	输出(终)文件地址
	语言数量
	状态数量
```
对于输入输出文件地址，相对路径和绝对路径都是可被接受的。
一种建议的使用方式是将typecast(.py/.exe)文件与目标文件放置在相同目录下，这将简化路径的输入。
如创建文件夹并将typeCast与待转化文档置内(.py与.exe按需选择其一即可)，或直接将待转换文件置于此文件夹下
```
typeCast/
	|--readme.txt
	|--typeCast.py
	|--typeCast.exe
	|--origin.xls
```
按照此目录结构存放时，当脚本请求输入文件时仅需输入文件名即可正常读取
```
...
...
请输入源 Excel 文件路径：origin.xls
请输入终 Excel 文件路径：output.xls
...
...
```
以上演示了在同目录下读取一个来自weinview的文本标签库origin.xls
并将按Dtools标准转换后在同目录下创建文档output.xls
```
typeCast/
	|--readme.txt
	|--typeCast.py
	|--typeCast.exe
	|--origin.xls
	|--output.xls
```
余下两项参数：语言数量 和 状态数量 是指输入文档中最大支持的语言数量和状态数量，该值是可缺省的，将会分别使用
默认值6和8填充。该值大于实际值不会造成显著影响，可以直接使用回车跳过输入。但当源文件中语言及状态大于默认值时
应当手动重写以避免文档被截断。

该文件目录包含typeCast的python源码和使用pyinstaller封装的.exe文件

在拥有python环境的计算机上.py文件可以直接运行，但需注意以下依赖库的安装与引入：
```
pip install
	pandas
	pyexcel
	os
```
该程序暂由 Eddie包 独立维护，程序诞生初衷旨在解决本人在weinview工程迁移中.xls文件转换的繁琐问题，由于
文本标签库在格式上与Kinco Dtools存在巨大差异，最为棘手，成为首要解决对象。在日后可能考虑更新版本增加对
事件信息的转换，或其他第三方数据库转换至Kinco Dtools标准的支持。

任何疑问或建议，请联系baozl@kinco.com，Kinco企业内部伙伴可通过企业微信搜索baozl找到我。
<br>在GitHub上可以找到最新版本与更加详细的信息：
<br>https://github.com/Eddie1206/typeCast

本项目采用MIT License。这意味着您可以自由地使用、修改和分发本项目的代码，但需注意应遵循许可证条件，
并在文件中包含来自Kinco Inc.的版权声明。

Kinco Inc为本人对版权申明的偏好，实际上步科公司在官方称呼上使用Kinco LTD。在此，Kinco Inc.与
Kinco LTD指代同一家公司，具有相同法律效力.特此强调，版权归深圳市步科电气有限公司所有。
