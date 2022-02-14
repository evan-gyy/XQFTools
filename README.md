# XQFTools


![](https://img.shields.io/badge/python-3.6%2B-brightgreen)

XQF办公自动化小工具,提供如下功能:

- 英语组卷导出文件的格式修改,支持Word、PPT
- 对于Word导入流程中的文件,生成重命名题目编号
- 基于排课文件,按系统要求的导入格式写入Excel
- 修改目录Excel中的编号顺序,按区县排序
- 批量删除Word文件中的图文框
- 其他花里胡哨的功能

## 目录

`[TOC]`

## 使用准备


可以直接使用.exe文件,或者使用Python运行

### 常规方式
将要修改的文件、exe文件和相关依赖文件(如有)放在同一个空目录下,双击exe文件运行

### Python
需要有相应Python环境,具体要求如下:

- Python>=3.6
- openpyxl==3.0.9
- pandas==1.3.5
- python_docx==0.8.11
- python_pptx==0.6.21
- pywin32==303
- urwid==2.1.2

使用pip安装依赖包:
```
pip install -r requriements.txt
```

## 组卷文件修改

> PaperExport/

### 功能
修改Word样式,并基于Word将内容写入相应PPT

### 说明
- 使用时确保当前目录只有需要修改的word和ppt,没有其他.docx和.pptx结尾的文件
- 若修改ppt,则需要同时存在word和ppt,程序会提取word中的内容写入ppt;若只修改word,则只需放入word文件
- 请确保文件夹中存在background_image_xqf.png,丢失会导致ppt中出现空白页

## Word导入重命名

> NumberRename/
> SplitQuestion/

### 功能
根据目录文件,重命名题目编号,生成题目、答案、解析三种编号,并按指定模式排列

### 说明
- NumberRename:单文件模式,使用指定Excel文件生成题目编号
- SplitQuestion:多文件模式,使用指定Excel和多个Word文件,批量生成Word文件对应的重命名编号文件
- 多文件模式的具体说明见文件夹内的视频

## 排课表格导入

> Paike/

### 功能
基于排课文件,生成对应的、支持导入系统的Excel文件

### 说明
- 运行后,需要依次选择排课文件和系统导出文件
- 排课文件示例:
	> Paike/总表.xlsx

- 系统导出文件示例:
	> Paike/导出(高一1v5)班级排课.xlsx


## Word批量删除图文框

> DelFrame/

### 功能
批量删除word文件中的图文框

### 说明
- 支持单文件/多文件模式
- 可以删除同一个目录下所有Word文件中的图文框

## Excel目录文件按区县排序

> SortCatalog/

### 功能
将指定目录文件中的内容按区县排序

### 说明
- 默认区县顺序:
  ```python
  self.order_list = [
              '宝山',
              '崇明',
              '长宁',
              '奉贤',
              '虹口',
              '黄浦',
              '静安',
              '嘉定',
              '金山',
              '闵行',
              '浦东',
              '普陀',
              '青浦',
              '松江',
              '徐汇',
              '杨浦',
          ]
  ```
- 输出文件样例:
	> res-原文件名.xlsx
- 支持多文件模式,可以批量修改当前目录下的文件













