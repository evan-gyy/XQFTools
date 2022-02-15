# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import win32com
from win32com.client import Dispatch
import os

def find_file(type, dir='.'):
    file_list = []
    for root, dirs, files in os.walk(dir):
        for f in files:
            if type in f:
                file_list.append(os.getcwd() + '/' + f)
    return file_list

def main():
    app = win32com.client.DispatchEx('Word.Application')
    app.Visible = 0
    app.DisplayAlerts = 0
    try:
        for f in find_file('.docx'):
            doc = app.Documents.Open(f)
            for fra in doc.Content.Frames:
                fra.Delete()
            doc.Save()
            print('done: {}'.format(f))
    except Exception as e:
        print(e)
    app.Quit()
    input('运行完成')

main()