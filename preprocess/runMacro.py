# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import win32com
from win32com.client import Dispatch
import traceback
from alive_progress import alive_bar
import os

file = r"D:\SynologyDrive\GYY\XQF\XQFTools\preprocess\八年级第一学期练习.docx"
app = win32com.client.DispatchEx('Word.Application')
app.Visible = 0
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(file)
    for i in range(doc.Shapes.Count - 1, -1, -1):
        print('\rshape: {}'.format(i), end='', flush=True)
        shp = doc.Shapes[i]
        if shp.Type == 17:
            string = shp.TextFrame.TextRange.Text[:shp.TextFrame.TextRange.Characters.Count - 1]
            if len(string) > 0:
                rng = shp.Anchor.Paragraphs[0].Range
                rng.InsertBefore(string)
            shp.Delete()
    doc.Save()
except:
    traceback.print_exc()
app.Quit()
