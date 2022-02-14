# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import pandas as pd
import numpy as np
import os

def find_file(type, path):
    file_list = []
    for root, dirs, files in os.walk(path):
        for f in files:
            if type in f:
                file_list.append(f)
    return file_list

def get_res():
    res_list = {}
    fl = []
    for f in find_file(".xlsx", "."):
        df = pd.read_excel(f, usecols=['习题编号'])
        i = df['习题编号'].dropna().tolist()
        print(i[0])
        current = int(i[0][7])
        if current in fl:
            print(current)
            res_list[current].extend(i)
        else:
            fl.append(current)
            res_list[current] = i
    for i in res_list.keys():
        print(i, len(res_list[i]))
    return res_list

def main():
    cat_df = pd.read_excel("../全国版小学数学口算 - 习题集模式目录.xlsx")
    res_list = get_res()
    grade_list = []
    drop_list = []
    for i in range(0, len(cat_df)):
        if isinstance(cat_df.iloc[i]['章'], str):
            grade = cat_df.iloc[i]['章'][0]
            if grade not in grade_list:
                grade_list.append(grade)
                print(len(grade_list))
        current = len(grade_list)
        num = cat_df.iloc[i]['习题编号']
        if isinstance(num, str):
            if num not in res_list[current]:
                drop_list.append(i)

    new_df = cat_df.drop(drop_list)
    new_df.to_excel("../new.xlsx", index=False)

main()