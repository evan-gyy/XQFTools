# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import pandas as pd
import os

class SortCatalog:
    def __init__(self):
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

    def find_file(self, type, multiple=True):
        file_list = []
        for root, dirs, files in os.walk("."):
            for f in files:
                if type in f and 'res' not in f:
                    file_list.append(f)
        if multiple:
            return file_list
        else:
            print("检测到以下{}文件: ".format(type))
            for f in file_list:
                print("{}: {}".format(file_list.index(f), f))
            while True:
                try:
                    target = file_list[int(input("请输入文件序号："))]
                    return target
                except:
                    print("发生错误：请正确输入文件前的序号（0-n）")

    def match_first(self, s):
        match_list = []
        match_idx = []
        for i in self.order_list:
            if i in s:
                match_list.append(i)
                match_idx.append(s.find(i))
        if match_idx:
            return match_list[match_idx.index(min(match_idx))]
        else:
            return 0

    def reorder(self, excel):
        df = pd.read_excel(excel)
        max_row = df.shape[0]
        new_df = pd.DataFrame(columns=df.columns)
        df['区'] = df['章'].apply(lambda x: self.match_first(x) if isinstance(x, str) else None)
        index_list = []
        for region in self.order_list:
            try:
                i = df[df['区'] == region].index.tolist()[0]
                index_list.append(i)
            except:
                continue
        missing = df[df['区'] == 0].index.tolist()
        if missing:
            index_list.extend(missing)
        sorted_list = sorted(index_list)
        for i in range(len(index_list)):
            j = sorted_list.index(index_list[i])
            start = sorted_list[j]
            end = sorted_list[j + 1] if j + 1 < len(sorted_list) else max_row
            df.loc[sorted_list[j], '章'] = str(i + 1) + "、" + df.loc[sorted_list[j], '章']
            new_df = pd.concat([new_df, df.iloc[start:end, :]])
        del new_df['区']
        new_df.to_excel('res-' + excel, index=False)

    def run(self):
        try:
            file_list = self.find_file('.xlsx')
            for file in file_list:
                print('processing: {}'.format(file))
                self.reorder(file)
        except Exception as e:
            print(e)
            print("运行出错：请核对当前目录下的Excel，并确保Excel可编辑")
            print("可尝试方法：进入Excel双击任一单元格并保存")


def main():
    sc = SortCatalog()
    sc.run()

main()