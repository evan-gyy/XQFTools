# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import os
import pandas as pd

class RenameNumber:
    def __init__(self):
        self.QUESTION = "_word_question"
        self.ANSWER = "_word_answer"
        self.RESOLVE = "_word_resolve"
        self.ql, self.al, self.rl = self.get_names()

    def find_file(self, type):
        file_list = []
        for root, dirs, files in os.walk("."):
            for f in files:
                if type in f:
                    file_list.append(f)
        print("检测到以下{}文件: ".format(type))
        for f in file_list:
            print("{}: {}".format(file_list.index(f), f))
        while True:
            try:
                target = file_list[int(input("请输入文件序号："))]
                return target
            except:
                print("发生错误：请正确输入文件前的序号（0-n）")

    def get_names(self):
        df = pd.read_excel(self.find_file('.xlsx'), usecols=[0])
        df.dropna(inplace=True)
        df['question'] = df.iloc[:, 0].apply(lambda x: x + self.QUESTION)
        df['answer'] = df.iloc[:, 0].apply(lambda x: x + self.ANSWER)
        df['resolve'] = df.iloc[:, 0].apply(lambda x: x + self.RESOLVE)
        ques_list = df['question'].tolist()
        ans_list = df['answer'].tolist()
        res_list = df['resolve'].tolist()
        return ques_list, ans_list, res_list

    def find_func(self, mode, ql, al, rl):
        """ functions for different mode """
        def case_func_1(ql, al, rl):
            result = []
            for i in range(len(ql)):
                result.append(ql[i])
                result.append(al[i])
                result.append(rl[i])
            return result

        def case_func_2(ql, al, rl):
            result = []
            for i in range(len(ql)):
                result.append(ql[i])
                result.append(rl[i])
                result.append(al[i])
            return result

        def case_func_3(ql, al, rl):
            result = []
            for i in range(len(ql)):
                result.append(ql[i])
            for i in range(len(al)):
                result.append(al[i])
                result.append(rl[i])
            return result

        def case_func_4(ql, al, rl):
            result = []
            result += ql
            for i in range(len(al)):
                result.append(rl[i])
                result.append(al[i])
            return result

        def case_func_5(ql, al, rl):
            result = []
            result += ql
            result += al
            result += rl
            return result

        def case_func_6(ql, al, rl):
            return ql

        def case_func_7(ql, al, rl):
            return al

        def case_func_8(ql, al, rl):
            return rl

        def case_func_9(ql, al, rl):
            result = []
            for i in range(len(ql)):
                result.append(ql[i])
                result.append(al[i])
            return result

        def case_func_10(ql, al, rl):
            result = []
            result += ql
            result += al
            return result

        func_map = {
            1: case_func_1,
            2: case_func_2,
            3: case_func_3,
            4: case_func_4,
            5: case_func_5,
            6: case_func_6,
            7: case_func_7,
            8: case_func_8,
            9: case_func_9,
            10: case_func_10,
        }
        if mode not in func_map:
            return 0

        return func_map[mode](ql, al, rl)

    def run_mode(self, mode):
        """ direct run with specific mode """
        return self.find_func(mode, self.ql, self.al, self.rl)

    def run_input(self):
        tips = """模式：
    ①题目1-答案1-解析1-题目2-答案2-解析2……
    ②题目1-解析1-答案1-题目2-解析2-答案2……
    ③题目1-题目2……答案1-解析1-答案2-解析2……
    ④题目1-题目2……解析1-答案1-解析2-答案2-……
    ⑤题目1-题目2……答案1-答案2……解析1-解析2……
    ⑥题目1-题目2-题目3……
    ⑦答案1-答案2-答案3……
    ⑧解析1-解析2-解析3……
    ⑨题目1-答案1-题目2-答案2……
    ⑩题目1-题目2……答案1-答案2……"""
        while True:
            print(tips)
            n = input("请输入模式编号(整数)：")
            try:
                n = int(n)
            except:
                print("请输入整数")
                continue

            result = self.find_func(n, self.ql, self.al, self.rl)
            if not isinstance(result, list):
                print("模式不存在，请重新输入")
                continue

            for i in result:
                print(i)

            with open("mode_{}_output.txt".format(n), "w") as f:
                for i in result:
                    f.write(i + '\n')

def main():
    rn = RenameNumber()
    r = rn.run_mode(9)
    print(r)

main()