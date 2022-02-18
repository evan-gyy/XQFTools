# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import os
import pandas as pd
import difflib

class NumberRename:
    """
    编号重命名模式：
    1：题目1-答案1-解析1-题目2-答案2-解析2……
    2：题目1-解析1-答案1-题目2-解析2-答案2……
    3：题目1-题目2……答案1-解析1-答案2-解析2……
    4：题目1-题目2……解析1-答案1-解析2-答案2-……
    5：题目1-题目2……答案1-答案2……解析1-解析2……
    6：题目1-题目2-题目3……
    7：答案1-答案2-答案3……
    8：解析1-解析2-解析3……
    9：题目1-答案1-题目2-答案2……
    10：题目1-题目2……答案1-答案2……
    """
    def __init__(self):
        self.QUESTION = "_word_question"
        self.ANSWER = "_word_answer"
        self.RESOLVE = "_word_resolve"

    def get_names(self, ques_num):
        ques_list = [i + self.QUESTION for i in ques_num]
        ans_list = [i + self.ANSWER for i in ques_num]
        res_list = [i + self.RESOLVE for i in ques_num]
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

    def run_mode(self, mode, ques_num):
        """ direct run with specific mode """
        ql, al, rl = self.get_names(ques_num)
        return self.find_func(mode, ql, al, rl)


class SplitQuestion:
    def __init__(self, path='.'):
        self.path = path
        self.excel = self.find_file('.xlsx', path)
        self.word_list = self.find_file('.docx', path, multiple=True)
        # self.word_list.sort(key=lambda x: int(x.split('-')[0]))
        self.new_ques_num = {}

    def find_file(self, type, path, multiple=False):
        file_list = []
        for f in os.listdir(path):
            if type in f and 'res' not in f and 'name-mode' not in f:
                if not multiple:
                    return f
                else:
                    file_list.append(f)
        return file_list

    def get_ques_num(self):
        cat_df = pd.read_excel(self.path + self.excel)
        titles = []
        ques_num_dict = {}
        count = 0
        for i in range(0, len(cat_df)):
            if isinstance(cat_df.iloc[i]['章'], str):
                count += 1
                cat_df.loc[i, '章'] = str(count) + "、" + cat_df.loc[i, '章']
                ques_num_dict[count] = []
                titles.append(cat_df.iloc[i]['章'])
            num = cat_df.iloc[i]['习题编号']
            if isinstance(num, str):
                ques_num_dict[count].append(num)
        cat_df.to_excel(self.path + 'res-' + self.excel, index=False)
        return ques_num_dict, titles

    def rename_number(self, old_ques_num):
        rn = NumberRename()
        ques_dict = {}
        for i in range(len(self.word_list)):
            new_num = rn.run_mode(self.mode_list[i], old_ques_num[i + 1])
            ques_dict[i + 1] = new_num
        return ques_dict

    def to_file(self, ques_dict, each=True, total=True):
        for f in os.listdir(self.path):
            if f.endswith('.txt'):
                os.remove(os.path.join(self.path, f))
        if each:
            for key, value in ques_dict.items():
                name = self.path + self.word_list[key - 1][:-5] + '_' + str(len(value)) + '.txt'
                if os.path.exists(name):
                    os.remove(name)
                with open(name, "w") as f:
                    for num in value:
                        f.write(num + '\n')
        if total:
            total_n = 0
            name = self.path + 'total.txt'
            f = open(name, "w")
            for key, value in ques_dict.items():
                total_n += len(value)
                for num in value:
                    f.write(num + '\n')
            f.close()
            new_name = self.path + 'total_{}_{}.txt'.format(len(self.word_list), total_n)
            if os.path.exists(new_name):
                os.remove(new_name)
            os.rename(name, new_name)

    def word_rename(self, path, titles, default_mode=1, imp=False):
        if not imp:
            order_list = [i + 1 for i in range(len(titles))]
            mode_list = [default_mode for j in range(len(titles))]
            self.generate_mode_file(order_list, mode_list)
            input('请核对编号-模式文件')

        rename_file = ""
        for f in os.listdir(path):
            if 'name-mode' in f:
                rename_file = f
        if rename_file == "":
            raise ValueError("无法找到编号-模式文件，请检查目录")
        rn_df = pd.read_excel(path + rename_file)
        if rn_df.shape[0] != len(titles):
            raise ValueError("编号-模式数量与文件实际数量不符，请核对")
        for t in range(len(titles)):
            try:
                match = difflib.get_close_matches(titles[t], self.word_list, 1, cutoff=0.7)[0]
            except:
                match = 0
            if match:
                os.rename(path + match, path + f"{rn_df.loc[t, 'order']}-{rn_df.loc[t, 'mode']}.docx")
            else:
                for f in self.word_list:
                    if f.split('-')[0] == str(t + 1):
                        os.rename(path + f, path + f"{rn_df.loc[t, 'order']}-{rn_df.loc[t, 'mode']}.docx")


    def generate_mode_file(self, order: list , mode: list):
        res_df = pd.DataFrame()
        res_df['order'] = order
        res_df['mode'] = mode
        for f in os.listdir(self.path):
            if 'name-mode' in f:
                ans = input('已有编号-模式文件，是否替换？(y/n)：')
                if ans in ['y', '']:
                    res_df.to_excel(self.path + 'name-mode-' + self.excel, index=False)
                    print('已生成编号-模式文件')
                else:
                    pass
                return 0
        res_df.to_excel(self.path + 'name-mode-' + self.excel, index=False)
        print('已生成编号-模式文件')

    def run(self):
        # Step 1
        self.old_ques_num, titles = self.get_ques_num()
        rename = False
        imp = False
        while True:
            ans = input("是否需要重命名？（y/n）：")
            if ans in ['y', '']:
                rename = True
                if input("是否导入编号-模式列表？（y/n）：") in ['y', '']:
                    imp = True
                break
            elif ans == "n":
                break
            else:
                print("请输入y或n")
        if rename and not imp:
            mode = int(input("请输入默认模式："))
            self.word_rename(self.path, titles, mode)
        elif imp:
            self.word_rename(self.path, titles, imp=True)
        input("请核对文件名及模式，按回车继续")
        # Step 2
        self.word_list = self.find_file('.docx', self.path, multiple=True)
        self.word_list.sort(key=lambda x: int(x.split('-')[0]))
        order_list = [x.split('-')[0] for x in self.word_list]
        print('文件列表：', self.word_list)

        self.mode_list = [int(i[:-5].split("-")[1]) for i in self.word_list]
        print('模式列表：', self.mode_list)

        # Step 3
        self.new_ques_num = self.rename_number(self.old_ques_num)
        self.to_file(self.new_ques_num)

def main():
    try:
        path = input("请输入文件夹路径：").replace("\\", "/") + '/'
        sq = SplitQuestion(path=path)
        sq.run()
        input('运行完成')
    except Exception as e:
        print(e)
        input('运行出错，请检查错误')

main()
