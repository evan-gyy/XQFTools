# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import os
import pandas as pd
from win32com.client import Dispatch
from ctypes import windll

class RenameNumber:
    def __init__(self):
        self.QUESTION = "_word_question"
        self.ANSWER = "_word_answer"
        self.RESOLVE = "_word_resolve"

    def get_names(self, ques_num):
        df = pd.DataFrame(ques_num)
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

    def run_mode(self, mode, ques_num):
        """ direct run with specific mode """
        ql, al, rl = self.get_names(ques_num)
        return self.find_func(mode, ql, al, rl)


class SplitQuestion:
    def __init__(self, path='.'):
        self.excel = self.find_file('.xlsx', path)
        self.word_list = self.find_file('.docx', path, multiple=True)
        self.word_list.sort(key=lambda x: len(x))
        print(self.word_list)
        self.mode_list = [int(i[:-5].split("-")[1]) for i in self.word_list]
        print(self.mode_list)
        self.new_ques_num = {}

    def find_file(self, type, path, multiple=False):
        file_list = []
        for root, dirs, files in os.walk(path):
            for f in files:
                if type in f:
                    if not multiple:
                        return os.path.abspath(f)
                    else:
                        file_list.append(f)
        return file_list

    def get_ques_num(self):
        cat_df = pd.read_excel(self.excel)
        ques_num_dict = {}
        count = 0
        for i in range(0, len(cat_df)):
            if isinstance(cat_df.iloc[i]['章'], str):
                count += 1
                ques_num_dict[count] = []
            num = cat_df.iloc[i]['习题编号']
            if isinstance(num, str):
                ques_num_dict[count].append(num)
        return ques_num_dict

    def rename_number(self, old_ques_num):
        rn = RenameNumber()
        ques_dict = {}
        for i in range(len(self.word_list)):
            new_num = rn.run_mode(self.mode_list[i], old_ques_num[i + 1])
            ques_dict[i + 1] = new_num
        return ques_dict

    def to_file(self, ques_dict, each=True, total=True):
        if each:
            for key, value in ques_dict.items():
                name = self.word_list[key - 1][:-5] + '_' + str(len(value)) + '.txt'
                with open(name, "w") as f:
                    for num in value:
                        f.write(num + '\n')
        if total:
            total_n = 0
            f = open('total.txt', "w")
            for key, value in ques_dict.items():
                total_n += len(value)
                for num in value:
                    f.write(num + '\n')
            f.close()
            os.rename('total.txt', 'total_{}_{}.txt'.format(len(self.word_list), total_n))

    def clear_board(self):
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()

    def copy_doc(self, app, doc, path, page_no, num):
        try:
            doc_add = app.Documents.Add()
            newFile = path + '{}.docx'.format(num)
            doc_add.SaveAs(os.path.abspath(newFile))  # 创建新文件
            doc_new = app.Documents.Open(os.path.abspath(newFile))
            # 页对象
            pages = doc.ActiveWindow.Panes(1).Pages.Count
            if page_no > pages:
                print("指定页索引超出已有页面")
            else:
                objRectangles = doc.ActiveWindow.Panes(1).Pages(page_no).Rectangles
                for i in range(objRectangles.Count):
                    self.clear_board()
                    objRectangles.Item(i + 1).Range.Copy()
                    doc_new.Range(doc_new.Content.End - 1, doc_new.Content.End - 1).Paste()
            doc_new.Save()
            doc_new.Close()
        except Exception as e:
            print(e)

    def cut_word(self):
        app = Dispatch('Word.Application')
        app.Visible = 0
        app.DisplayAlerts = 0
        path = 'result/'
        if not os.path.exists("./" + path):
            os.makedirs("./" + path)
        try:
            for i in range(len(self.word_list)):
                print(self.word_list[i])
                doc = app.Documents.Open(os.path.abspath(self.word_list[i]))
                pages = doc.ActiveWindow.Panes(1).Pages.Count
                for page in range(1, pages + 1):
                    print(page)
                    self.copy_doc(app, doc, path, page, self.new_ques_num[i+1][page-1])
                doc.Close()
        except Exception as e:
            print(e)
        finally:
            app.Quit()

    def run(self):
        old_ques_num = self.get_ques_num()
        self.new_ques_num = self.rename_number(old_ques_num)
        self.to_file(self.new_ques_num)
        # self.cut_word()

def main():
    sq = SplitQuestion()
    sq.run()
    input('运行完成')

main()
