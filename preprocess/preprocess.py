# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import win32com
from win32com.client import Dispatch
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Inches
import os


class Preprocess:
    def __init__(self):
        self.file_list = self.find_file('.docx')
        print(self.file_list)

    def find_file(self, type, dir='.'):
        file_list = []
        for f in os.listdir(dir):
            if type in f:
                file_list.append(os.path.abspath(f))
        return file_list

    def del_frame(self):
        app = win32com.client.DispatchEx('Word.Application')
        app.Visible = 0
        app.DisplayAlerts = 0
        for file in self.file_list:
            doc = app.Documents.Open(file)
            # 删除图文框
            for fra in doc.Content.Frames:
                fra.Delete()
            # 替换文本
            # old = ['A、', 'B、', 'C、', 'D、', '^b', '^m', '^n']
            # new = ['A．', 'B．', 'C．', 'D．', '', '', '']
            old = ['A、', 'B、', 'C、', 'D、']
            new = ['A．', 'B．', 'C．', 'D．']
            for i in range(len(old)):
                app.Selection.Find.ClearFormatting()
                app.Selection.Find.Replacement.ClearFormatting()
                app.Selection.Find.Execute(old[i], False, False, False, False, False, True, 1, False, new[i], 2)

            doc.Save()
            print('deleting frame: {}'.format(os.path.basename(file)))
        app.Quit()

    def set_font(self, run, font_name):
        run.font.name = u"Times New Roman"
        run.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    def set_para_format(self, para):
        pf = para.paragraph_format
        tabs = [2, 10, 18, 26]
        try:
            # print(pf.tab_stops[0].position.inches)
            pf.tab_stops.clear_all()
            for i in tabs:
                pf.tab_stops.add_tab_stop(Inches(i / 6))
        except:
            pass

    def iter_block_items(self, parent):
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
            # print(parent_elm.xml)
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("something's not right")
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def format_word(self, file):
        print('formatting: {}'.format(os.path.basename(file)))
        abs_path = os.path.abspath(file)
        doc = Document(abs_path)
        doc.settings.odd_and_even_pages_header_footer = False
        # 装订线、页边距、页眉页脚
        for section in doc.sections:
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(3.17)
            section.right_margin = Cm(3.17)
            section.gutter = Cm(0)
            # 删除页眉、页脚
            section.different_first_page_header_footer = False
            section.header.is_linked_to_previous = True
            section.footer.is_linked_to_previous = True
        # 修改字体
        for block in self.iter_block_items(doc):
            if isinstance(block, Paragraph):
                self.set_para_format(block)
                for run in block.runs:
                    self.set_font(run, u'汉仪书宋二简')
            elif isinstance(block, Table):
                for row in block.rows:
                    for ce in row.cells:
                        for para in ce.paragraphs:
                            self.set_para_format(para)
                            for run in para.runs:
                                self.set_font(run, u'汉仪书宋二简')

        doc.save(abs_path)

    def run(self):
        self.del_frame()
        for f in self.file_list:
            self.format_word(f)
        input('运行完成')


pp = Preprocess()
pp.run()