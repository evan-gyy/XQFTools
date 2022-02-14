# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy
import time

import requests
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from pyquery import PyQuery as pq
from math_clf import math_grade_one_last
# from math_clf import paddle_ocr
from math_clf import tss_ocr
import os
import sys
from PIL import Image
from datetime import datetime

class Simulator:
    def __init__(self):
        self.basic_url = 'https://www.wulouai.com/user-center/my-course-catalogue-record?id=3003270&show_mode=part_operate'
        self.clf_count = [[0 for i in range(5)],
                          [0 for i in range(3)],
                          [0 for i in range(6)]]

    def login(self):
        self.driver = webdriver.Firefox()
        self.wait = WebDriverWait(self.driver, 10)
        cookies = {
            'name': 'remember_user_member_59ba36addc2b2f9401580f014c7f58ea4e30989d',
            'value': 'eyJpdiI6IlNhNmdQVlBra2tmdUhrVDRmS281S2c9PSIsInZhbHVlIjoiK2d2dkpMYnRMT2Q0d00zWDI1ek1WdlZ6dXlIKzdscERUcEttOUJmZEdDeUJENXJqc0lRb2VBb1dET3RYXC8wdlgxUGxFUmUwWkJDalIxcjdOMTZGYU13SGpzWlQxMWRuVHdIbXZlY2ZDeW1FPSIsIm1hYyI6IjdmMTdmZWQ5Zjc3MTFlYzA5YzBlZTIzZWUxM2I1Yzg3YWQ0MjQ1ZDY2NDRiNzBkYjVmNDNiODU0ODhmY2U3MTEifQ'
        }
        self.driver.get("https://www.wulouai.com/login")
        self.driver.add_cookie(cookie_dict=cookies)
        self.driver.get("https://www.wulouai.com/user-center/my-course-catalogue-record?id=3003270&show_mode=part_operate")

    def classify(self):
        pic_path = self.get_pic()
        now = datetime.now()
        with open('log/log{}.txt'.format(now.strftime('%Y-%m-%d_%H_%M_%S')), 'w') as f:
            j = 0
            for i in pic_path:
                # text = paddle_ocr(i)
                self.modify_pic(i, i)
                text = tss_ocr(i)
                cls = math_grade_one_last(text)
                print(cls)
                if cls != 0:
                    self.count(cls)
                    self.choose_ele('.custom_select_text', j, cls)

                f.write(text)
                f.write('\n')
                f.write(str(cls))
                f.write('\n\n')
                j += 1
        self.del_pic(pic_path)

    def count(self, res):
        # 分类计数逻辑
        target = self.clf_count[res[-1][0]][res[-1][1]]
        if target >= 100:
            if self.clf_count[res[-1][0]][-1] >= 100:
                if self.clf_count[0][-1] >= 100 and self.clf_count[1][-1] >= 100 and self.clf_count[2][-1] >= 100:
                    self.driver.close()
            else:
                self.clf_count[res[-1][0]][-1] += 1
        else:
            self.clf_count[res[-1][0]][res[-1][1]] += 1

    def get_pic(self):
        # 获取图片元素列表
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.question_show_div > img')))
        doc = pq(self.driver.page_source)
        pic_element = doc('.question_show_div > img').items()
        pic_url = [i.attr['src'] for i in pic_element]
        # 批量下载图片
        pic_path = []
        for i in pic_url:
            pic = requests.get(i)
            _, pic_name = os.path.split(i)
            path = 'img/{}'.format(pic_name)
            pic_path.append(path)
            with open(path, 'wb') as f:
                f.write(pic.content)

        return pic_path

    def del_pic(self, path):
        for i in path:
            os.remove(i)

    def modify_pic(self, img_path, save_path):
        image = Image.open(img_path, 'r')
        width = image.size[0]
        height = image.size[1]
        if (width != height):
            bigside = width if width > height else height
            background = Image.new('RGBA', (bigside, bigside), (255, 255, 255, 255))
            offset = (int(round(((bigside - width) / 2), 0)), int(round(((bigside - height) / 2), 0)))
            background.paste(image, offset)
            background.save(save_path)

    def choose_ele(self, css, n, cls):
        self.roll(css, n)
        button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'table > tbody > tr:nth-child({n+2}) li')))
        ActionChains(self.driver).move_to_element(button).perform()
        button.click()
        cat_one = self.driver.find_element_by_link_text(cls[0])[-1]
        ActionChains(self.driver).move_to_element(cat_one).perform()
        cat_two = self.driver.find_element_by_link_text(cls[1])[-1]
        ActionChains(self.driver).move_to_element(cat_two).perform()
        cat_two.click()

    def roll(self, css, n):
        ele = self.driver.find_elements_by_css_selector(css)[n]
        self.driver.execute_script("arguments[0].scrollIntoView();", ele)
        return ele

    def run(self):
        self.login()
        # self.classify()
        self.choose_ele('.custom_select_text', 4, math_grade_one_last(tss_ocr('img/test.png')))
        # self.driver.close()

if __name__ == '__main__':
    img_path = 'img/bj.png'
    s = Simulator()
    s.run()
    # s.modify_pic(img_path, '{}-1.png'.format(img_path))