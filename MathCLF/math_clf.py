# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: gyy

import pytesseract
from PIL import Image
# from paddleocr import PaddleOCR, draw_ocr
import os
import re

home_dir = 'D:\\Programs\\TSS\\'
os.environ['TESSDATA_PREFIX'] = home_dir + 'tessdata'
pytesseract.pytesseract.tesseract_cmd = home_dir + 'tesseract.exe'

def tss_ocr(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, config='--psm 6').strip()
    return text.strip()

# def paddle_ocr(image_path):
#     ocr = PaddleOCR(use_angle_cls=True, lang='ch')
#     res = ocr.ocr(image_path, cls=True)
#     text = [line[1][0] for line in res]
#     return ' '.join(text).strip()

def math_grade_one_first():
    pass

def math_grade_one_last(text):
    class_dict = {
        '20以内退位减法': ['十几减9  例如：14-9',
                     '十几减8  例如：12-8',
                     '十几减7、6  例如：14-7',
                     '十几减5、4、3、2  例如：13-5',
                     '综合练习'],
        '100以内数的认识和加减': ['比较大小  例如：54○63',
                         '整十数加减一位数  例如：20+3 48-20',
                         '综合练习'],
        '100以内加法和减法（一）': ['整十数加减整十数  例如：40-20  50+30',
                          '两位数加一位数、整十数（不进位）  例如：43+5 、46+20',
                          '两位数加一位数（进位）  例如：35+7',
                          '两位数减一位数、整十数（不退位）  例如：37-6、48-20',
                          '两位数减一位数（退位）  例如：45-8',
                          '综合练习']
    }
    cls_key = list(class_dict.keys())

    try:
        # 比较题
        if len(text.split(' ')) > 1:
            return cls_key[1], class_dict[cls_key[1]][0], [1, 0]
        text = text.replace('§', '5')
        text = re.sub(r"[=]", "", text)
        print('text: {}'.format(text))

        num_list_str = re.split('-|\+', re.sub(r"[()]", "", text))
        num_list = [int(i) for i in num_list_str]
        # print('num_list: {}'.format(num_list))

        if len(num_list) > 2 and '(' not in text and max(num_list) < 10:
            formula = text[:text.index(num_list_str[1])+len(num_list_str[1])]
            a = eval(formula)
            if a < 10:
                return 0
            b = num_list[2]
        elif '(' in text:
            a = num_list[1]
            b = num_list[2]
        else:
            a = max(num_list)
            try:
                b = num_list[num_list.index(a) + 1]
            except:
                b = num_list[num_list.index(a) - 1]

        # 20以内减法
        if a >= 10 and a <= 20:
            if '-' in text:
                if b == 9:
                    return cls_key[0], class_dict[cls_key[0]][0], [0, 0]
                elif b == 8:
                    return cls_key[0], class_dict[cls_key[0]][1], [0, 1]
                elif b == 7 or b == 6:
                    return cls_key[0], class_dict[cls_key[0]][2], [0, 2]
                elif b >= 2 and b <= 5:
                    return cls_key[0], class_dict[cls_key[0]][3], [0, 3]
                else:
                    return 0
            else:
                return 0
        # 100以内
        elif a <= 100:
            if a % 10 == 0:
                if b < 10:
                    return cls_key[1], class_dict[cls_key[1]][1], [1, 1]
                elif b % 10 == 0:
                    return cls_key[2], class_dict[cls_key[2]][0], [2, 0]
                else:
                    return 0
            elif b < 10 or b % 10 == 0:
                if '+' in text:
                    # 判断进位
                    if (a + b) // 10 == max(a, b):
                        return cls_key[2], class_dict[cls_key[2]][1], [2, 1]
                    else:
                        return cls_key[2], class_dict[cls_key[2]][2], [2, 2]
                elif '-' in text:
                    if (a - b) // 10 == max(a, b):
                        return cls_key[2], class_dict[cls_key[2]][3], [2, 3]
                    else:
                        return cls_key[2], class_dict[cls_key[2]][4], [2, 4]
                else:
                    return 0
            else:
                return 0
        else:
            return 0
    except:
        return 0


if __name__ == '__main__':
    img_path = 'img/9+(15-6).png'
    text = tss_ocr(img_path)
    cls = str(math_grade_one_last(text))
    print(cls)