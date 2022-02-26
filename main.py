import sys
import os
sys.path.append(os.getcwd())
from docxtpl import DocxTemplate
import pandas as pd
base_dir = os.path.dirname(os.path.abspath(__file__))


def gen_detail():
    df = pd.read_excel(base_dir+'/data/detail2.xlsx').to_dict(orient='records')
    return df


def pursuit_fill():
    doc = DocxTemplate(base_dir+'/data/template3.docx')  # 加载模板文件
    detail = gen_detail()
    for idx, d in enumerate(detail):
        doc.render(d)  # 填充数据
        doc.save('所函-{}-{}.docx'.format(d.get('t1'), idx))  # 保存目标文件


pursuit_fill()

# data_dic = {
# 't1':'燕子',
# 't2':'杨柳',
# 't3':'桃花',
# 't4':'针尖',
# 't5':'头涔涔',
# 't6':'泪潸潸',
# 't7':'茫茫然',
# 't8':'伶伶俐俐',
# }
#
# doc = DocxTemplate('tpl.docx') #加载模板文件
# doc.render(data_dic) #填充数据
# doc.save('demo.docx') #保存目标文件