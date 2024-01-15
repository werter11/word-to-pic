# -*- coding: utf-8 -*-
import os
import re
import docx


# 抓取word文件中的图片
def fetch_image(doc_path, desc_path):
    doc = docx.Document(doc_path)
    dict_rel = doc.part._rels  # rels其实是个目录
    for rel in dict_rel:
        rel = dict_rel[rel]
        print("rel", rel.target_ref)
        if "image" in rel.target_ref:
            # create_dir(desc_path)
            img_name = re.findall("/(.*)", rel.target_ref)[0]  # windos:/
            print("img_name", img_name)
            word_name = os.path.splitext(doc_path)[0]
            print("word_name", word_name)
            if os.sep in word_name:
                new_name = word_name.split('\\')[-1]
            else:
                new_name = word_name.split('/')[-1]
            img_name = f'{new_name}_{img_name}'
            with open(f'{desc_path}/{img_name}', "wb") as f:
                f.write(rel.target_part.blob)


def get_image():
    doc_name = r"C:\Users\14270\OneDrive\桌面\念奴娇_赤壁怀古.docx"
    desc_path = r"C:\Users\14270\OneDrive\桌面\新建文件夹"
    pwd = os.path.dirname(os.path.abspath(desc_path))
    print("[get_image]desc_path", desc_path)
    desc_path = os.path.join(pwd, desc_path)  # 目标路径
    fetch_image(doc_name, desc_path)


# 创建目录
def create_dir(desc_path):
    if not os.path.exists(desc_path):
        os.makedirs(desc_path)


if __name__ == '__main__':
    # create_doc()
    # fetch_doc()
    # update_doc()
    # create_doc_table()
    # fetch_doc_table()
    # modify_doc_table()
    get_image()