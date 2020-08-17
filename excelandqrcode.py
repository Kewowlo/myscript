# -*- coding: utf-8 -*- 
#!/usr/bin/env Python

from pyzbar.pyzbar import decode
from PIL import Image
import os,shutil
import xlrd
from urllib.parse import urlparse
from urllib.parse import parse_qs

if __name__ == '__main__':
    filepath = r"C:\Users\kewowlo\eclipse-workspace\testZxing\zxing\0729\rar\湾头\(湾头派出所第一警务区)二维码 (1)\xl\media"
    excelpath = r"C:\Users\kewowlo\eclipse-workspace\testZxing\zxing\0729\xlsx\湾头\(湾头派出所第一警务区)二维码 (1).xlsx"
    newpath = r"C:\Users\kewowlo\eclipse-workspace\testZxing\zxing\0729\二维码门牌1\湾头\(湾头派出所第一警务区)二维码(1)"

    
    data = xlrd.open_workbook(excelpath)
    table = data.sheets()[0]

    addr_list = table.col_values(0)[2:]
    type_list = table.col_values(1)[2:]
    note_list = table.col_values(5)[2:]
    # print (addr_list, type_list, note_list)
    if os.path.exists(newpath) == False:
        os.makedirs(newpath)
    for name, _type, note in zip(addr_list, type_list, note_list):
        print (name, _type, note)
        if note !="":
            notedir = os.path.join(newpath,note)
            print(notedir)
            
            if os.path.exists(notedir) == False:
                os.mkdir(notedir)
            temp = os.path.join(filepath, name+".png")
            if os.path.isfile(temp):
                shutil.move(temp, os.path.join(notedir, name+".png"))
                print(temp,"移动成功")
            else:
                print(temp,"移动失败")
        else:
            if _type == "":
                _type = "不明"
            typedir = os.path.join(newpath,_type)
            print(typedir)
            
            if os.path.exists(typedir) == False:
                os.mkdir(typedir)
            temp = os.path.join(filepath, name+".png")
            if os.path.isfile(temp):
                shutil.move(temp, os.path.join(typedir, name+".png"))
                print(temp,"移动成功")
            else:
                print(temp,"移动失败")
            
        