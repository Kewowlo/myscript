# -*- coding: utf-8 -*- 
#!/usr/bin/env Python

from pyzbar.pyzbar import decode
from PIL import Image
import os
import xlrd
from urllib.parse import urlparse
from urllib.parse import parse_qs

if __name__ == '__main__':
    filepath = r"D:\陈建南\qqfile\门牌号二维码36个村\朝霞门牌二维码\门牌二维码\艾杨各庄村"
    # excelpath = r"C:\Users\kewowlo\Desktop\0729\xlsx\李典\(李典派出所第六警务区)二维码 (1).xlsx"
    # newpath = r"C:\Users\kewowlo\Desktop\0729\二维码门牌1"

    
    # data = xlrd.open_workbook(excelpath)
    # table = data.sheets()[0]

    # addr_list = table.col_values(0)[2:]
    # type_list = table.col_values(5)[2:]
    # print str(type_list).decode("unicode-escape")
    for root, dirs, files in os.walk(filepath, topdown=False):
        
        for name in files:

            if name.endswith('.png') == False:
                continue
            imagepath = os.path.join(root, name)
            print(imagepath)
            result = decode(Image.open(imagepath))
            
            ans = urlparse(result[0].data.decode("utf8"))
            if ans.fragment != 'null':
                print(ans)

            # print(parse_qs(ans[4])["DZMC"])
            # imagename = parse_qs(ans[4])["DZMC"][0] + ".png"
            # if os.path.exists(os.path.join(root, imagename)) == False:
                
            #      os.rename(imagepath,os.path.join(root, imagename))