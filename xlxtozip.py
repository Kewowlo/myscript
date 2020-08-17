# coding=utf-8
'''
File Name：   readexcelimg
Author：      tim
Date：        2018/7/26 19:52
Description： 读取excel中的图片，打印图片路径
    先将excel转换成zip包，解压zip包，包下面有文件夹存放了图片，读取这个图片
'''

import os
import zipfile
from changeOffice import Change 
import win32com.client as win32
# 判断是否是文件和判断文件是否存在
def isfile_exist(file_path):
    if not os.path.isfile(file_path):
        print("It's not a file or no such file exist ! %s" % file_path)
        return False
    else:
        return True


# 修改指定目录下的文件类型名，将excel后缀名修改为.zip
def change_file_name(file_path, new_type='.zip'):
    if not isfile_exist(file_path):
        return ''

    extend = os.path.splitext(file_path)[1]  # 获取文件拓展名
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % file_path)
        return False

    file_name = os.path.basename(file_path)  # 获取文件名
    new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip

    dir_path = os.path.dirname(file_path)  # 获取文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新的文件路径
    if os.path.exists(new_path):
        os.remove(new_path)

    os.rename(file_path, new_path)  # 保存新文件，旧文件会替换掉

    return new_path  # 返回新的文件路径，压缩包


# 解压文件
def unzip_file(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False

    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # 获取文件所在目录
    for files in file_zip.namelist():
        file_zip.extract(files, os.path.join(zipfile_path, zipdir))  # 解压到指定文件目录

    file_zip.close()
    return True


# 读取解压后的文件夹，打印图片路径
def read_img(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False

    dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    pic_dir = 'xl' + os.sep + 'media'  # excel变成压缩包后，再解压，图片在media目录
    pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)

    file_list = os.listdir(pic_path)
    for file in file_list:
        filepath = os.path.join(pic_path, file)
        print(filepath)


# 组合各个函数
def compenent(excel_file_path):
    print(excel_file_path)
    zip_file_path = change_file_name(excel_file_path)
    if zip_file_path != '':
        if unzip_file(zip_file_path):
            read_img(zip_file_path)


def xls2xlsx(excel_file_path):
    
    for parent, dirnames, filenames in os.walk(excel_file_path):
        for fn in filenames:
            if fn.endswith('.xlsx') == False:
                continue
            filedir = os.path.join(parent, fn)
            newfile = os.path.join(parent, os.path.splitext(fn)[0]+".zip")
            os.rename(filedir,newfile)
            print(filedir)
            continue
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            # 后缀名的大小写不通配，需按实际修改：xls，或XLS
            wb.SaveAs(filedir.replace('xls', 'xlsx'), FileFormat=51)  # 我这里原文件是大写
            wb.Close()                                 
            excel.Application.Quit()

# main
if __name__ == '__main__':
    # compenent(u'C:/Users/kewowlo/Desktop/二维码 - 副本/五里庙(225003已核已发厂家0721)/(五里庙派出所第三警务区)二维码 (1).xlsx')
    root = r"D:\陈建南\qqfile\门牌号二维码36个村"
    xls2xlsx(root)