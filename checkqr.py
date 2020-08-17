# -*- coding: utf-8 -*- 
import os,shutil,math
import xlrd
import pandas 
def dump_excel(excel_path):#去掉表格里重复的
    
    for root, dirs, files in os.walk(excel_path, topdown=False):
        for name in files:
            if name.endswith('.xls') == False:
                continue
            filepath = os.path.join(root, name)
            table = pandas.read_excel(filepath,index_col = None, header = 0)
            # print(table)
            #print(table.columns.values)
            table.drop_duplicates(subset=['标准地址'], inplace=True)
            # print(table)
            
            # for i in table.columns.values:
            #     print(i)
            for i in ["数量","大门牌","小门牌","楼栋牌","单元牌","楼层牌","户室牌","含二维码"]:
                sum = 0
                for j in table[i][1:] :
                    if math.isnan(j):
                        continue
                    sum += j
                table[i][0] = sum
            #print(table)
            
            newfilepath = os.path.join(root,"新"+name)
            pandas.DataFrame(table).to_excel(newfilepath, sheet_name='Sheet1', index=False, header=True)

def get_duplication(excel_path):#得到表格里重复的
    newtable = []
    for root, dirs, files in os.walk(excel_path, topdown=False):
        for name in files:
            if name.endswith('.xls') == False or name[0] == '新':
                continue
            filepath = os.path.join(root, name)
            table = pandas.read_excel(filepath,index_col = None, header = 0)
            table['is_duplicated'] = table.duplicated(keep = 'first')
            newtable.append(table.loc[table['is_duplicated'] == True])
            
    writer = pandas.ExcelWriter(r'C:\Users\kewowlo\Desktop\1.xls')
    pandas.concat(newtable).to_excel(writer,'Sheet1',index=False)
    
    writer.save()

def check_excel(excel_path):#检查
    ans = {"数量":0,"大门牌":0,"小门牌":0,"楼栋牌":0,"单元牌":0,"楼层牌":0,"户室牌":0,"含二维码":0}
    for root, dirs, files in os.walk(excel_path, topdown=False):
        newfile = [] #二维码文件
        temp_excel = ""
        for name in files:
            if name.endswith('.xls') and name[0] != "新":
                #print(name)
                temp_excel = name
            if name.endswith('.jpg'):
                newfile.append(os.path.splitext(name)[0][3:])
        if temp_excel == "":
            continue
        data = xlrd.open_workbook(os.path.join(root, temp_excel))
        table = data.sheets()[0]
        namelist = table.col_values(0)[2:]
        flaglist = table.col_values(8)[2:]
        ans['数量'] += int(table.col_values(1)[1])
        ans['大门牌'] += int(table.col_values(2)[1])
        ans['小门牌'] += int(table.col_values(3)[1])
        ans['楼栋牌'] += int(table.col_values(4)[1])
        ans['单元牌'] += int(table.col_values(5)[1])
        ans['楼层牌'] += int(table.col_values(6)[1])
        ans['户室牌'] += int(table.col_values(7)[1])
        ans['含二维码'] += int(table.col_values(8)[1])
        #print(newfile)
        for name, flag in  zip(namelist, flaglist):
            if name in newfile and int(flag) == 1: #在二维码
                pass
            elif flag == "" : # 不是二维码
                pass
            else :#
                print(os.path.join(root,name))
    return ans
                
if __name__ == "__main__":
    excel_path = r"D:\陈建南\qqfile\江岸二维码"
    get_duplication(excel_path)
    
    

    