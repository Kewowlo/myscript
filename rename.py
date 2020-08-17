# -*- coding: utf-8 -*- 
import os

if __name__ == "__main__":
    
    filepath = r"C:\Users\kewowlo\eclipse-workspace\testZxing\zxing\0729\二维码门牌1"
    ans = 0
    for root, dirs, files in os.walk(filepath, topdown=False):
        _,a = os.path.split(root)
        print( a )
        if a.find("不") != -1:
            continue
        for name in files:
            
            # imagepath = os.path.join(root, name)

            # os.rename(imagepath,os.path.join(root, str(ans)+'.png'))
            ans = ans + 1

    print(ans)