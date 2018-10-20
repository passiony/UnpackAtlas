#coding=utf-8

import os
import xlrd
from PIL import Image
from PIL import ImageEnhance

#import pyautogui
#import re

'''
#把当前目录下的10*10张jpeg格式图片拼接成一张大图片

#图片压缩后的大小
cell_width = 200
cell_height = 300

space_width=10;
space_height=10;

#每行每列显示图片数量
col_max = 3
row_max = 3

#参数初始化
all_path = []
num = 0
pic_max=col_max*row_max

dirName = os.getcwd()

for root, dirs, files in os.walk(dirName):
        for file in files:
            if "jpg" in file or "jpeg" in file or "png" in file:
                    all_path.append(os.path.join(root, file))

#初始化目标
toImage = Image.new('RGBA',((cell_width+space_width)*col_max,(cell_height+space_height)*row_max))

#排列图片
for index in range(len(all_path)):
    pic = Image.open(all_path[index])
    temp = pic.resize((cell_width, cell_height))

    pos = (int(index%col_max*(cell_width+space_width)),int(index/col_max*(cell_height+space_height)))
    print("第" + str(num) + "存放位置" + str(pos))
    toImage.paste(temp, pos)

print("---------------------")
print("merged 生成完毕！！！")
toImage.save('merged.png')

# 亮度增强
origin = ImageEnhance.Brightness(toImage)
brightness = 1.25
result = origin.enhance(brightness)
result.show()

'''

dirName = os.getcwd()
for root, dirs, files in os.walk(dirName):
        for file in files:
            if "xlsx" in file:
                xlsx_path = os.path.join(root, file)
                break

workbook = xlrd.open_workbook(xlsx_path)#打开excel文件
# sheet_names= workbook.sheet_names()
# for sheet_name in sheet_names: # 获取sheets
sheet = workbook.sheet_by_name("Sheet1")
nrows = sheet.nrows# 拿到总共行数
colnames = sheet.row_values(0)# 第一行数据

#输出信息
for i in range(1, nrows): #也就是从Excel第二行开始，第一行表头不算
    row = sheet.row_values(i)# 某一行数据
    if row:
        for j in range(len(colnames)):# 表头与数据对应
            print(u"%s::%s"%(colnames[j],row[j]));
