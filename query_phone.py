# -*- coding: utf-8 -*-
import phone
import xlwt
import xlrd
import time
def getPhone(filename):
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    original_list = table.col_values(0)
    phone_list = []
    for each in original_list:
        phone_list.append(int(each))
    return phone_list
def getInfo(phoneNum):
    info = phone.Phone().find(phoneNum)
    data = []    
    data.append(info['phone'])
    data.append(info['province'])
    data.append(info['city'])
    data.append(info['phone_type'])
    data.append(info['area_code'])
    data.append(info['zip_code'])
    return data

def saveDataToExcel(datalist,path):
    #标题栏背景色
    styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue; font: bold on;'); # 80% like
    #创建一个工作簿
    book=xlwt.Workbook(encoding='utf-8',style_compression=0)
    #创建一张表
    sheet=book.add_sheet('手机归属地查询',cell_overwrite_ok=True)
    #标题栏
    titleList=('手机号码','卡号归属省份','卡号归属城市','卡 类 型','区 号','邮 编')
    #设置第一列尺寸
    first_col = sheet.col(0)
    first_col.width=256*30
    #写入标题栏
    for i in range(0,6):
        sheet.write(0,i,titleList[i], styleBlueBkg)
    #写入Chat信息
    for i in range(0,len(datalist)):
        data=datalist[i]
        for j in range(0,len(data)):
            sheet.write(i+1,j,data[j])
    #保存文件到指定路径
    book.save(path)



if __name__ == "__main__":
    phone_list = getPhone('号码归属.xlsx')
    
    results = []  # 手机号码段信息列表
 
    for each_phone in phone_list:
        result = getInfo(each_phone)
        results.append(result)
 
    
    saveDataToExcel(results,'号码归属结果.xls')
