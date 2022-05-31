import sys
import you_get  # 导入you-get库
import openpyxl
import time

wb = openpyxl.load_workbook('收藏夹信息.xlsx')
print(type(wb))  # 结果: <class 'openpyxl.workbook.workbook.Workbook'>形式
# 获取所有表的表名
sheets_names = wb.sheetnames
print(sheets_names)  # 结果: ['表1', '表2']
sheet = wb[str(input('输入要下载视频的sheet表名称，然后按下回车'))]
time.sleep(2)  # 休眠2秒
i = 8
x = int(input('输入当前表格最后一行之后任意的行数，然后按下回车'))
while i < x:
    print(i)
    f1 = sheet['F'+str(i)]  # A1 表示A列中的第一行，这儿的列号采用的是从A开始的
    i = i + 10
    print(f1)  # 获取单元格中的内容
    content = f1.value
    print(content)  # 结果是: Rank
# 设置下载目录
    directory = 'E:\Bilibili'
# 设置要下载的视频地址
    url = 'https://www.bilibili.com/video/'+str(content)
    sys.argv = ['you-get', '-l', '-o', directory, url]
    try:
         you_get.main()
    except:
            pass