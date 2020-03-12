import openpyxl
import os
import re
import pyperclip
# ki hieu
kihieu = re.compile(r'''
^(\w\d{2,4}) # ma so
(\s|,)
(\w{5,7}) # size
(\s|,)
(.*?) # mau
(\s|,)
(\d{1,3}) # so luong
''', re.VERBOSE)
# Setup file
wb = openpyxl.load_workbook('Xuat.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
maxcol = sheet.max_column
# giao dien
print('1.c để thêm thủ công')
print('2.z để thêm hàng loat qua notepad')
print('3.` để thoát')
# them hang loat
while True:
    x = input()
    if x == 'c':
        maxcol = sheet.max_column
        test = open('ma.txt')
        test1 = test.readlines()
        for t in test1:
            t = t.strip()
            n = kihieu.search(t)
            if n == None:
                print('Loi')
            else:
                maso = n.group(1)
                size = n.group(3)
                mau = n.group(5)
                soluong = int(n.group(7))
                size = size[0:4]+str(size[4]).capitalize()
                c = 2
                i = 2
                r = 3
                while c <= maxcol:
                    if size == sheet.cell(row=1, column=c).value:
                        while c <= maxcol:
                            if mau == str(sheet.cell(row=2, column=c).value).strip():
                                excol = c
                                c = maxcol+1
                                size = ' '
                            else:
                                c = c+1
                    else:
                        c = c+1
                while r <= sheet.max_row:
                    if maso == str(sheet.cell(row=r, column=1).value).strip() or maso == sheet.cell(row=r, column=1).value:
                        if sheet.cell(row=r, column=excol).value == None:
                            print('ô ko tồn tại')
                        else:
                            num = int(sheet.cell(
                                row=r, column=excol).value)-soluong
                            sheet.cell(row=r, column=excol).value = num
                            r = sheet.max_row+1
                            print('Đã chỉnh xong')
                    else:
                        r = r+1
