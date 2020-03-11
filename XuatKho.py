import openpyxl
import pyperclip
import os
import re
import datetime
# set up file
'''dirname = os.path.join(os.getcwd(), 'ToolExcel')
os.chdir(dirname)'''
wb = openpyxl.load_workbook('Xuat.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
# viet regex
kihieu = re.compile(r'''
^(\w\d{2,4}?) # ma so
(\s|,)
(\w{5}) # size
(\s|,)
(.*?) # mau
(\s|,)
(\d) # so luong
''', re.VERBOSE)
# thu nhap cookies
# giao dien
print('Lưu ý save excel trước khi dùng')
print('Lưu ý gõ ko dấu')
print('1.c để quét nội dung (can copy vao clipboard truoc)')
print('2.z để quét nội dung trong notepad')
print('3.` để thoát')
while True:
    # kiem tra dau vao
    x = input()
    if x == '`':
        break
    # chuc nang 1
    elif x == 'c':
        test = pyperclip.paste()
        maxcol = sheet.max_column
        test = pyperclip.paste()
        n = kihieu.search(test)
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
                        if mau == sheet.cell(row=2, column=c).value:
                            excol = c
                            c = maxcol+1
                            size = ' '
                        else:
                            c = c+1
                else:
                    c = c+1
            while r <= sheet.max_row:
                if maso == str(sheet.cell(row=r, column=1).value).strip():
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
# chuc nang 2
    elif x == 'z':
        test = open('ma.txt')
        test1 = test.readlines()
        for t in test1:
            t = t.strip()
            g = kihieu.search(t)
            if g == None:
                print('Loi')
            else:
                maso = g.group(1)
                size = g.group(3)
                mau = g.group(5)
                soluong = int(g.group(7))
                size = size[0:4]+str(size[4]).capitalize()
                c = 2
                i = 2
                r = 3
                while c <= maxcol:
                    if size == sheet.cell(row=1, column=c).value:
                        while c <= maxcol:
                            if mau == sheet.cell(row=2, column=c).value:
                                excol = c
                                c = maxcol+1
                                size = ' '
                            else:
                                c = c+1
                    else:
                        c = c+1
                while r <= sheet.max_row:
                    if maso == str(sheet.cell(row=r, column=1).value).strip():
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
        test.close()
# cuoi
fdate = datetime.date.today().strftime('%d/%m/%Y')
wb.save('Xuat'+fdate+'.xlsx')
print('Đã save với tên Xuat'+fdate+'.xlsx')
