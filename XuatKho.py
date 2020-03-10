import openpyxl
import pyperclip
import os
import re
dirname = os.path.join(os.getcwd(), 'ToolExcel')
os.chdir(dirname)
wb = openpyxl.load_workbook('Xuat.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
kihieu = re.compile(r'''
^(\w\d{2,4}?) # ma so
(\s|,)
(\w{5}) # size
(\s|,)
(.*?) # mau
(\s|,)
(\d) # so luong
''', re.VERBOSE)
print('Lưu ý save excel trước khi dùng')
print('1.c để quét nội dung (can coppy vao clipboard truoc)')
print('2.` để thoát')
while True:
    x = input()
    if x == '`':
        break
    elif x == 'c':
        test = pyperclip.paste()
    column = sheet['A']
    maxcol = len(column)
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
        c = 1
        i = 2
        r = 3
        while c <= maxcol:
            if size == sheet.cell(row=1, column=c).value:
                while i <= maxcol:
                    if mau == sheet.cell(row=2, column=i).value:
                        excol = i
                        i = maxcol+1
                        size = ' '
                        c = maxcol+1
                    else:
                        i = i+1
            else:
                c = c+1
        while r <= sheet.max_row:
            if maso == sheet.cell(row=r, column=1).value:
                if sheet.cell(row=r, column=excol).value == None:
                    print('ô ko tồn tại')
                else:
                    num = int(sheet.cell(row=r, column=excol).value)-soluong
                    sheet.cell(row=r, column=excol).value = num
                    r = sheet.max_row+1
                    print('Đã chỉnh xong')
            else:
                r = r+1
wb.save('Xuat1.xlsx')
print('Đã save với tên Xuát.xlsx')
