import os
import datetime
import re
exname = []
tt = 1
print('An so thu tu cua so do de chon')
print('Neu ko nhap gi thi mac dinh se chon file duoc tao gan day nhat voi dinh dang Xuat ngay-thang')
print("Cac file excel trong o dia la: ")
for path, dirname, filename in os.walk(os.getcwd()):
    for name in filename:
        if name.endswith(".xlsx"):
            print(str(tt)+' '+name)
            tt += 1
            exname.append(name)
tt = input()
if tt == "c":
    ngay = re.compile(r'(\w{1,12})(\s)(\d{1,2})(-)(\d{1,2})')
    today = datetime.date.today()
    daynow = int(today.day)
    monthnow = int(today.month)
    minday = 100
    minmonth = 100
    for name in exname:
        n = ngay.search(name)
        if n == None:
            continue
        else:
            day = int(n.group(3))
            month = int(n.group(5))
            if minday > daynow-day and minmonth > monthnow-month:
                minday = daynow-day
                minmonth = monthnow-month
                chosename = name
print(chosename)
exname.clear()
