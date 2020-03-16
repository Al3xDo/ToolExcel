import pyperclip
while True:
    print('1. c để copy đoạn văn bản')
    print('2. ` để thoát')
    x = input()
    if x == 'c':
        print('các đoạn cách nhau bằng space')
        letter = pyperclip.paste()
        test = letter.split()
        for i in range(0, len(test)):
            while True
            print('c để tiếp tục')
            pyperclip.copy(test[i])
