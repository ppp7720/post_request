import requests, bs4, openpyxl
from tkinter import *
from tkinter import messagebox

def start():
    if len(num1.get()) != 13 or len(num2.get()) != 13 or int(num1.get()) > int(num2.get()):
        messagebox.showwarning("오류", "등기번호를 확인해주세요.(13자리)")
    else:
        go()
        messagebox.showwarning("완료", "조회가 완료되었습니다. '배송조회결과.xlsx' 파일을 확인하세요.")

def go():
    url = 'http://openapi.epost.go.kr/trace/retrieveLongitudinalCombinedService/retrieveLongitudinalCombinedService/getLongitudinalCombinedList'
    serviceKey = 'key'
    rgist1 = int(num1.get())
    rgist2 = int(num2.get())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["regiNo", "senderName", "senderData", "receiveName", "date", "time", "statue", "location"])

    for rgist in range(rgist1, rgist2 + 1):
        response = requests.get(url, params={'ServiceKey': serviceKey, 'rgist': rgist}).text
        xml = bs4.BeautifulSoup(response, 'xml').trackInfo
        if xml: ws.append(list(xml.strings)[:4] + list(xml.contents[-1].strings)[1:])
        else: ws.append([str(rgist), "조회결과가 없습니다."])

    wb.save("배송조회결과.xlsx")
    

root = Tk()
root.geometry('240x80+100+200')

root.title('등기배송조회')

num1 = StringVar()
num2 = StringVar()

Label(root, text="조회시작번호").grid(row=0, column=0)
Label(root, text="마지막번호").grid(row=1, column=0)

Entry(root, textvariable=num1).grid(row=0, column=1)
Entry(root, textvariable=num2).grid(row=1, column=1)

Button(root, text="엑셀출력", command=start).grid(row=3, column=1)

root.mainloop()