from tkinter import *
import tkinter.filedialog
from openpyxl import Workbook
from openpyxl import load_workbook
from main import entry

root = Tk()
root.title('徐二狗专用')
root.geometry('500x300')

targetWb = None
submitting = False

def startwork(name):
    global targetWb
    global submitting
    print('submitting',submitting)
    if submitting:
        print('正在执行')
        return
    submitting = True
    print('name',name)
    result  = entry(targetWb,name)
    if result == 1:
        labe2 = Label(root,text="处理成功")
        labe2.pack()
        submitting = False

def placeBtns(item):
    button = Button(root,text=item,command=lambda : startwork(item))
    button.pack()
    return

def cmd():
    global targetWb
    filename = tkinter.filedialog.askopenfilename()
    targetWb = load_workbook(filename)
    label2 = Label(root,text='文件路径：'+str(filename))
    label2.pack()
    sheets = targetWb.sheetnames
    label = Label(root,text="请选择要处理的表格")
    label.pack()
    for item in sheets:
        placeBtns(item)


btn = Button(root,text="选择要处理的文件",command=cmd)
btn.pack()
root.mainloop()
