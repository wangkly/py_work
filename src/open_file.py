from tkinter import *
import tkinter.filedialog
from openpyxl import Workbook
from openpyxl import load_workbook
from main import entry

root = Tk()
root.geometry('500x300')

targetWb = None

def startwork(name):
    global targetWb
    print('name',name)
    entry(targetWb,name)
    pass

def placeBtns(item):
    button = Button(root,text=item,command=lambda : startwork(item))
    button.pack()
    return

def cmd():
    global targetWb
    filename = tkinter.filedialog.askopenfilename()
    targetWb = load_workbook(filename)
    sheets = targetWb.sheetnames
    label = Label(root,text="请选择要处理的表格")
    label.pack()
    for item in sheets:
        placeBtns(item)
    label2 = Label(root,text='文件路径：'+str(filename))
    label2.pack()

btn = Button(root,text="选择要处理的文件",command=cmd)
btn.place(x=100,y=100)
btn.pack()
root.mainloop()
