from tkinter import *
import tkinter.filedialog
from openpyxl import Workbook
from openpyxl import load_workbook
from main import entry

root = Tk()
root.title('徐二狗专用')
root.geometry('500x300')

targetWb = None
destnation = None
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
    result  = entry(targetWb,destnation,name)
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
    label = Label(root,text="第三步:请选择要处理的表格")
    label.pack()
    for item in sheets:
        placeBtns(item)

def cmd2():
    global destnation
    destnation = tkinter.filedialog.askopenfilename()
    label3 = Label(root,text='模板文件路径：'+str(destnation))
    label3.pack()

btn2 = Button(root,text="第一步：选择要生成文件的模板",command=cmd2)
btn2.pack()

btn = Button(root,text="第二步：选择辅助余额表",command=cmd)
btn.pack()

root.mainloop()
