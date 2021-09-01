'''
引用区:应该尽量减少对整个包的引用，以缩小文件的大小

'''
import time
from tkinter import Frame, Tk, StringVar, Label, Entry, Button
from time import sleep
# import requests
# import base64
import pyautogui
print(pyautogui.position())

from os import environ
from paddleocr import PaddleOCR, draw_ocr

'''
   |￣￣￣￣￣￣￣￣￣￣|
   |   不要摸鱼哦！    | 
   |＿＿＿＿＿＿＿＿＿＿|   
    (\__/) ||
    (•ㅅ•) ||
    /  　 づ
'''

'''
给类提供一个进行文字比对的函数textCompare()
传入两个str,第一个为输入,第二个为被比对
'''
global orgList
orgList = ['ORG','90000.100000','90000.28000','90000.10000','90000.A2100','11000','23000']

def textCompare(inputString, textString):
    if inputString==textString:
        return True
    if len(textString)-len(inputString)>3:
        return False
    for i in range(len(textString)-len(inputString)+1):
        while((textString[i] ==inputString[0])&(len(textString[i:])==len(inputString))):
            if textString[i:]==inputString:
                return i
            else:
                return False
    return False

'''
多次滚轮滑动
'''
def rollNum(n):
    for i in range(n):
        pyautogui.scroll(-1)

'''
截图并识别
'''
def recongize(frame):
    imgPath = '.\pic\cut.png'
    im = pyautogui.screenshot(region=frame)  # 截图右边部分
    im.save(imgPath)
    return (ocr.ocr(imgPath))

def monthTrans(mon):
    monDict = {'YTD':'YID','1':'Jan','2':'Feb','3':'Mar','4':'Apr',
            '5':'May', '6':'Jn','6A':'JunA','7':'E','8':'Aug',
            '9':'Sep','10':'oct','11':'Nov','12':'Dec','12A':'DecA'}
    return monDict[mon]

'''
类RobotWindow:继承tkinter,提供界面来作为输入
'''
class RobotWindow(Frame):
    def __init__(self):
        global root
        root = Tk()
        root.title('C1插件下载机器人')
        root.geometry('200x300+1000+420')
        # 定义几个变量来接收
        self.year = StringVar()
        self.month = StringVar()
        self.organ = StringVar()
        self.mess = StringVar()

        # 加载的函数
        self.startWindow()
        # 启动的函数
        root.mainloop()

    def startWindow(self):
        '''
        这个函数给主界面添加上几个简单的Dom
        '''
        yearLabel = Label(root, text="年:").place(x=10, y=10)
        yearEntry = Entry(root, textvariable=self.year).place(x=50, y=10)

        monthLabel = Label(root, text="月:").place(x=10, y=45)
        monthEntry = Entry(root, textvariable=self.month).place(x=50, y=45)

        organLabel = Label(root, text="组织:").place(x=10, y=80)
        organEntry = Entry(root, textvariable=self.organ).place(x=50, y=80)

        cookButton = Button(root, text='Cook!', command=self.cook).place(x=10, y=115)
        quitButton = Button(root, text='古德拜', command=root.quit).place(x=80, y=115)

        messageLabel = Label(root, textvariable=self.mess).place(x=10, y=180)

    def cook(self):
        self.subCook(orgList[int(self.organ.get())])

    def subCook(self,mark):
        '''
        第一步：去读取输入的内容，并且启动脚本的程序
        '''
        # 通过图例找到插件按钮
        point = pyautogui.locateCenterOnScreen('.\pic\step1.png')
        if point == None:
            self.mess.set('请在Excel上操作')
            return  # 在此就结束函数了
        else:

            pyautogui.click(x=point[0], y=point[1] + 10, clicks=2)

            pyautogui.moveRel(0, 55, 0.7)
            pyautogui.click()


        '''
        第二步：出来中间那个框框了
        part.1：左边选‘年’，‘期间’，‘组织’
        part.2：右边根据输入来
        '''
        # part.1
        size = pyautogui.size()
        environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'  # 必要的
        global ocr
        ocr = PaddleOCR(use_gpu=False)  # 取得模型
        pyautogui.moveTo(size[0] / 2 - 80, size[1] / 2)  # 鼠标过去

        frame = (370,180,580,600)
        for i in range(4):
            sleep(1)  # 稍微等一下，不然界面没出来
            result = recongize(frame)
            # 对当前页面进行识别
            for text in result:
                for i in ['年', '期间', '组织']:
                    if textCompare(i, text[1][0]) !=False:
                        pyautogui.click(395,(text[0][0][1] + text[0][3][1]) / 2 + frame[1])
                    else:
                        pass
            rollNum(4)

        # part.2
        frame = (950, 180, 580, 600)
        result = recongize(frame)
        for text in reversed(result):
            pyautogui.click(text[0][0][0] + 960, (text[0][0][1] + text[0][3][1]) / 2 + 180)


        # 已经全部点开了
        for i in [self.year.get(), monthTrans(self.month.get()), mark]:
            print(i)
            tt = True  # 设计个标记
            while (tt):  # 使用一个循环一直向下找
                # 一个页面的识别
                sleep(1)
                result=recongize(frame)
                # 对一个页面的文字判别，找到了直接进去下一个属性，不翻页
                # 没找到才翻页
                for text in result:
                    print(text[1][0])
                    if textCompare(i, text[1][0]) !=False:
                        tt = False
                        if str(i) ==mark:
                            pyautogui.moveTo(1007,(text[0][0][1] + text[0][3][1]) / 2 + frame[1])
                        else:
                            pyautogui.click(1007,(text[0][0][1] + text[0][3][1]) / 2 + frame[1])
                        break
                if tt == True:
                    rollNum(7)

        x,y = pyautogui.position()
        x = 1004
        for i in range(6):
            pyautogui.click(x,y)
            pyautogui.click(x,y+35)
            rollNum(1)



        pyautogui.click(1483, 801)

        # part.3
        pyautogui.click(387,275,1)
        pyautogui.click(416, 312, 1)
        pyautogui.click(447, 493, 1)
        pyautogui.click(500, 628, 1)
        pyautogui.click(1409, 803, 1)
        pyautogui.click(1447, 799, 1)

        print(time.asctime( time.localtime(time.time()) ))




'''
用GUI将程序封装起来
'''
if __name__ == "__main__":
    robotWindow = RobotWindow()
