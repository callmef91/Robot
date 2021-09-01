from os import environ

import pyautogui
from paddleocr import PaddleOCR, draw_ocr
'''
该文件中存放一些工具类
'''

'''
生成模型
'''
environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'  # 必要的
global ocr
ocr = PaddleOCR(use_gpu=False)  # 取得模型


'''
给类提供一个进行文字比对的函数textCompare()
传入两个str,第一个为输入,第二个为被比对
'''
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


