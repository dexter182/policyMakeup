#encoding:cp936
import pymouse
import pykeyboard
import win32clipboard as cb
import time

mouse = pymouse.PyMouse()
kb = pykeyboard.PyKeyboard()

def copy(text):
    cb.OpenClipboard()
    cb.EmptyClipboard()
    cb.SetClipboardText(text)
    cb.CloseClipboard()

def paste():
    cb.OpenClipboard()
    data = cb.GetClipboardData()
    cb.CloseClipboard()
    return data

def click_commit():
    mouse.click(g.btnCommitPos[0],g.btnCommitPos[1])

def click_paste(position,text):
    mouse.click(position[0],position[1])
    copy(text)
    kb.press_key(kb.control_key)
    kb.tap_key('v')
    kb.release_key(kb.control_key)

def init_values():
    """ 返回一个值为空的输入项字典 """

    g.boxValue = g.boxPos.copy()
    for k in g.boxValue.keys():
        g.boxValue[k] = ''
    return g.boxValue

def iter_paste():
    """ 一个个进行点击-复制粘贴的主循环 """

    for k,v in g.boxPos.items():
        data = g.boxValue[k]
        click_paste(v,data)
        time.sleep(0.5)

    click_commit()

class g(object):
    boxPos = {}
    boxValue = {}
    btnCommitPos = []

g.boxPos = {
    # 该字典容纳保单录入各项在页面中的位置

    'Voucherid':[555,100] # 单证号

    }

