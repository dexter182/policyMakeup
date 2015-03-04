#coding:cp936
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

def click_confirm():
    mouse.click(g.btnConfirmPos[0],g.btnConfirmPos[1])

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

    click_confirm()

class g(object):
    boxPos = {}
    boxValue = {}
    btnConfirmPos = [763,653] # 确定按钮的位置

g.boxPos = {
    # 该字典容纳保单录入各项在页面中的位置

    'voucherid':[575,204] # 单证号码
    'bdid':[1035,200] # 保单号
    'posid':[543,229] # 终端ID
    'policyid':[1027,229] # 业务方案代码
    'saledate':[540,258] # 销售日期
    'startdate':[531,282] # 保险起期
    'enddate':[1027,285] # 保险终期
    'operid':[576,313] # 操作员id
    'opername':[1023,312] # 操作员姓名
    'feeamt':[1017,338] # 保费（分）
    'tbrname':[534,365] # 投保人姓名
    'tbridno':[1022,391] # 投保人证件号码
    }

