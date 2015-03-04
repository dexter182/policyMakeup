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
    """ ����һ��ֵΪ�յ��������ֵ� """

    g.boxValue = g.boxPos.copy()
    for k in g.boxValue.keys():
        g.boxValue[k] = ''
    return g.boxValue

def iter_paste():
    """ һ�������е��-����ճ������ѭ�� """

    for k,v in g.boxPos.items():
        data = g.boxValue[k]
        click_paste(v,data)
        time.sleep(0.5)

    click_confirm()

class g(object):
    boxPos = {}
    boxValue = {}
    btnConfirmPos = [763,653] # ȷ����ť��λ��

g.boxPos = {
    # ���ֵ����ɱ���¼�������ҳ���е�λ��

    'voucherid':[575,204] # ��֤����
    'bdid':[1035,200] # ������
    'posid':[543,229] # �ն�ID
    'policyid':[1027,229] # ҵ�񷽰�����
    'saledate':[540,258] # ��������
    'startdate':[531,282] # ��������
    'enddate':[1027,285] # ��������
    'operid':[576,313] # ����Աid
    'opername':[1023,312] # ����Ա����
    'feeamt':[1017,338] # ���ѣ��֣�
    'tbrname':[534,365] # Ͷ��������
    'tbridno':[1022,391] # Ͷ����֤������
    }

