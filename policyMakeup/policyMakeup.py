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

    click_commit()

class g(object):
    boxPos = {}
    boxValue = {}
    btnCommitPos = []

g.boxPos = {
    # ���ֵ����ɱ���¼�������ҳ���е�λ��

    'Voucherid':[555,100] # ��֤��

    }

