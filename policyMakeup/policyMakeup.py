import pymouse
import pykeyboard
import win32clipboard as cb
import time
import xml.etree.ElementTree as ET
import fileinput
import re

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
    kb.press_key(kb.control_l_key)
    kb.tap_key('v')
    kb.release_key(kb.control_l_key)

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

def change_xml_head(line):
    m_loghead = g.loghead_pat.search(line)
    if m_loghead:
        original_head = m_loghead.group()
        line = line.replace(original_head,g.xml_head)
        return line
    return 0


def parse_xml(logfile):
    """ ������־ """
    for line in fileinput.input(logfile):
        line = change_xml_head(line)
        if line:
            tree = ET.fromstring(line)
            root = tree.getroot()
            for info_tag in root:
                for child in info_tag:
                    g.tag[child.tag] = child.text
            g.boxValue['posid'] = g.tag['POS_ID']
            g.boxValue['operid'] =  g.tag['OPER_ID']
            g.boxValue['opername'] = g.tag['OPER_NAME']
            parse_cdate(g.tag['REC_LIST'])
        init_values()
        time.sleep(3)
        iter_paste()

def parse_cdata(cdata):
    data = cdata.split('/')
    saledate = trans_date(data[0]) # ��������
    policyid = data[1]+data[2] # ��������
    voucherid = data[3] # ��֤����
    startdate = trans_date(data[4]) # ��������
    enddate = trans_date(data[5]) # ����ֹ��
    feeamt = data[7] # ����
    bdid = data[9] # ������
    tbrname = data[11] # Ͷ��������
    tbridno = data[15] # Ͷ����id����
    g.boxValue['voucherid'] = voucherid
    g.boxValue['bdid'] = bdid
    g.boxValue['policyid'] = policyid
    g.boxValue['saledate'] = saledate
    g.boxValue['startdate'] = startdate
    g.boxValue['enddate'] = enddate
    g.boxValue['feeamt'] = feeamt
    g.boxValue['tbrname'] = tbrname
    g.boxValue['tbridno'] = tbridno
    
def trans_date(date_str):
    date_pat = re.compile("(....)(..)(..)(..)(..)(..)")
    match = date_pat.search(date_str)
    year = match.group(1)
    month = match.group(2)
    day = match.group(3)
    hour = match.group(4)
    min = match.group(5)
    sec = match.group(6)
    return year+'-'+month+'-'+day+' '+hour+':'+min+':'+sec

class g(object):
    boxPos = {}
    boxValue = {}
    btnConfirmPos = [763,653] # ȷ����ť��λ��
    loghead_pat = re.compile(r".*\[com.pos.http.dao.impl.DB_SALEDAOImpl\]-<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>")
    xml_head = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
    logfile = r"d:\tcweb0302a.log"
    tag = {}

g.boxPos = {
    # ���ֵ����ɱ���¼�������ҳ���е�λ��
    'voucherid':[575,204], # ��֤����
    'bdid':[1035,200], # ������
    'posid':[543,229], # �ն�ID
    'policyid':[1027,229], # ҵ�񷽰�����
    'saledate':[540,258], # ��������
    'startdate':[531,282], # ��������
    'enddate':[1027,285], # ��������
    'operid':[576,313], # ����Աid
    'opername':[1023,312], # ����Ա����
    'feeamt':[1017,338], # ���ѣ��֣�
    'tbrname':[534,365], # Ͷ��������
    'tbridno':[1022,391] # Ͷ����֤������
    }

#init_values()
#for k,v in g.boxValue.items():
#    g.boxValue[k]='3123'
#time.sleep(3)
#iter_paste()
parse_xml()