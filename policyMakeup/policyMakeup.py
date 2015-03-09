#coding:cp936
import cx_Oracle
import pymouse
import pykeyboard
import win32clipboard as cb
import time
import xml.etree.ElementTree as ET
import fileinput
import re

#utf8_encode = lambda x: x.encode("utf8") if type(x) == unicode else x
#utf8_decode = lambda x: x.decode("utf8") if type(x) == str else x
#cp_encode = lambda x: x.encode("GB2312") if type(x) == unicode else x
#cp_decode = lambda x: x.decode("GB2312") if type(x) == str else x

mouse = pymouse.PyMouse()
kb = pykeyboard.PyKeyboard()
conn = cx_Oracle.connect("pos/qwer1234!@10.24.57.50/gretms")


def copy(text):
    cb.OpenClipboard()
    cb.EmptyClipboard()
    cb.SetClipboardText(text,cb.CF_UNICODETEXT)
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
    time.sleep(0.2)
    #kb.press_key(kb.control_l_key)
    #kb.tap_key('v')
    #kb.release_key(kb.control_l_key)
    mouse.click(position[0],position[1],2)
    time.sleep(0.2)
    kb.tap_key('p')


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
        time.sleep(1)

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
        cur = conn.cursor()
        g.boxValue = init_values()
        line = change_xml_head(line)
        if line:
            print '-------------------------------------------------'
            f=open(g.tempxml,'w')
            f.write(line)
            f.close()
            tree = ET.parse(g.tempxml)
            root = tree.getroot()
            for info_tag in root:
                for child in info_tag:
                    g.tag[child.tag] = child.text
            g.boxValue['posid'] = g.tag['POS_ID']
            g.boxValue['operid'] =  g.tag['OPER_ID']
            g.boxValue['opername'] = g.tag['OPER_NAME']
            parse_cdata(g.tag['REC_LIST'])
            if g.boxValue['tbrname'] == '':
                g.boxValue['tbrname'] = '0'
            if g.boxValue['tbridno'] == '':
                g.boxValue['tbridno'] = '0'
            print "g.boxValue"
            print g.boxValue
            g.boxValue['voucherid'] = str(g.boxValue['voucherid'])
            sql1 = g.select_sql+"'"+g.boxValue['voucherid']+"'"
            sql2 = "select s_regposid from vc_voucherinfo where s_voucherid ="+"'"+g.boxValue['voucherid']+"'"
            cur.execute(sql1)
            for row in cur:
                result = row[0]
            print "���ݿ��в�ѯvoucherid��",g.boxValue['voucherid'],

            if result == g.boxValue['voucherid']:
                print result,'�Ѵ��ڣ��Թ�'
            else:
                cur.execute(sql2)
                for row in cur:
                    lock = row[0]
                    if lock:
                        print "��֤����"
                    else:
                        print result,
                        print 'û�У�����'
                        time.sleep(3)
                        iter_paste()
                        time.sleep(1)
                        mouse.click(g.success_btn[0],g.success_btn[1])
                        clear_box()
            


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

def clear_box():
    mouse.click(g.last_page[0],g.last_page[1])
    time.sleep(1)
    mouse.click(g.main_page[0],g.main_page[1])

class g(object):
    boxPos = {}
    boxValue = {}
    btnConfirmPos = [777,612] # ȷ����ť��λ��
    success_btn = [719,429]
    loghead_pat = re.compile(r".*\[com.pos.http.dao.impl.DB_SALEDAOImpl\]-<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>")
    xml_head = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
    logfile = r"d:\hehe.xml"
    tempxml = r".\temp.xml"
    tag = {}
    value_null = [684,421] # �ύ�󱨴���֤�Ų���Ϊ�ա�
    select_sql = "select s_voucherid from sp_bdsalelist where s_voucherid ="
    last_page = [115,359]
    main_page = [107,378]


g.boxPos = {
    # ���ֵ����ɱ���¼�������ҳ���е�λ��
    'voucherid':[535,164], # ��֤����
    'bdid':[1063,171], # ������
    'posid':[523,193], # �ն�ID
    'policyid':[1043,196], # ҵ�񷽰�����
    'saledate':[562,220], # ��������
    'startdate':[523,250], # ��������
    'enddate':[1029,245], # ��������
    'operid':[572,275], # ����Աid
    'opername':[1036,274], # ����Ա����
    'feeamt':[1057,297], # ���ѣ��֣�
    'tbrname':[543,327], # Ͷ��������
    'tbridno':[1024,357] # Ͷ����֤������
    }

#init_values()
#for k,v in g.boxValue.items():
#    g.boxValue[k]='3123'
#time.sleep(3)
#iter_paste()
parse_xml(g.logfile)
