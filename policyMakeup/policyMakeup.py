#coding:cp936
import cx_Oracle
import pymouse
import pykeyboard
import win32clipboard as cb
import time
import xml.etree.ElementTree as ET
import fileinput
import re
import os

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
    """ ���ȷ����ť """
    mouse.click(g.config['btnConfirmPos'][0],g.config['btnConfirmPos'][1])

def click_paste(position,text):
    """ һ��ѡ���ı���ճ���Ĳ��� """
    mouse.click(position[0],position[1])
    copy(text)
    time.sleep(0.2)
    #kb.press_key(kb.control_l_key)
    #kb.tap_key('v')
    #kb.release_key(kb.control_l_key)
    mouse.click(position[0],position[1],2)
    time.sleep(0.2)
    kb.tap_key('p')

def load_config():
    """ ���������ļ� """
    if os.path.exists('config.ini'):
        for line in fileinput.input('config.ini'):
            if not re.search('=',line):
                continue
            data = line.split()
            key = data[0]
            val = data[2]
            if re.search(',',val):
                val = [ int(x) for x in val.split(',')]
            if key in g.config.keys():
                g.config[key] = val
            if key in g.boxPos.keys():
                g.boxPos[key] = val

def save_config():
    for k,v in g.boxPos.items()+g.config.items():
        if type(v) == list:
            v= ','.join(map(str,v))
        if type(v) == int:
            v = str(v)
        with open('config.ini','a') as f:
            f.write("%s = %s\n" %(k,v))

def init_values():
    """ ����һ��ֵΪ�յ��������ֵ� """

    g.boxValue = g.boxPos.copy()
    for k in g.boxValue.keys():
        g.boxValue[k] = ''
    return g.boxValue

def iter_paste():
    """ ����һ�ű����Ķ��� """

    for k,v in g.boxPos.items():
        data = g.boxValue[k]
        click_paste(v,data)
        time.sleep(0.5)

    click_confirm()

def change_xml_head(line):
    """ ����־���ɱ�׼��xmlͷ��������ϸcdata��ĵ�16���Ժ���ʱ�������룬�ᵼ��xml����ʧ�ܣ���Ҫɾ��"""
    m_loghead = g.loghead_pat.search(line)
    if m_loghead:
        original_head = m_loghead.group()
        line = line.replace(original_head,g.xml_head)
        m = re.search(g.legal_str,line)
        if m:
            line = m.group(1)+m.group(2)+'/////'+m.group(3)
        return line
    return 0



def parse_xml(logfile):
    """ ������־ """
    for line in fileinput.input(logfile):
        cur = conn.cursor()
        g.boxValue = init_values() # ��ʼ����ű�����Ϣ���ֵ�
        line = change_xml_head(line) # ��xmlͷ
        if line:
            print '-------------------------------------------------'
            f=open(g.config['tempxml'],'w') # �ѵ�ǰ�е�����д����ʱtemp.xml���Խ����˵���xml�ļ�
            f.write(line)
            f.close()
            tree = ET.parse(g.config['tempxml']) # ����temp.xml
            root = tree.getroot()
            for info_tag in root:
                for child in info_tag:
                    g.tag[child.tag] = child.text
            g.boxValue['posid'] = g.tag['POS_ID']
            g.boxValue['operid'] =  g.tag['OPER_ID']
            g.boxValue['opername'] = g.tag['OPER_NAME']
            parse_cdata(g.tag['REC_LIST']) # ����cdata�е����ݣ���Ͷ���˺����֤��Ϊ�գ���'0'���
            if g.boxValue['tbrname'] == '':
                g.boxValue['tbrname'] = '0'
            if g.boxValue['tbridno'] == '':
                g.boxValue['tbridno'] = '0'
            print "g.boxValue"
            print g.boxValue
            g.boxValue['voucherid'] = str(g.boxValue['voucherid'])
            sql1 = g.select_sql+"'"+g.boxValue['voucherid']+"'" # sql1������֤��ǰ��֤���Ƿ���sp_bdsalelist�������м�¼
            print 'sql1',sql1
            sql2 = "select s_regposid from vc_voucherinfo where s_voucherid ="+"'"+g.boxValue['voucherid']+"'" # sql2������֤��ǰ��֤���Ƿ��ѱ�����
            cur.execute(sql1)
            for row in cur:
                result = row[0]
            print row  
            print 'test result,',result
            print "���ݿ��в�ѯvoucherid��",g.boxValue['voucherid'],


            if result == g.boxValue['voucherid']:
                print result,'�����Ѵ��ڣ��Թ�'
            else:
                cur.execute(sql2)
                for row in cur:
                    lock = row[0]
                    if lock:
                        print "��֤����"
                    else:
                        print result,
                        print 'û�У�����'
                        time.sleep(int(g.config['init_time'])) # ������xml�󣬲���֮ǰͣ����ʱ��
                        iter_paste() # ��ʼ����
                        time.sleep(0.5) # ��������ճ�����������ȷ�ϡ���ť֮ǰ��ʱ����
                        mouse.click(g.config['success_btn'][0],g.config['success_btn'][1]) # �����ʾ"�����ɹ�"�ĶԻ���
                        clear_box(int(g.config['clear_time'])) # �л�ҳ�������������
            


def parse_cdata(cdata):
    data = cdata.split('/')
    saledate = trans_date(data[0]) # ��������
    if len(data[2]) == 1:
        data[2] = '0' + data[2]
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
    """ ת����������׼��"yyyy-mm-dd 24hh:mi:ss"��ʽ"""
    date_pat = re.compile("(....)(..)(..)(..)(..)(..)")
    match = date_pat.search(date_str)
    year = match.group(1)
    month = match.group(2)
    day = match.group(3)
    hour = match.group(4)
    min = match.group(5)
    sec = match.group(6)
    return year+'-'+month+'-'+day+' '+hour+':'+min+':'+sec

def clear_box(clear_time=1):
    """ ����ҳ�棬������ı�������� """
    mouse.click(g.config['last_page'][0],g.config['last_page'][1])
    time.sleep(1)
    mouse.click(g.config['main_page'][0],g.config['main_page'][1])
    time.sleep(clear_time)

class g(object):
    boxPos = {}
    boxValue = {}
    config = {}
    loghead_pat = re.compile(r".*\[com.pos.http.dao.impl.DB_SALEDAOImpl\]-<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>")
    xml_head = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
    tag = {}
    select_sql = "select s_voucherid from sp_bdsalelist where s_voucherid ="
    legal_str = r"(^.*<REC_LIST><!\[CDATA\[[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/[^/]*/)[^/]*(/[^/]*/[^/]*/[^/]*/[^/]*/).*(\]\].*$)"

g.config = {   
    'last_page':[93,402],
    'main_page':[118,419],
    'btnConfirmPos':[787,657], # ȷ����ť��λ��
    'success_btn':[691,421],
    'value_null':[674,423], # �ύ�󱨴���֤�Ų���Ϊ�ա�
    'logfile':r"d:\hehe.xml",
    'tempxml':r".\temp.xml",
    'clear_time':1,
    'init_time':2
    }

g.boxPos = {
    # ���ֵ����ɱ���¼�������ҳ���е�����
    'voucherid':[540,201], # ��֤����
    'bdid':[1044,208], # ������
    'posid':[619,231], # �ն�ID
    'policyid':[1036,231], # ҵ�񷽰�����
    'saledate':[536,260], # ��������
    'startdate':[593,283], # ��������
    'enddate':[1048,281], # ��������
    'operid':[535,310], # ����Աid
    'opername':[1039,310], # ����Ա����
    'feeamt':[1030,339], # ���ѣ��֣�
    'tbrname':[554,361], # Ͷ��������
    'tbridno':[1022,393] # Ͷ����֤������
    }


load_config()
parse_xml(g.config['logfile'])
