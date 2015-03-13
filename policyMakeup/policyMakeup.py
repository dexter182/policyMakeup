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
    """ 点击确定按钮 """
    mouse.click(g.config['btnConfirmPos'][0],g.config['btnConfirmPos'][1])

def click_paste(position,text):
    """ 一次选中文本框并粘贴的操作 """
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
    """ 加载配置文件 """
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
    """ 返回一个值为空的输入项字典 """

    g.boxValue = g.boxPos.copy()
    for k in g.boxValue.keys():
        g.boxValue[k] = ''
    return g.boxValue

def iter_paste():
    """ 补齐一张保单的动作 """

    for k,v in g.boxPos.items():
        data = g.boxValue[k]
        click_paste(v,data)
        time.sleep(0.5)

    click_confirm()

def change_xml_head(line):
    """ 把日志换成标准的xml头，另保单明细cdata里的第16项以后有时会有乱码，会导致xml解析失败，需要删掉"""
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
    """ 解析日志 """
    for line in fileinput.input(logfile):
        cur = conn.cursor()
        g.boxValue = init_values() # 初始化存放保单信息的字典
        line = change_xml_head(line) # 换xml头
        if line:
            print '-------------------------------------------------'
            f=open(g.config['tempxml'],'w') # 把当前行的内容写入临时temp.xml，以解析此单个xml文件
            f.write(line)
            f.close()
            tree = ET.parse(g.config['tempxml']) # 解析temp.xml
            root = tree.getroot()
            for info_tag in root:
                for child in info_tag:
                    g.tag[child.tag] = child.text
            g.boxValue['posid'] = g.tag['POS_ID']
            g.boxValue['operid'] =  g.tag['OPER_ID']
            g.boxValue['opername'] = g.tag['OPER_NAME']
            parse_cdata(g.tag['REC_LIST']) # 解析cdata中的内容，若投保人和身份证号为空，用'0'填充
            if g.boxValue['tbrname'] == '':
                g.boxValue['tbrname'] = '0'
            if g.boxValue['tbridno'] == '':
                g.boxValue['tbridno'] = '0'
            print "g.boxValue"
            print g.boxValue
            g.boxValue['voucherid'] = str(g.boxValue['voucherid'])
            sql1 = g.select_sql+"'"+g.boxValue['voucherid']+"'" # sql1用来验证当前单证号是否在sp_bdsalelist表中已有记录
            print 'sql1',sql1
            sql2 = "select s_regposid from vc_voucherinfo where s_voucherid ="+"'"+g.boxValue['voucherid']+"'" # sql2用来验证当前单证号是否已被申请
            cur.execute(sql1)
            for row in cur:
                result = row[0]
            print row  
            print 'test result,',result
            print "数据库中查询voucherid：",g.boxValue['voucherid'],


            if result == g.boxValue['voucherid']:
                print result,'保单已存在，略过'
            else:
                cur.execute(sql2)
                for row in cur:
                    lock = row[0]
                    if lock:
                        print "单证被锁"
                    else:
                        print result,
                        print '没有，补单'
                        time.sleep(int(g.config['init_time'])) # 解析完xml后，补单之前停留的时间
                        iter_paste() # 开始补单
                        time.sleep(0.5) # 所有数据粘贴完后，与点击“确认”按钮之前的时间间隔
                        mouse.click(g.config['success_btn'][0],g.config['success_btn'][1]) # 点击提示"补单成功"的对话框
                        clear_box(int(g.config['clear_time'])) # 切换页面以重置输入框
            


def parse_cdata(cdata):
    data = cdata.split('/')
    saledate = trans_date(data[0]) # 销售日期
    if len(data[2]) == 1:
        data[2] = '0' + data[2]
    policyid = data[1]+data[2] # 方案代码
    voucherid = data[3] # 单证号码
    startdate = trans_date(data[4]) # 保单起期
    enddate = trans_date(data[5]) # 保单止期
    feeamt = data[7] # 保费
    bdid = data[9] # 保单号
    tbrname = data[11] # 投保人姓名
    tbridno = data[15] # 投保人id号码
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
    """ 转化日期至标准的"yyyy-mm-dd 24hh:mi:ss"格式"""
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
    """ 更换页面，来清空文本框的内容 """
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
    'btnConfirmPos':[787,657], # 确定按钮的位置
    'success_btn':[691,421],
    'value_null':[674,423], # 提交后报错“单证号不能为空”
    'logfile':r"d:\hehe.xml",
    'tempxml':r".\temp.xml",
    'clear_time':1,
    'init_time':2
    }

g.boxPos = {
    # 该字典容纳保单录入各项在页面中的坐标
    'voucherid':[540,201], # 单证号码
    'bdid':[1044,208], # 保单号
    'posid':[619,231], # 终端ID
    'policyid':[1036,231], # 业务方案代码
    'saledate':[536,260], # 销售日期
    'startdate':[593,283], # 保险起期
    'enddate':[1048,281], # 保险终期
    'operid':[535,310], # 操作员id
    'opername':[1039,310], # 操作员姓名
    'feeamt':[1030,339], # 保费（分）
    'tbrname':[554,361], # 投保人姓名
    'tbridno':[1022,393] # 投保人证件号码
    }


load_config()
parse_xml(g.config['logfile'])
