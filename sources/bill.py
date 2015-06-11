#!/usr/bin/env python
# -*- coding: utf-8 -*-

import smtplib
from time import sleep
from time import ctime
from datetime import *
from email.mime.text import MIMEText
import email.MIMEMultipart
import csv

import os.path
import sys
import mimetypes
import ConfigParser
import string
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import codecs


#获取进入时间
dt = datetime.now()

#全局cache
msg_buff = []

#主要键值地图
keyword_map = dict()

#次要键值地图
other_map = []


def trim(string):
        string = string.strip().replace('\r\n','').replace('\n','')
        return string



class Worker:
    def __init__(self):
        self.conf = ConfigParser.ConfigParser()
        self.default_conf_file_path = 'setting.conf'

        if self.is_exist(self.default_conf_file_path):
            self.default_conf_file_path = 'setting.conf'
        elif self.is_exist('../example/sample.setting.conf'):
            print(u'在当前目录无法查找到设置文件，准备读取演示配置')
            self.default_conf_file_path = '../example/sample.setting.conf'
        else:
            print(u'无法找到配置文件')
            exit(12)

        self.conf.read(self.default_conf_file_path)


    def is_exist(self,path):
        if os.path.isfile(path):
            return True
        else:
            return False



class mail_Worker(Worker):
    def __init__(self):
        Worker.__init__(self)
        try:
            self.mail_host = self.conf.get('mail','host')
            self.mail_user = self.conf.get('mail','user')
            self.mail_pass = self.conf.get('mail','password')
            self.mail_postfix = self.conf.get('mail','postfix')
            self.mail_title_prefix = self.conf.get('mail','title_prefix')
            self.mail_bill_name = self.conf.get('mail','bill_name')

        except ConfigParser.NoSectionError:
            print(u'配置文件不正常')
            exit(13)
        except ConfigParser.NoOptionError:
            print(u'配置文件不正常')
            exit(14)

        try:
            self.s = smtplib.SMTP()
            self.s.connect(self.mail_host)
            self.s.login(self.mail_user,self.mail_pass)
        except Exception, e:
            print str(e)
            exit(18)


    def __del__(self):
        self.s.close()



    def send(self, to_list, content):
        '''
        to_list: send to somebody
        sub: subject
        content: content
        send_mail ("dblogs", "sub", "content")
        '''

        try:

            self.me=self.mail_user+"<"+self.mail_user+"@"+self.mail_postfix+">"
            self.msg = email.MIMEMultipart.MIMEMultipart()

            self.msg=MIMEText(content,_subtype='html',_charset="utf-8")
            self.msg['Subject'] = self.mail_title_prefix+dt.strftime('%Y/%m')+self.mail_bill_name
            self.msg['Subject'] = self.msg['Subject'].encode('gbk')

            self.msg['From'] = self.me
            self.msg['To'] = ";".join(to_list)
            self.msg['date'] = ctime()
            print('aaa')

            print(self.msg['Subject'])
        except UnicodeDecodeError:
            self.mail_title_prefix = self.conf.get('mail','title_prefix')
            self.mail_bill_name = self.conf.get('mail','bill_name')

            self.me=self.mail_user+"<"+self.mail_user+"@"+self.mail_postfix+">"
            self.msg = email.MIMEMultipart.MIMEMultipart()

            self.msg=MIMEText(content,_subtype='html',_charset="utf-8")
            self.msg['Subject'] = self.mail_title_prefix.decode('gbk')+dt.strftime('%Y/%m')+self.mail_bill_name.decode('gbk')

            self.msg['From'] = self.me
            self.msg['To'] = to_list
            self.msg['date'] = ctime()



        try:
            self.s.sendmail(self.me, to_list, self.msg.as_string())
            return True
        except Exception, e:
            print str(e)
            print(u'检查配置文件中邮件配置部份')
            exit(17)

class msg_Worker(Worker):
    def __init__(self):
        Worker.__init__(self)
        try:
            self.template_header = self.conf.get('template','header')
            self.template_footer = self.conf.get('template','footer')
            self.datafile = self.conf.get('data','datafile')

        except ConfigParser.NoSectionError:
            print(u'配置文件不正常')
            exit(13)
        except ConfigParser.NoOptionError:
            print(u'配置文件不正常')
            exit(14)

        #始初化模版
        if self.is_exist(self.template_header) & self.is_exist(self.template_footer):
            fp = open(self.template_header)
            self.msg_header = trim(fp.read())
            fp.close()
            fp = open(self.template_footer)
            self.msg_footer = trim(fp.read())
            fp.close()
        else:
            print(u'读取模版出问题了')
            exit(16)


    def csv_map(self):
        if (self.is_exist(self.datafile)):
            BLOCKSIZE = 8438
            with codecs.open(self.datafile, "r", "gbk") as sourceFile:
                with codecs.open('swap', "w", "utf-8") as targetFile:
                    while True:
                        try:
                            contents = sourceFile.read(BLOCKSIZE)
                            if not contents:
                                break
                            targetFile.write(contents)
                            self.datafile='swap'
                        except UnicodeDecodeError:
                            #print('文档编码错误,尝试使用别的编码打开文档')
                            pass

            reader = csv.reader(open(self.datafile))
        else:
            print(u'读取演示数据，并准备发送')
            reader = csv.reader(open("../example/sample.csv"))

        index_of_msg_key = 0
        for row in reader:
            if(len(row)==0):
                #print('Data has been buffer')
                break
            else:
                msg_buff.append(row)

        for key in msg_buff[0]:

            if trim(key) == '序号':
                keyword_map["k0"] = index_of_msg_key
            elif trim(key) == '姓名':
                keyword_map["name"] = index_of_msg_key
            elif trim(key) == '部门':
                keyword_map["k2"] = index_of_msg_key
            elif trim(key) == '出勤天数':
                keyword_map["k3"] = index_of_msg_key
            elif trim(key) == '基本工资':
                keyword_map["k4"] = index_of_msg_key
            elif trim(key) == '绩效工资':
                keyword_map["k5"] = index_of_msg_key
            elif trim(key) == '其他津贴':
                keyword_map["k6"] = index_of_msg_key
            elif trim(key) == '销售提成':
                keyword_map["k7"] = index_of_msg_key
            elif trim(key) == '扣请假款':
                keyword_map["k8"] = index_of_msg_key
            elif trim(key) == '应发工资':
                keyword_map["k9"] = index_of_msg_key
            elif trim(key) == '公司交社保':
                keyword_map["k10"] = index_of_msg_key
            elif trim(key) == '个人交社保':
                keyword_map["k11"] = index_of_msg_key
            elif trim(key) == '住房公积金（公司）':
                keyword_map["k12"] = index_of_msg_key
            elif trim(key) == '住房公积金（个人）':
                keyword_map["k13"] = index_of_msg_key
            elif trim(key) == '应税所得':
                keyword_map["k14"] = index_of_msg_key
            elif trim(key) == '税率':
                keyword_map["k15"] = index_of_msg_key
            elif trim(key) == '速算扣除':
                keyword_map["k16"] = index_of_msg_key
            elif trim(key) == '应缴个税':
                keyword_map["k17"] = index_of_msg_key
            elif trim(key) == '扣（补）其它':
                keyword_map["k18"] = index_of_msg_key
            elif trim(key) == '扣饭卡':
                keyword_map["k19"] = index_of_msg_key
            elif trim(key) == '扣班车费':
                keyword_map["k20"] = index_of_msg_key
            elif trim(key) == '邮箱号':
                keyword_map["email"] = index_of_msg_key
            elif trim(key) == '实发工资':
                keyword_map["actual_payment"] = index_of_msg_key
            else:
                other_map.append(index_of_msg_key)

            index_of_msg_key = index_of_msg_key + 1




    def msg_generator(self, msg):
        buff = '<div class="employee_info">'+msg[keyword_map["name"]]
        buff = buff + '<i class="department">'+msg[keyword_map["k2"]]+'</i></div>'
        buff = buff + '<div class="payment_date">['+dt.strftime('%Y/%m')  +'工资单]</div>'
        buff = buff + '<div class="normal" ><br>'
        buff = buff + '出勤天数:'+msg[keyword_map["k3"]]+'<br><br>'
        buff = buff + '基本工资:'+msg[keyword_map["k4"]]+'<br>'
        buff = buff + '绩效工资:'+msg[keyword_map["k5"]]+'<br>'
        buff = buff + '其他津贴:'+msg[keyword_map["k6"]]+'<br>'
        buff = buff + '销售提成:'+msg[keyword_map["k7"]]+'<br>'
        buff = buff + '扣请假款:'+msg[keyword_map["k8"]]+'<br><br>'
        buff = buff + '<div class="should_payment">应发工资:'+msg[keyword_map["k9"]]+'</div>'
        buff = buff + '公司交社保:'+msg[keyword_map["k10"]]+'<br>'
        buff = buff + '个人交社保:'+msg[keyword_map["k11"]]+'<br>'
        buff = buff + '住房公积金（公司）:'+msg[keyword_map["k12"]]+'<br>'
        buff = buff + '住房公积金（个人）:'+msg[keyword_map["k13"]]+'<br><br>'
        buff = buff + '应税所得:'+msg[keyword_map["k14"]]+'<br>'
        buff = buff + '税率:'+msg[keyword_map["k15"]]+'<br>'
        buff = buff + '速算扣除:'+msg[keyword_map["k16"]]+'<br>'
        buff = buff + '应缴个税:'+msg[keyword_map["k17"]]+'<br><br>'
        buff = buff + '扣（补）其它:'+msg[keyword_map["k18"]]+'<br>'
        buff = buff + '扣饭卡:'+msg[keyword_map["k19"]]+'<br>'
        buff = buff + '扣班车费:'+msg[keyword_map["k20"]]+'<br><br>'

        for key in other_map:
            buff = buff + msg_buff[0][key] + ':' + msg[key] + '<br>'

        buff = buff + '<div class="actual_payment">实发工资:'+msg[keyword_map["actual_payment"]]+'*<br></div>'

        return self.msg_header+buff+self.msg_footer






if __name__ == '__main__':
    print(u'开始读取数据')
    w1 = msg_Worker()
    w1.csv_map()
    print(u'读取数据结束')
    w2 = mail_Worker()
    print(u'开始发送邮件')
    msg_count = len(msg_buff);
    print(u'总共有'+str(msg_count-1)+u'封邮件需要发送')

    for i in range(1, msg_count):
        msg_body = w1.msg_generator(msg_buff[i])
        mailto_list = msg_buff[i][keyword_map["email"]]

        sleep(0.1)
        if w2.send(mailto_list,msg_body):
            try:
                name=msg_buff[i][keyword_map["name"]].decode('ascii').encode('utf-8')
                tips_info = u'第'+str(i)+u'邮件<'+name+ u':'+msg_buff[i][keyword_map["email"]]+u']成功!'
            except UnicodeDecodeError:
                name=msg_buff[i][keyword_map["name"]]
                tips_info = u'第'+str(i)+u'邮件['+name+ u':'+msg_buff[i][keyword_map["email"]]+u']成功!'

            print(tips_info)
        else:
            try:
                name=msg_buff[i][keyword_map["name"]].decode('ascii').encode('utf-8')
                tips_info = u'第'+str(i)+u'邮件<'+name+ u':'+msg_buff[i][keyword_map["email"]]+u']成功!'
            except UnicodeDecodeError:
                name=msg_buff[i][keyword_map["name"]]
                tips_info = u'第'+str(i)+u'邮件['+name+ u':'+msg_buff[i][keyword_map["email"]]+u']成功!'

            print(tips_info)

    print(u"全部工作完成")
    sleep(10)



