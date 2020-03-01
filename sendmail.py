import time
import xlwt
from xlrd import open_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders



mail_host = "mail.chinare.local"
mail_from = "qiutx@chinare.com.cn"
mail_user = "qiutx"
mail_password="qtx&640421"


excelName = "f:\\atte_pg\\副本信息技术中心2019-03考勤统计表5.7版.xls"

bk = open_workbook(excelName)
workbook = xlwt.Workbook(encoding='utf-8')


def get_emailaddr(e_name):
	excelname1 = "f:\\atte_pg\\atte_datatest.xlsx"
	bk1 = open_workbook(excelName1)
	workbook1 = xlwt.Workbook(encoding='utf-8')
	sname1 = bk1.sheet_by_name("namelist")
	sname_nrows1 = sname1.nrows
	for i in range(1,sname_nrows1):
		if e_name == sname1.cell(i, 0).value:
			emailaddr=sname1.cell(i,5).value
			print("emailaddr",emailaddr)
			return emailaddr
		else:
			continue


def send_mail(mail_to, subject, body):
	msg = MIMEMultipart()
	msg['subject'] = subject
	msg['from'] = mail_from
	msg['to'] = mail_to
	msg['date'] = time.ctime()
	txt = MIMEText(body, 'plain', 'utf-8')
	msg.attach(txt)
	try:
		s = smtplib.SMTP()

		s.connect(mail_host)
		s.login(mail_user, mail_password)
		s.sendmail(mail_from, mail_to, msg.as_string())
		s.close()
		return True
	except:
		print("err")
		return False
sname = bk.sheet_by_name("kaoqin")

sname_nrows = sname.nrows

subject ="关于考勤报表结果沟通"

for i in range(3,sname_nrows):

	q_name=sname.cell(i,0).value
	print(q_name)
	print("DEBUG:sending mail to " ,q_name)
	e_str=sname.cell(i, 8).value
	body=q_name+"您好:"+"您这月得考勤状态如下"+" "+e_str+"请在2天内做出合理解释，例如忘记带卡，公出等,并通过邮件进行反馈。\n\n规划部综合处\n丘彤轩 "
	print("e_str",e_str)
	if e_str!="":
		mailto = get_emailaddr(q_name)
		print("mail=",mailto)
	else:
		continue

	if send_mail(mailto, subject, body):
			print("INFO:success to send mail to %s")
	else:
		print ("ERROR:fail to send mail to %s")
		continue
	print("done...python is great!")





print ("计算完成------------------->>>>>>>")



