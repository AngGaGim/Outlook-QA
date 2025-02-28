import imaplib
import email
from email.header import decode_header

# https://learn.microsoft.com/en-us/answers/questions/1130148/outlook-imap-login-failed-using-python

# oauth
#https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

import imaplib


# 测试qq
# mail_user = '1364217409'
# mail_pass = 'yeupkahlmuvzjcgg'
# host = 'imap.qq.com'
# server = imaplib.IMAP4_SSL(host)
# print('服务器连接成功')
# server.login(mail_user,mail_pass)
# print('邮箱登陆成功')




# mail_user = '用户名@xxxxxxx.com.cn'
mail_user = 'wsjda5s' # 这里千万不可以有后缀！有后缀就如上面那行的话会报错
# mail_user = 'wsjda5s@outlook.com'
mail_pass = 'Hong12345'
host = 'imap-mail.outlook.com'# 服务器地址
# host = 'imap.outlook.com'
host = 'outlook.office365.com'

server = imaplib.IMAP4_SSL(host)
print('服务器连接成功')
server.login(mail_user,mail_pass)
print('邮箱登陆成功')
