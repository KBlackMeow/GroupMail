import smtplib
import xlrd

from sendmailUI import *
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5 import QtCore, QtGui, QtWidgets

from email.mime.text import MIMEText
from email.header import Header

def test():
    print("testtt!")

def send():

    # 打开文件，读取学生名单
    data = xlrd.open_workbook('receivers.xlsx')

    table = data.sheet_by_index(0)

    ############################################################################################################
    # 发件人信息设置

    # 第三方SMTP服务，请在此处设置PKU邮箱的账号信息，只需填写user和password信息，user为发件邮箱，password为发件邮箱的登录密码
    host = "smtp.pku.edu.cn"  # 设置服务器
    user = "19012xxxxx@pku.edu.cn"  # 用户名
    password = "xxxxxxx"  # 口令

    # sender请设置为发件人pku邮箱(即为上面配置的user账号)
    sender = '1901xxxxx@pku.edu.cn'  # 发送方

    #############################################################################################################

    # 学生姓名列表
    names = table.col_values(table.row_values(0).index("姓名"))[1:]

    # 学生排名信息列表
    orders = [int(i) for i in table.col_values(table.row_values(0).index("排序"))[1:]]

    # 接收方 学生邮箱地址列表
    receivers = table.col_values(table.row_values(0).index("邮箱"))[1:]

    for i in range(len(names)):
        name = names[i]
        receiver = receivers[i]
        order = orders[i]
        # 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
        # 邮件内容请在这里定义，如果要修改总人数，也请在这里修改。注意，这里的总人数只能设置一次，并且是保持不变的，即名单中的学生收到的邮件
        # 中的总人数均为130，如果人数会变动（比如两个班或者年级），请分别整理出两个班级的信息Excel表格，并修改相应的总人数，分别运行程序
        message = MIMEText('''{}同学你好,\n祝贺你得奖。你的排序是 {}/130.'''.format(name, order), 'plain', 'utf-8')
        message['From'] = Header(sender)  # 发送者
        message['To'] = Header(receiver)  # 接受者

        subject = '【重要】校内门户奖励申请'
        message['subject'] = Header(subject, 'utf-8')

        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(host, 25)  # 25 为 SMTP 端口号
            smtpObj.login(user, password)
            smtpObj.sendmail(sender, receiver, message.as_string())
            print(name + "邮件发送成功")
        except smtplib.SMTPException as e:
            print(name + "Error: 无法发送邮件", e)

    print("Done!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    test()
    sys.exit(app.exec_())