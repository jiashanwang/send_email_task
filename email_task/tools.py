import datetime, xlwt,random,smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import parseaddr, formataddr


def proxy_to_dict(proxy_list):
    '''
     将数据库查询出来的 <<RowProxy实例对象>>列表进行统一的转换
     查询结果并不是数据库实例对象（因为字段不完整，是result 对象）
     :param proxy_list: 待转换的result 实例对象列表
     :return: 字典列表
    '''
    dict_list = []
    for item in proxy_list:
        dict_list.append(dict(zip(item.keys(), item)))
    return dict_list


def send_mail_annex(name,xml_order_name, receiver_emails, flag=None):
    '''
    发邮件（带附件）
    name: 邮件接受人的姓名
    xml_order_name:待读取的邮件文件名字
    receiver_emails：邮件接收列表
    flag: True（有订单） ，默认无订单数据
    '''
    if flag is not None:
        content = "<html><head></head><body><p> 您好，" + name + " : </p><p>附件为昨日订单数据</p></body></html>"
        msg = MIMEMultipart()
        msg.attach(MIMEText(content, 'html', 'utf-8'))
        with open('./' + xml_order_name, 'rb') as f:
            # 设置附件的MIME和文件名，这里是png类型:
            mime = MIMEBase('application', 'octet-stream', filename='Excel_test.xls')
            # 加上必要的头信息:
            mime.add_header('Content-Disposition', 'attachment', filename='Excel_test.xls')
            mime.add_header('Content-ID', '<0>')
            mime.add_header('X-Attachment-Id', '0')
            # 把附件的内容读进来:
            mime.set_payload(f.read())
            # 用Base64编码:
            encoders.encode_base64(mime)
            # 添加到MIMEMultipart:
            msg.attach(mime)
    else:
        content = "<html><head></head><body><p> 您好，" + name + " : </p><p>昨日无订单数据</p></body></html>"
        msg = MIMEText(content, 'html', 'utf-8')
    # 输入Email地址和口令:
    sender_email = "1283305468@qq.com"
    sender_password = "kapobpzhnkaehhdc"
    sender_name = "学长Store"
    # 输入SMTP服务器地址:
    smtp_server = 'smtp.qq.com'
    msg['From'] = formataddr([sender_name, sender_email])  # 发件人邮箱名称、账号
    msg['To'] = ",".join(receiver_emails)
    msg['Subject'] = "昨日订单数据"  # 邮件主题
    server = smtplib.SMTP(smtp_server, 25)  # SMTP协议默认端口是25
    server.login(sender_email, sender_password)  # 登陆邮件服务器
    server.sendmail(sender_email, receiver_emails, msg.as_string())
    server.quit()

def write_xls(title_list, filed_list, data_list,school_id,dorm_id):
    '''
    将数据写入 excel 表格
    '''
    i = 0
    work_book = xlwt.Workbook(encoding='utf-8')
    work_sheet = work_book.add_sheet("订单明细")
    # font_title 为标题的字体格式
    font_title = xlwt.Font()  # 创建字体样式
    font_title.name = '华文宋体'
    font_title.bold = True
    # 字体颜色
    font_title.colour_index = i
    # 字体大小，18为字号，20为衡量单位
    font_title.height = 20 * 18

    # font_body 为内容的字体央视
    font_body = xlwt.Font()
    font_body.name = '华文宋体'
    font_title.colour_index = i
    font_title.height = 20 * 12

    # 设置单元格对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01
    # 设置自动换行
    alignment.wrap = 1
    # 设置边框
    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    # 初始化样式 (标题样式)
    style_title = xlwt.XFStyle()
    style_title.font = font_title
    style_title.alignment = alignment
    style_title.borders = borders
    # 初始化样式 (内容样式)
    style_body = xlwt.XFStyle()
    style_body.font = font_body
    style_body.alignment = alignment
    style_body.borders = borders

    # 写入标题
    for index, item in enumerate(title_list):
        work_sheet.write(0, index, item, style_title)
    # 写入内容
    total_price = 0
    for index, item in enumerate(data_list, start=1):
        total_price +=item.get("totalPrice")
        for num, val in enumerate(filed_list):
            work_sheet.write(index, num, item.get(val), style_body)
    data_count = len(data_list) + 1

    work_sheet.write(data_count,2, "累计金额", style_title)
    work_sheet.write(data_count, 3,total_price , style_title)
    curr_date = str(datetime.datetime.now()).split(" ")[0] + "-" + str(school_id) + "-" + str(dorm_id) + "-order.xls"
    work_book.save(curr_date)
    return curr_date


def send_mail(mail_content, receiver_emails):
    '''
    mail_content:要发送的邮件正文，html 格式
    receiver_emails：收件人列表 ， list 格式
    '''

    msg = MIMEText(mail_content, 'html', 'utf-8')
    # 两个发件人，随机切换账号，避免同一个邮件发送过多被系统限制  password 是邮箱授权码
    senders = [{"email": "1283305468@qq.com", "password": "kapobpzhnkaehhdc"},
               {"email": "690865953@qq.com", "password": "pbrnwibduavtbege"}]
    current_sender = random.choice(senders)
    sender_email = current_sender.get("email")
    sender_password = current_sender.get("password")

    sender_name = "学长Store"
    # SMTP服务器地址:
    smtp_server = 'smtp.qq.com'
    # 判断收件人是一个还是多个
    if len(receiver_emails) == 1:
        receiver_email = receiver_emails[0]
    elif len(receiver_emails) > 1:
        receiver_email = ",".join(receiver_emails)

    msg['From'] = formataddr([sender_name, sender_email])  # 发件人邮箱名称、账号
    msg['To'] = receiver_email
    msg['Subject'] = "您有新的订单"  # 邮件主题
    server = smtplib.SMTP(smtp_server, 25)  # SMTP协议默认端口是25
    server.login(sender_email, sender_password)  # 登陆邮件服务器
    server.sendmail(sender_email, receiver_emails, msg.as_string())
    server.quit()

