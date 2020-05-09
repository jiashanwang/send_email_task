import sqlalchemy, json,os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from tools import proxy_to_dict, write_xls, send_mail_annex
from sqlalchemy import text
from sqlalchemy import Float

engine = create_engine('mysql://root:hbwjshan@163.com@127.0.0.1/store_v2?charset=utf8', echo=True)

Session = sessionmaker(bind=engine)

if __name__ == "__main__":
    # sql = text("select * from goods", typemap={'guide_price': Float})
    # 姓名，电话，商品，金额，学校，楼栋，宿舍号，订单创建时间，备注
    sql_admin = text('select * from admin')
    sql_user = text(
        "select * from user where to_days(CURDATE())-to_days(createTime)<=1 and school_id=:school_id and dorm_id=:dorm_id")
    session_admin = Session()  # 每当需要与数据库对话时，都要实例化一个Session
    result_admin = session_admin.execute(sql_admin).fetchall()
    session_admin.close()
    result_list_admin = proxy_to_dict(result_admin)
    title_list = ["姓名", "电话", "商品", "金额", "学校", "楼栋", "宿舍号", "创建时间", "备注"]
    filed_list = ["userName", "phone", "productList", "totalPrice", "school", "dormBuild", "dormNumber", "createTime",
                  "mask"]
    for item in result_list_admin:
        school_id = item.get("school_id")
        dorm_id = item.get("dorm_id")
        user_name = item.get("name")
        emails = item.get("email")
        if ";" in emails:
            receiver_emails = emails.split(";")
        else:
            receiver_emails = []
            receiver_emails.append(emails)
        session_user = Session()
        result_user = session_user.execute(sql_user, {"school_id": school_id, "dorm_id": dorm_id}).fetchall()
        if len(result_user) == 0:
            # 当前楼栋没有用户订单  不发邮件
            send_mail_annex(user_name, "", receiver_emails)
        else:
            # 当前楼栋有用户订单
            result_list_user = proxy_to_dict(result_user)
            for user_data in result_list_user:
                product_list = json.loads(json.loads(user_data.get("productList")))
                product_infos = ""
                for product in product_list:
                    product_infos = product.get("goods_name") + " * " + str(product.get("goods_number")) + " ; "
                user_data['productList'] = product_infos
            xml_order_name = write_xls(title_list, filed_list, result_list_user, school_id, dorm_id)
            # 文件写入成功后，发邮件
            send_mail_annex(user_name, xml_order_name, receiver_emails, True)

    # 删除当前产生的excel 文件
    current_dir = os.path.dirname(__file__)
    print(current_dir)
    files = os.listdir(current_dir)
    for file in files:
        if file.endswith(".xls"):
            os.remove(os.path.join(current_dir, file))

