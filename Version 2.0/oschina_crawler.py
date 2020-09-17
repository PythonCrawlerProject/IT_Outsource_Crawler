'''
开源中国抓取
'''

from datetime import datetime

import html2text
import requests
from openpyxl import Workbook

from config import oschina_headers, time_point
from utils import get_contact, get_ua, get_mysql_connection, create_table,  add_default_data
from sender import send_message


def get_id(url):
    try:
        response = requests.get(url, headers=oschina_headers.update({'User-Agent': get_ua()}))
        if response.status_code == 200:
            data = response.json()
            datas = data['data']['data']
            id_list = [(d['id'], d['type']) for d in datas]
            return id_list
        else:
            return 19, response.status_code
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def get_one_page(url):
    try:
        response = requests.get(url, headers=oschina_headers)
        if response.status_code == 200:
            data = response.json()
            data = data['data']
            description = data['prd']
            status = data['status']
            app = data['application']
            time_str = data['publishTime']
            tmp_str = data['statusLastTime']
            pub_time = datetime.strptime(time_str if time_str else tmp_str, "%Y-%m-%d %H:%M:%S")
            return [description, status, app, pub_time]
        else:
            return 33, response.status_code
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrderModel, WebsiteModel):
    print('开始爬取开源中国订单')
    sheet = wb.create_sheet('开源中国', 1)
    sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(2)
    for i in range(10, 0, -1):
        url = 'https://zb.oschina.net/project/contractor-browse-project-and-reward?applicationAreas=&moneyMinByYuan=&moneyMaxByYuan=&sortBy=30&currentTime=&pageSize=20&currentPage=%d' % i
        id_list = get_id(url)
        if isinstance(id_list, list):
            for oid, otype in id_list:
                if otype == 2:
                    url = 'https://zb.oschina.net/reward/detail?id=%d' % oid
                    link = 'https://zb.oschina.net/reward/detail.html?id=%s' % oid
                else:
                    url = 'https://zb.oschina.net/project/detail?id=%s' % oid
                    link = 'https://zb.oschina.net/project/detail.html?id=%s' % oid
                result = get_one_page(url)
                if isinstance(result, list):
                    publish_time = result[3]
                    if publish_time < time_point:
                        continue
                    desc = html2text.html2text(result[0]).strip()
                    is_valid = True if result[1] == 3 else False
                    contact = get_contact(desc)

                    oid = 'oc-{}'.format(oid//10)
                    order_query = session.query(OrderModel).filter_by(desc=desc, pub_time=publish_time).first()
                    if order_query:
                        is_valided = order_query.is_valid
                        order_query.is_valid = is_valid
                        if is_valid == True:
                            sheet.append([count, desc, link, publish_time, contact, ''])
                            count += 1
                            if is_valided == False:
                                order_query.is_delete = False
                        if is_valided == True and is_valid == False:
                            order_query.is_delete = True
                    else:
                        order = OrderModel(id=oid, desc=desc, link=link, contact=contact, category=result[2], pub_time=publish_time, is_valid=is_valid, is_delete=False if is_valid else True)
                        order.website = website
                        session.add(order)
                        if is_valid == True:
                            sheet.append([count, desc, link, publish_time, contact, ''])
                            count += 1
                elif isinstance(result, tuple):
                    message = '开源中国详情爬取第%d行出错：%s'  % (result[0], result[1])
                    print(message)
                    send_message(message)
            session.commit()
        elif isinstance(id_list, tuple):
            message = '开源中国爬取第%d行出错：%s' % (id_list[0], id_list[1])
            print(message)
            send_message(message)
    print('结束爬取开源中国订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
