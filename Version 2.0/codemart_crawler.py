'''
码市订单抓取
'''

import time
from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from config import codemart_headers, time_point
from utils import get_contact, get_mysql_connection, create_table, add_default_data
from sender import send_message


def get_one_page(url):
    result_list = []
    try:
        response = requests.get(url, headers=codemart_headers)
        if response.status_code == 200:
            data = response.json()
            rewards = data['rewards']
            for reward in rewards:
                data_dict = {
                    'id': reward['id'],
                    'name': reward['name'],
                    'description': reward['description'],
                    'duration': reward['duration'],
                    'cate': reward['typeText'],
                    'status': reward['statusText'],
                    'pubtime': reward['pubTime']
                }
                result_list.append(data_dict)
            return result_list
        else:
            return None, response.status_code
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrderModel, WebsiteModel):
    print('开始爬取码市订单')
    sheet = wb['Sheet']
    sheet.title = '码市'
    sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(1)
    for i in range(10, 0, -1):
        url = 'https://codemart.com/api/project?page=%d' % i
        result = get_one_page(url)
        if isinstance(result, list):
            for r in result:
                time_stamp = int(r['pubtime']) / 1000
                publish_time = datetime.fromtimestamp(time_stamp)
                if publish_time < time_point:
                    continue
                desc = ILLEGAL_CHARACTERS_RE.sub(r'', r['description'])
                cid = 'cm-{}'.format(r['id'])
                contact = get_contact(desc)
                link = 'https://codemart.com/project/{}'.format(r['id'])
                is_valid = True if r['status'] == '招募中' else False
                order_query = session.query(OrderModel).get(cid)
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
                    order = OrderModel(id=cid, desc=desc, link=link, contact=contact, category=r['cate'],
                                       pub_time=publish_time, is_valid=is_valid, is_delete=False if is_valid else True)
                    order.website = website
                    session.add(order)
                    if is_valid == True:
                        sheet.append([count, desc, link, publish_time, contact, ''])
                        count += 1
            session.commit()
        elif isinstance(result, tuple):
            message = '码市爬取第%d行出错：%s' % (result[0], result[1])
            print(message)
            send_message(message)
    print('结束爬取码市订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
