'''
人人开发抓取
'''

from datetime import datetime

import requests
from lxml import etree
from openpyxl import Workbook

from utils import get_contact, get_ua, get_mysql_connection, create_table,  add_default_data
from sender import send_message


def get_info(url):
    info_list = []
    try:
        text = requests.get(url, headers={'User-Agent':get_ua()}).text
        html = etree.HTML(text)
        orders = html.xpath('//*[@id="r-list-wrapper"]/div[2]/div')
        for order in orders:
            info = {}
            link = 'http://www.rrkf.com' + order.xpath('./div[1]/div/h4/a/@href')[0]
            desc = order.xpath('./div[1]/div/p/text()')[0]
            info['link'] = link
            info['desc'] = desc.strip()
            info_list.append(info)
        return info_list
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def get_detail(url):
    try:
        text = requests.get(url, headers={'User-Agent': get_ua()}).text
        html = etree.HTML(text)
        status_str = html.xpath('//*[@id="step-box"]/ul/li[1]/span/span/text()')
        status = status_str[0] if status_str else '定标及以后'
        pub_date = html.xpath('//*[@id="step-box"]/ul/li[1]/div/span[2]/text()')
        pub_time = pub_date[0] if pub_date else None
        return [status, pub_time]
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrderModel, WebsiteModel):
    print('开始爬取人人开发订单')
    sheet = wb.create_sheet('人人开发', 2)
    sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(3)
    for i in range(10, 0, -1):
        url = 'http://www.rrkf.com/serv/request?&currentPage=%d' % i
        info_list = get_info(url)
        if isinstance(info_list, list):
            for info in info_list:
                desc = info['desc']
                link = info['link']
                details = get_detail(link)
                if isinstance(details, list):
                    rid = 'rr-{}'.format(link.split('=')[1])
                    contact = get_contact(desc)
                    is_valid = True if '剩余' in details[0] else False
                    pub_time = datetime.strptime(details[1], "%Y-%m-%d %H:%M:%S") if details[1] else None
                    order_query = session.query(OrderModel).get(rid)
                    if order_query:
                        is_valided = order_query.is_valid
                        order_query.is_valid = is_valid
                        if is_valid == True:
                            sheet.append([count, desc, link, pub_time, contact, ''])
                            count += 1
                            if is_valided == False:
                                order_query.is_delete = False
                        if is_valided == True and is_valid == False:
                            order_query.is_delete = True
                    else:
                        order = OrderModel(id=rid, desc=desc, link=link, contact=contact, category='',
                                          pub_time=pub_time, is_valid=is_valid, is_delete=False if is_valid else True)
                        order.website = website
                        session.add(order)
                        if is_valid == True:
                            sheet.append([count, desc, link, pub_time, contact, ''])
                            count += 1
                else:
                    message = '人人开发详情爬取第%d行出错：%s' % (details[0], details[1])
                    print(message)
                    send_message(message)
            session.commit()
        elif isinstance(info_list, tuple):
            message = '人人开发爬取第%d行出错：%s' % (info_list[0], info_list[1])
            print(message)
            send_message(message)
    print('结束爬取人人开发订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
