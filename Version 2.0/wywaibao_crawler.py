'''
51外包抓取
'''

import random
import time
from datetime import datetime

import requests
from lxml import etree
from openpyxl import Workbook

from config import time_point
from sender import send_message
from utils import get_contact, get_ua, get_mysql_connection, create_table, add_default_data


def get_links(url):
    link_list = []
    try:
        text = requests.get(url, headers={'User-Agent':get_ua()}).text
        html = etree.HTML(text)
        orders = html.xpath('//*[@class="xiangmu_item"]')
        for order in orders:
            link = 'http://www.51waibao.net/' + order.xpath('./div[1]/div[1]/a/@href')[0]
            link_list.append(link)
        return link_list
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def get_detail(url):
    try:
        text = requests.get(url, headers={'User-Agent': get_ua()}).text
        html = etree.HTML(text)
        info = html.xpath('//*[@id="form1"]/div[6]/div[3]')[0]
        wid = info.xpath('./div[1]/div[1]/ul/li[1]/text()')[0].split('waibao')[1]
        cate = info.xpath('./div[1]/div[1]/ul/li[2]/text()')[0][6:]
        status = info.xpath('./div[1]/div[1]/ul/li[6]/text()')[0]
        pub_time = info.xpath('./div[1]/div[1]/ul/li[7]/text()')[0][6:]
        desc_list = info.xpath('./div[2]/div[2]//text()')
        desc = '\n'.join([dl.strip() for dl in desc_list])
        return [wid, cate, status, pub_time, desc]
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrdrModel, WebsiteModel):
    print('开始爬取51外包订单')
    sheet = wb.create_sheet('51外包', 4)
    sheet.append(['单据编号', '订单描述', '链接', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(5)
    for i in range(10, 0, -1):
        url = 'http://www.51waibao.net/Project.html?page=%d' % i
        link_list = get_links(url)
        if isinstance(link_list, list):
            for link in link_list:
                result = get_detail(link)
                if isinstance(result, list):
                    date_str = result[3]
                    publish_time = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                    if publish_time < time_point:
                        continue
                    desc = result[4]
                    contact = get_contact(desc)
                    wid = 'wy-' + result[0]
                    is_valid = False if '项目已过期' in result[2] else True
                    order_query = session.query(OrdrModel).get(wid)
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
                        order = OrdrModel(id=wid, desc=desc, link=link, contact=contact, category=result[1], pub_time=publish_time, is_valid=is_valid, is_delete=False if is_valid else True)
                        order.website = website
                        session.add(order)
                        if is_valid == True:
                            sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
                            count += 1
                else:
                    message = '51外包详情爬取第%d行出错：%s' % (result[0], result[1])
                    print(message)
                    send_message(message)
                time.sleep(random.random() / 10)
            session.commit()
        elif isinstance(link_list, tuple):
            message = '51外包爬取第%d行出错：%s' % (link_list[0], link_list[1])
            print(message)
            send_message(message)
    print('结束爬取51外包订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
