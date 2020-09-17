'''
实现抓取
'''

import re
import time
import random
from datetime import datetime

import requests
from lxml import etree
from openpyxl import Workbook

from utils import get_contact, get_ua, get_mysql_connection, create_table,  add_default_data
from sender import send_message
from config import emoji_regex


def get_info(url):
    info_list = []
    try:
        text = requests.get(url, headers={'User-Agent':get_ua()}).text
        orders = re.findall(r'<div class="job">(.*?)<div class="clearfix"></div>', text, re.S | re.M)
        for order in orders:
            info = {}
            link = 'http://www.shixian.com' + re.search(r'<a target="_blank" href="(.+?)">', order).groups()[0]
            desc_str = re.search(r'<p class="describe text-inline-limit">(.*?)</p>', order, re.S | re.M).groups()[0]
            desc = emoji_regex.sub('[Emoji]', desc_str)
            start_time = re.search(r'.*?(\d{4}-\d{2}-\d{2}).*?', order, re.S | re.M).groups()[0]
            info['link'] = str(link)
            info['desc'] = desc.strip()
            info['start_time'] = start_time + ' 23:59:59'
            info_list.append(info)
        return info_list

    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def get_category(url):
    try:
        text = requests.get(url, headers={'User-Agent': get_ua()}).text
        html = etree.HTML(text)
        cate_temp = html.xpath('/html/body/div[3]/div[1]/article/section[1]/dl/dd[1]/span/text()')
        cate = cate_temp[0] if cate_temp else ''
        return cate
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrderModel, WebsiteModel):
    print('开始爬取实现订单')
    sheet = wb.create_sheet('实现', 3)
    sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(4)
    for i in range(10, 0, -1):
        url = 'https://shixian.com/job/all?page=%d&sort_arrow=down' % i
        info_list = get_info(url)
        if isinstance(info_list, list):
            for info in info_list:
                desc = info['desc']
                link = info['link']
                contact = get_contact(desc)
                dl_time = datetime.strptime(info['start_time'], "%Y-%m-%d %H:%M:%S")
                is_valid = True if datetime.now() <= dl_time else False
                sid= 'sx-' + link.split('/')[-1]
                cate = get_category(link)
                if isinstance(cate, str):
                    order_query = session.query(OrderModel).get(sid)
                    if order_query:
                        is_valided = order_query.is_valid
                        order_query.is_valid = is_valid
                        if is_valid == True:
                            sheet.append([count, desc, link, '', contact, ''])
                            count += 1
                            if is_valided == False:
                                order_query.is_delete = False
                        if is_valided == True and is_valid == False:
                            order_query.is_delete = True
                    else:
                        order = OrderModel(id=sid, desc=desc, link=link, contact=contact, category=cate,
                                          pub_time=None, is_valid=is_valid, is_delete=False if is_valid else True)
                        order.website = website
                        session.add(order)
                        if is_valid == True:
                            sheet.append([count, desc, link, '', contact, ''])
                            count += 1
                else:
                    message = '实现详情爬取第%d行出错：%s' % (cate[0], cate[1])
                    print(message)
                    send_message(message)
                time.sleep(random.random()/10)
            session.commit()
        elif isinstance(info_list, tuple):
            message = '实现爬取第%d行出错：%s' % (info_list[0], info_list[1])
            print(message)
            send_message(message)
    print('结束爬取实现订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
