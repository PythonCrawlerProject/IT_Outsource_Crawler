'''
猿急送抓取
'''

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
        html = etree.HTML(text)
        orders = html.xpath('//*[@id="db_adapt_id"]/div[position()>2]')
        for order in orders:
            info = {}
            link = str(order.xpath('./a/@href')[0])
            desc_str = order.xpath('./div[1]/div[1]/div/a/p/text()')[0]
            desc = emoji_regex.sub('[Emoji]', desc_str)
            status = order.xpath('./div[2]/a/text()')[0]
            info['link'] = link
            info['desc'] = desc.strip()
            info['status'] = status
            info_list.append(info)
        return info_list
    except Exception as e:
        return e.__traceback__.tb_lineno, e.args[0]


def main(wb, session, OrderModel, WebsiteModel):
    print('开始爬取猿急送订单')
    sheet = wb.create_sheet('猿急送', 5)
    sheet.append(['单据编号', '订单描述', '链接', '发布时间', '联系方式', '分配人员'])
    count = 1
    website = session.query(WebsiteModel).get(6)
    for i in range(10, 0, -1):
        url = 'https://www.yuanjisong.com/job/allcity/page%d' % i
        info_list = get_info(url)
        if isinstance(info_list, list):
            for info in info_list:
                desc = info['desc']
                link = info['link']
                contact = get_contact(desc)
                is_valid = True if info['status'] == '投递职位' else False
                yid = 'yj-{}'.format(int(link.split('/')[-1]))
                order_query = session.query(OrderModel).get(yid)
                if order_query:
                    is_valided = order_query.is_valid
                    order_query.is_valid = is_valid
                    # if is_valided == False and is_valid == True:
                    #     sheet.append([count, desc, link, contact])
                    #     count += 1
                    #     order_query.is_delete = False
                    if is_valid == True:
                        sheet.append([count, desc, link, '', contact, ''])
                        count += 1
                        if is_valided == False:
                            order_query.is_delete = False
                    if is_valided == True and is_valid == False:
                        order_query.is_delete = True
                else:
                    order = OrderModel(id=yid, desc=desc, link=link, contact=contact, category='',
                                      pub_time=None, is_valid=is_valid, is_delete=False if is_valid else True)
                    order.website = website
                    session.add(order)
                    if is_valid == True:
                        sheet.append([count, desc, link, '', contact, ''])
                        count += 1
            session.commit()
        elif isinstance(info_list, tuple):
            message = '猿急送爬取第%d行出错：%s' % (info_list[0], info_list[1])
            print(message)
            send_message(message)
    print('结束爬取猿急送订单')


if __name__ == '__main__':
    wb = Workbook()
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)
    main(wb, session, Order, Website)
    now = datetime.now()
    wb.save(r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
