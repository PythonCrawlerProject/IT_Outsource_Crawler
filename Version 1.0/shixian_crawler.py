'''
实现抓取
'''

from datetime import datetime

import requests
from lxml import etree
from openpyxl import Workbook

from utils import get_contact, get_ua
from sender import send_message


def get_info(url):
    info_list = []
    try:
        text = requests.get(url, headers={'User-Agent':get_ua()}).text
        html = etree.HTML(text)
        orders = html.xpath('//*[@class="job"]')
        for order in orders:
            info = {}
            link = 'http://www.shixian.com' + order.xpath('./div[1]/a/@href')[0]
            desc = order.xpath('./div[1]/a/p/text()')[0]
            # release_time = order.xpath('./div[1]/div/div/span/text()')[0]
            # if '1 天前发布' in release_time or '小时' in release_time:
            #     info['link'] = link
            #     info['desc'] = desc.strip()
            #     info_list.append(info)
            # else:
            #     continue
            info['link'] = link
            info['desc'] = desc.strip()
            info_list.append(info)
        return info_list
    except Exception as e:
        return None, e.args[0]


def main(wb):
    print('开始爬取实现订单')
    sheet = wb.create_sheet('实现', 5)
    sheet.append(['单据编号', '订单描述', '链接', '分配人员'])
    count = 1
    for i in range(10):
        url = 'https://shixian.com/job/all?page={}&sort_arrow=down'.format(i + 1)
        info_list = get_info(url)
        if isinstance(info_list, list):
            for info in info_list:
                desc = info['desc']
                contact = get_contact(desc)
                sheet.append([count, desc, info['link'], contact])
                count += 1
        elif isinstance(info_list, tuple):
            message = '实现爬取出错：%s' % info_list[1]
            print(message)
            send_message(message)
    print('结束爬取实现订单')


if __name__ == '__main__':
    wb = Workbook()
    main(wb)
    now = datetime.now()
    wb.save(r'%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
