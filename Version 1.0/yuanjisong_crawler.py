'''
猿急送抓取
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
        orders = html.xpath('//*[@id="db_adapt_id"]/div[position()>2]')
        for order in orders:
            info = {}
            link = order.xpath('./div[1]/div[2]/a/@href')[0]
            desc = order.xpath('./div[1]/div[1]/div/a/p/text()')[0]
            info['link'] = link
            info['desc'] = desc.strip()
            info_list.append(info)
        return info_list
    except Exception as e:
        return None, e.args[0]


def main(wb):
    print('开始爬取猿急送订单')
    sheet = wb.create_sheet('猿急送', 4)
    sheet.append(['单据编号', '订单描述', '链接', '分配人员'])
    count = 1
    for i in range(10):
        url = 'https://www.yuanjisong.com/job/allcity/page{}'.format(i + 1)
        info_list = get_info(url)
        if isinstance(info_list, list):
            for info in info_list:
                desc = info['desc']
                contact = get_contact(desc)
                sheet.append([count, desc, info['link'], contact])
                count += 1
        elif isinstance(info_list, tuple):
            message = '猿急送爬取出错：%s' % info_list[1]
            print(message)
            send_message(message)
    print('结束爬取猿急送订单')


if __name__ == '__main__':
    wb = Workbook()
    main(wb)
    now = datetime.now()
    wb.save(r'%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
