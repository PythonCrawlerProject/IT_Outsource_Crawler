'''
开源中国抓取
'''

from datetime import datetime

import html2text
import requests
from openpyxl import Workbook

from config import oschina_headers
from utils import get_contact, get_ua
from sender import send_message


def get_id(url):
    try:
        response = requests.get(url, headers=oschina_headers.update({'User-Agent': get_ua()}))
        if response.status_code == 200:
            data = response.json()
            try:
                datas = data['data']['data']
                id_list = [d['id'] for d in datas]
                return id_list
            except Exception as e:
                return None, e.args[0]
        else:
            return None, response.status_code
    except Exception as e:
        return None, e.args[0]


def get_one_page(url):
    try:
        response = requests.get(url, headers=oschina_headers)
        if response.status_code == 200:
            data = response.json()
            try:
                description = data['data']['prd']
                return description
            except Exception as e:
                return None, e.args[0]
        else:
            return None, response.status_code
    except Exception as e:
        return None, e.args[0]


def main(wb):
    print('开始爬取开源中国订单')
    sheet = wb.create_sheet('开源中国', 1)
    sheet.append(['单据编号', '订单描述', '链接', '分配人员'])
    count = 1
    for i in range(10):
        url = 'https://zb.oschina.net/project/contractor-browse-project-and-reward?applicationAreas=&moneyMinByYuan=&moneyMaxByYuan=&sortBy=30&currentTime=&pageSize=20&currentPage='.format(
            i + 1)
        id_list = get_id(url)
        if isinstance(id_list, list):
            for id in id_list:
                url = 'https://zb.oschina.net/project/detail?id=%s' % id
                desc = get_one_page(url)
                if isinstance(desc, str):
                    desc = html2text.html2text(desc).strip()
                    contact = get_contact(desc)
                    sheet.append([count, desc, url, contact])
                    count += 1
                elif isinstance(desc, tuple):
                    print('开源中国详情爬取出错：%s' % desc[1])
        elif isinstance(id_list, tuple):
            message = '开源中国爬取出错：%s' % id_list[1]
            print(message)
            send_message(message)
    print('结束爬取开源中国订单')


if __name__ == '__main__':
    wb = Workbook()
    main(wb)
    now = datetime.now()
    wb.save(r'%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
