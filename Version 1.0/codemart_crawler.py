'''
码市订单抓取
'''

import time
from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from config import codemart_headers
from utils import get_contact
from sender import send_message


def get_one_page(url, start_time):
    result_list = []
    try:
        response = requests.get(url, headers=codemart_headers)
        if response.status_code == 200:
            data = response.json()
            try:
                rewards = data['rewards']
                for reward in rewards:
                    # pub_time = float(reward['pubTime']) / 1000
                    # if start_time - pub_time < 7200:
                    #     data_dict = {
                    #         'id': reward['id'],
                    #         'name': reward['name'],
                    #         'description': reward['description'],
                    #         'duration': reward['duration'],
                    #     }
                    #     result_list.append(data_dict)
                    #     continue
                    # else:
                    #     return result_list
                    data_dict = {
                        'id': reward['id'],
                        'name': reward['name'],
                        'description': reward['description'],
                        'duration': reward['duration'],
                    }
                    result_list.append(data_dict)
                return result_list
            except Exception as e:
                return None, e.args[0]
        else:
            return None, response.status_code
    except Exception as e:
        return None, e.args[0]


def main(wb):
    print('开始爬取码市订单')
    start_time = time.time()
    sheet = wb['Sheet']
    sheet.title = '码市'
    sheet.append(['单据编号', '订单描述', '链接', '分配人员'])
    count = 1
    for i in range(10):
        url = 'https://codemart.com/api/project?page={}'.format(i + 1)
        result = get_one_page(url, start_time)
        if isinstance(result, list):
            for r in result:
                desc = ILLEGAL_CHARACTERS_RE.sub(r'', r['description'])
                contact = get_contact(desc)
                sheet.append([count, desc, 'https://codemart.com/project/{}'.format(r['id']), contact])
                count += 1
        elif isinstance(result, tuple):
            message = '码市爬取出错：%s' % result[1]
            print(message)
            send_message(message)
    print('结束爬取码市订单')


if __name__ == '__main__':
    wb = Workbook()
    main(wb)
    now = datetime.now()
    wb.save(r'%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S"))
