'''
爬取主程序
'''

import time
from datetime import datetime

from openpyxl import Workbook
from apscheduler.schedulers.blocking import BlockingScheduler

import shixian_crawler, rrkf_crawler, wywaibao_crawler, codemart_crawler, yuanjisong_crawler, oschina_crawler
from sender import get_media_id, send_file,send_message

sched = BlockingScheduler()


def crawl_save_upload():
    '''调用函数实现抓取、保存和上传数据文件'''
    print('-----数据抓取开始-----')
    wb = Workbook()
    codemart_crawler.main(wb)
    oschina_crawler.main(wb)
    rrkf_crawler.main(wb)
    wywaibao_crawler.main(wb)
    yuanjisong_crawler.main(wb)
    shixian_crawler.main(wb)
    print('-----数据抓取结束-----')

    print('-----文件保存开始-----')
    now = datetime.now()
    file = r'data/%s.xlsx' % now.strftime("%Y-%m-%d %H-%M-%S")
    wb.save(file)
    time.sleep(3)
    print('-----文件保存结束-----')

    print('-----文件上传开始-----')
    media_id = get_media_id(file)
    if isinstance(media_id, str):
        upload_result = send_file(media_id)
        if upload_result == True:
            print('文件上传成功：%s' % file)
        else:
            message = '文件上传失败：%s' % upload_result[1]
            print(message)
            send_message(message)
    else:
        message = '获取media_id失败：%s' % media_id[1]
        print(message)
        send_message(message)

    print('-----文件上传结束-----')


@sched.scheduled_job('interval', seconds=1800)
def schedule():
    '''设定执行计划'''
    now = datetime.now()
    print('当前时间：%s' % now.strftime("%Y-%m-%d %H:%M:%S"))
    hour = now.hour
    if hour >= 8 and hour <= 22:
        print('程序执行开始')
        crawl_save_upload()
        print('程序执行结束\n')
    else:
        pass


if __name__ == '__main__':
    '''主函数'''
    sched.start()