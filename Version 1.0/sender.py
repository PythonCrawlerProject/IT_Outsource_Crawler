'''
数据文件上传和发送
'''

import os
import requests

from config import upload_wechat_key, notify_wechat_key


def get_media_id(filename):
    '''上传文件获取media_id'''
    try:
        headers = {"Content-Type": "multipart/form-data"}
        send_url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={}&type=file'.format(upload_wechat_key)
        file = {
            (filename, open(filename, "rb")),
        }
        res = requests.post(url=send_url, headers=headers, files=file).json()
        media_id = res['media_id']
        return media_id
    except Exception as e:
        return None, e.args[0]


def send_file(media_id):
    '''发送文件'''
    try:
        headers = {'Content-Type': 'application/json'}
        data = {
            "msgtype": "file",
            "file": {
                "media_id": media_id
            }
        }
        send_url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={}'.format(upload_wechat_key)
        r = requests.post(url=send_url, headers=headers, json=data).json()
        if r['errcode'] == 0:
            return True
        else:
            return None, r['errmsg']
    except Exception as e:
        return None, e.args[0]


def send_message(message):
    '发送错误消息'
    try:
        headers = {'Content-Type': 'application/json'}
        data = {
            "msgtype": "text",
            "text": {
                "content": message,
                "mentioned_mobile_list": ["15682210532"]
            }
        }
        send_url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={}'.format(notify_wechat_key)
        r = requests.post(url=send_url, headers=headers, json=data).json()
        if r['errcode'] == 0:
            return True
        else:
            return None, r['errmsg']
    except Exception as e:
        return None, e.args[0]


if __name__ == '__main__':
    files = os.listdir('./data')
    files = [f for f in files if f.endswith('xlsx')]
    latest_file = files[-1]
    media_id = get_media_id('data/' + latest_file)
    upload_result = send_file(media_id)
    if upload_result == True:
        print('上传成功')
    else:
        send_message('上传失败：'+upload_result[1])
