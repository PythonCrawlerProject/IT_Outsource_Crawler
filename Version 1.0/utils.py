'''
工具函数文件
'''

import re

from fake_useragent import UserAgent

from config import contact_regex


def get_ua():
    '''获取随机请求头'''
    try:
        return UserAgent().chrome
    except:
        get_ua()


def get_contact(desc):
    '''根据字符串获取联系人'''
    contact_group = re.findall(contact_regex, desc, re.VERBOSE)
    if len(contact_group):
        contact_group = [e for t in contact_group for e in t if e != '']
        contact_group = list(set(contact_group))
        return '|'.join(contact_group)
    return ''
