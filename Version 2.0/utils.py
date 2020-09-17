'''
工具函数文件
'''

import re
import os
from datetime import datetime

from fake_useragent import UserAgent
from sqlalchemy import create_engine
from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey, Boolean
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy.ext.declarative import declarative_base

from config import contact_regex, DB_URL, web_name_list, web_url_list, reserve_file_count
from sender import send_message


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

def delete_data():
    files = os.listdir('./data')
    file_count = len(files)
    if file_count > reserve_file_count:
        delete_count = file_count-reserve_file_count
        delete_files = files[:delete_count]
        for file in delete_files:
            os.remove('data/' + file)
        message = '已删除过期文件%d个' % delete_count
        print(message)
        send_message(message)

def get_mysql_connection():
    '''连接MySQL数据库'''
    engine = create_engine(DB_URL)
    Base = declarative_base(engine)
    session = sessionmaker(bind=engine)()

    return engine, Base, session


def create_table(engine, Base):
    '''创建表'''
    class Website(Base):
        __tablename__ = 'website'
        id = Column(Integer, primary_key=True, autoincrement=True)
        name = Column(String(10), nullable=False)
        link = Column(String(40), nullable=False)

        orders = relationship('Order', backref='website')

    class Order(Base):
        __tablename__ = 'order'
        id = Column(String(50), primary_key=True)
        desc = Column(Text, nullable=False)
        link = Column(String(80), nullable=False)
        contact = Column(String(30))
        category = Column(String(15), nullable=True)
        pub_time = Column(DateTime, nullable=True)
        is_valid = Column(Boolean, nullable=False)
        add_time = Column(DateTime, default=datetime.now)
        wid = Column(Integer, ForeignKey('website.id'), nullable=False)
        is_delete = Column(Boolean, default=False)

    if (not engine.dialect.has_table(engine, 'website')) or (not engine.dialect.has_table(engine, 'order')):
        Base.metadata.create_all()
        print('表创建成功')

    return Order, Website


def add_default_data(session, WebsiteModel):
    origin_data = session.query(WebsiteModel).all()
    if len(origin_data) != 6:
        for data in origin_data:
            session.delete(data)
        session.commit()
        for name, url in zip(web_name_list, web_url_list):
            website = WebsiteModel(name=name, link=url)
            session.add(website)
        session.commit()
        print('插入数据成功')


if __name__ == '__main__':
    engine, Base, session = get_mysql_connection()
    Order, Website = create_table(engine, Base)
    add_default_data(session, Website)