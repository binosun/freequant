
# coding: utf-8

from sqlalchemy import Column, String,Integer, create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base

# 创建对象的基类:
Base = declarative_base()

# 定义User对象:
class SellCount(Base):
    # 表的名字:
    __tablename__ = 'sell_count'

    # 表的结构:
    product_id = Column(Integer)
    name = Column(String(20))


# 初始化数据库连接:
engine = create_engine('mysql+mysqlconnector://root:root@localhost:3306/tmall')
# 创建DBSession类型:
DBSession = sessionmaker(bind=engine)