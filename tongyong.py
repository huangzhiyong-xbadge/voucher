import pandas as pd
import numpy as np
import datetime
import os
import re

# forVoucherNo 用于循环赋予 excel 们凭证号（排名不分先后）
# forItemNo 用于给每一个 excel 内部存在的科目


def forVoucherNo(datadir):
    # 递归获取这里面的所有子孙目录下的 xlsx 后缀，然后循环赋予凭证号
    pass


def forItemNo(datadir):
    # 递归获取这里面的所有子孙目录下的 xlsx 后缀，分别处理分录号
    pass


if __name__ == "__main__":
    pass
    forVoucherNo(datadir=r'.\data')
    forItemNo(datadir=r'.\data')
