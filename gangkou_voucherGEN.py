import pandas as pd
import numpy as np
import datetime
import os
import re

# todo 最后循环生成单凭号
# todo ['辅助账摘要']

# 获取当前月的上一个月的最后一天
last_date = datetime.date(datetime.date.today().year,
                          datetime.date.today().month,
                          1) - datetime.timedelta(1)
last_month = str(last_date.month)
last_date = last_date.strftime("%Y-%m-%d")
colname = [
    '公司', '记账日期', '业务日期', '会计期间', '凭证类型', '凭证号', '分录号', '摘要', '科目', '科目名称',
    '币种', '汇率', '方向', '原币金额', '数量', '单价', '借方金额', '贷方金额', '制单人', '过账人', '审核人',
    '附件数量', '过账标记', '机制凭证模块', '删除标记', '凭证序号', '单位', '参考信息', '是否有现金流量',
    '现金流量标记', '业务编号', '结算方式', '结算号', '辅助账摘要', '核算项目1', '编码1', '名称1', '核算项目2',
    '编码2', '名称2', '核算项目3', '编码3', '名称3', '核算项目4', '编码4', '名称4', '核算项目5', '编码5',
    '名称5', '核算项目6', '编码6', '名称6', '核算项目7', '编码7', '名称7', '核算 项目8', '编码8',
    '名称8', '发票号', '换票证号', '客户', '费用类别', '收款人', '物料', '财务组织', '供应商', '辅助账业务日期',
    '到期日'
]
df = pd.DataFrame(columns=colname)


def constant_value(df):
    df[[
        '公司', '业务日期', '会计期间', '辅助账业务日期', '凭证类型', '币种', '汇率', '制单人', '过账标记',
        '删除标记', '现金流量标记', '数量', '单价'
    ]] = [
        '2.01.01.01.01',
        last_date,
        last_month,
        last_date,
        '银收',
        'BB01',
        1,
        'test',
        'FALSE',
        'FALSE',
        6,
        0,
        0,
    ]
    cols = ["原币金额", "借方金额", "贷方金额"]
    df[cols] = df[cols].replace({'0': np.nan, 0: np.nan})
    print(df.loc[(df["原币金额"].notna()), ['原币金额', '借方金额', '贷方金额']])
    df = df.loc[df["原币金额"].notna()]
    return df


def genExcel_gangkou_saomazhifu():
    pass



if __name__ == "__main__":
    pass
    genExcel_gangkou_saomazhifu()