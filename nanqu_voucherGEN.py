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


def genExcel_nanqu_shuaka(excelpath_sheet15, excelpath_sheet2, save_dir):
    # 南区刷卡
    # excelpath_sheet15, excelpath_sheet2, excelpath_sheet18:报表十五、报表二
    # save_dir 末端不加反斜杠
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet15)  # 报表十五
    # df_data_1.iloc[:,0].name
    # old_col = list(df_data_1.columns)
    # old_col[0] = 'index'
    # df_data_1.columns = old_col
    # df_data_1 = df_data_1.loc[df_data_1['index'].isna()]
    df_data_1_1 = df_data_1.loc[df_data_1[df_data_1.iloc[:, 0].name].isna()]
    # 筛选去掉“合计”行
    # print(df_data_1_1)
    df_1 = df.copy(deep=True)
    df_1['原币金额'] = df_data_1_1['结算金额']
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = df_data_1_1['结算金额']
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '收各单位水费（信用卡） （南区）', '1002', '银行存款', '银行账户', '2.03.001',
        "建行松苑办'44001780308051308979"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_data_2 = pd.read_excel(excelpath_sheet2,
                              sheet_name='营业厅收费汇总报表_刷卡')  # 报表二
    # print(df_data_2)
    df_2 = df.copy(deep=True)
    amount_df_2_firstSubtrahend = [
        round(abs(each), 2)
        for each in df_data_1.loc[df_data_1[df_data_1.iloc[:, 0].name] == '合计',
                                  '商户费用'].tolist()
    ]
    # 第一行被减数报表十五“商户费用”（手续费）列合计
    amount_df_2_all = [
        round(each, 2) for each in df_data_2.loc[
            df_data_2['序号'] == '合计',
            ['水费', '污水费', '垃圾处理费', '预收款收支', '违约金', '税额']].values.tolist()[0]
    ]
    df_2['原币金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['摘要'] = [
        '收各单位水费（信用卡） （南区）', '代收污水费 （南区）', '代收垃圾费 （南区）', '收到预收水费（南区）', '收到违约金 南区', '收到违约金 南区 销项税'
    ]
    df_2['科目'] = [
        '1122.001', '2241.003.001', '2241.003.002', '2203.004', '6301.003', '2221.016.002'
    ]
    df_2['科目名称'] = [
        '应收账款_自来水', '其他应付款_外部单位往来款_污水费', '其他应付款_外部单位往来款_垃圾费', '预收账款_水费', '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%'
    ]
    df_2['核算项目1'] = ['', '供应商', '供应商', '', '行政组织', '']
    df_2['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '', '2.01.01.01.01.25', ''
    ]
    df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '', '南区营业厅', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = ''
    print(df_sum)
    df_sum.to_excel(save_dir + "\\" + f'nanqu_shuaka.xlsx', index=False)
    return df_sum


def genExcel_nanqu_xianjin(excelpath_sheet1, save_dir):
    # 南区现金
    # save_dir 末端不加反斜杠
    # excelpath_sheet1:报表一
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet1,
                              sheet_name='营业厅收费汇总报表_现金')  # 报表一
    df_data_1_1 = df_data_1.loc[df_data_1['序号'] == '合计']
    df_1 = df.copy(deep=True)
    df_1['原币金额'] = df_data_1_1['合计']
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = df_data_1_1['合计']
    df_1['摘要'] = f'收到{last_month}月水费 南区'
    df_1['科目'] = '1001'
    df_1['科目名称'] = '库存现金'
    df_sum = df_sum.append(df_1, ignore_index=True)

    amount_df_2_all = [
        round(each, 2) for each in df_data_1.loc[
            df_data_1['序号'] == '合计',
            ['水费', '污水费', '垃圾处理费', '预收款收支', '违约金', '税额']].values.tolist()[0]
    ]
    df_2 = df.copy(deep=True)
    df_2['原币金额'] = amount_df_2_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = amount_df_2_all
    df_2['摘要'] = [
        f'收到{last_month}月水费 南区',
        '代收污水处理费 南区',
        '代收垃圾处理费 南区',
        '收到预收水费 南区',
        '收到违约金 南区',
        '收到违约金 南区 销项税',
    ]
    df_2['科目'] = [
        '1122.001',
        '2241.003.001',
        '2241.003.002',
        '2203.004',
        '6301.003',
        '2221.016.002',
    ]
    df_2['科目名称'] = [
        '应收账款_自来水',
        '其他应付款_外部单位往来款_污水费',
        '其他应付款_外部单位往来款_垃圾费',
        '预收账款_水费',
        '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%',
    ]
    df_2['核算项目1'] = ['', '供应商', '供应商', '', '行政组织', '']
    df_2['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '', '2.01.01.01.01.25', ''
    ]
    df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '', '南区营业厅', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00003'
    print(df_sum)
    df_sum.to_excel(save_dir + "\\" + f'nanqu_xianjin.xlsx',
                    index=False)
    return df_sum


def genExcel_nanqu_yinhanghuazhang(excelspath_sheet28_of_dir, save_dir):
    # 南区-银行划账
    # save_dir 末端不加反斜杠
    # excelspath_sheet28_of_dir:报表二十八s所在的文件路径

    if os.path.exists(save_dir + '\\nanqu_yinhanghuazhang') == False:
        os.mkdir(save_dir + '\\nanqu_yinhanghuazhang')
    excelspath_sheet28 = [
        os.path.join(excelspath_sheet28_of_dir, i)
        for i in os.listdir(excelspath_sheet28_of_dir) if i.endswith('.xlsx')
    ]
    # print(excelspath_sheet28)
    for excelpath_sheet28 in excelspath_sheet28:
        # 把定值全表深拷贝一份最为初始
        df_sum = df.copy(deep=True)
        df_data_1 = pd.read_excel(excelpath_sheet28,
                                  sheet_name='划帐情况汇总')  # 报表二十八
        df_data_1.fillna(0, inplace=True)
        # print(df_data_1)
        acountID = str(df_data_1['划帐ID'].values.tolist()[0])

        df_1 = df.copy(deep=True)
        df_1['原币金额'] = df_data_1.loc[df_data_1['项目'] == '总金额',
                                     '实收金额'].tolist()[0]
        df_1['方向'] = '1'  # 借方
        df_1['借方金额'] = df_data_1.loc[df_data_1['项目'] == '总金额',
                                     '实收金额'].tolist()[0]
        df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
            '收各单位水费（信用卡） （南区）', '1002', '银行存款', '银行账户', '2.01.003',
            "工商行香山支行'2011002109022109510"
        ]
        df_sum = df_sum.append(df_1, ignore_index=True)

        df_data_2 = df_data_1.copy(deep=True)
        df_data_2.index = df_data_2['项目']
        a = df_data_2.columns.tolist()
        a.remove('项目')
        df_data_2 = df_data_2[a].T
        # 以项目作index然后转置矩阵
        amount_df_2_all = df_data_2.loc[
            '实收金额', ['基本水费', '污水费', '垃圾费', '滞纳金']].values.tolist()
        amount_df_2_all = amount_df_2_all[0:3] + [
            amount_df_2_all[3] - amount_df_2_all[3] * 0.06
        ] + [amount_df_2_all[3] * 0.06
             ] + df_data_2.loc['重复金额', ['总金额']].values.tolist()
        amount_df_2_all = [round(each, 2) for each in amount_df_2_all]
        # print(amount_df_2_all)
        df_2 = df.copy(deep=True)
        df_2['原币金额'] = amount_df_2_all
        df_2['方向'] = '0'  # 贷方
        df_2['贷方金额'] = amount_df_2_all
        df_2['摘要'] = [
            '收各单位水费（信用卡） （南区）', '代收污水费 （南区）', '代收垃圾费 （南区）', '收水费违约金 （南区）', '水费违约金销项税 （南区）', '收到预收水费 南区'
        ]
        df_2['科目'] = [
            '1122.001', '2241.003.001', '2241.003.002', '6301.003',
            '2221.016.002', '2203.004'
        ]
        df_2['科目名称'] = [
            '应收账款_自来水', '其他应付款_外部单位往来款_污水费', '其他应付款_外部单位往来款_垃圾费',
            '营业外收入_违约金收入', '应交税费_简易计税_简易计税3%', '预收账款_水费'
        ]
        df_2['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
        df_2['编码1'] = [
            '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.25', '', ''
        ]
        df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '南区营业厅', '', '']
        df_sum = df_sum.append(df_2, ignore_index=True)

        df_sum = constant_value(df_sum)
        # df_sum['凭证号'] = ''
        print(df_sum)
        df_sum.to_excel(save_dir + '\\nanqu_yinhanghuazhang' + '\\' +
                        f'nanqu_yinhanghuazhang{acountID}.xlsx',
                        index=False)
        # return df_sum



if __name__ == "__main__":
    pass
    # genExcel_nanqu_shuaka(
    #     excelpath_sheet15=
    #          r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\扫码-刷卡\2021-10-13\南区-刷卡\南区-刷卡.xlsx',
    #     excelpath_sheet2=
    #          r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\南区\营业厅收费汇总报表\db_营业厅收费汇总报表.xlsx',
    #     save_dir=r'.\pingzheng',
    # )
    # genExcel_nanqu_xianjin(
    #     excelpath_sheet1=
    #          r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\南区\营业厅收费汇总报表\db_营业厅收费汇总报表.xlsx',
    #     save_dir=r'.\pingzheng',
    # )
    genExcel_nanqu_yinhanghuazhang(
        excelspath_sheet28_of_dir=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\南区\划帐情况汇总',
        save_dir=r'.\pingzheng'
    )