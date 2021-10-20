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


def genExcel_chengqu_shuaka(excelpath_sheet15, excelpath_sheet2,
                            excelpath_sheet18, save_dir, type1):
    # 竹苑刷卡--->城区
    # excelpath_sheet15, excelpath_sheet2, excelpath_sheet18:报表十五、报表二、报表十八
    # save_dir 末端不加反斜杠
    # type1 城区营业厅  str     zhuyuan or liantang
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
        '收各单位水费（信用卡）', '1002', '银行存款', '银行账户', '2.03.001',
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
            ['水费', '污水费', '垃圾处理费', '违约金', '税额', '预收款收支']].values.tolist()[0]
    ]
    df_2['原币金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['摘要'] = [
        '收各单位水费（信用卡）', '代收污水费', '代收垃圾费', '收水费违约金', '水费违约金销项税', '收到预收水费'
    ]
    df_2['科目'] = [
        '1122.001', '2241.003.001', '2241.003.002', '6301.003', '2221.016.002',
        '2203.004'
    ]
    df_2['科目名称'] = [
        '应收账款_自来水', '其他应付款_外部单位往来款_污水费', '其他应付款_外部单位往来款_垃圾费', '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%', '预收账款_水费'
    ]
    df_2['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
    df_2['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.22.01', '', ''
    ]
    df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '城区分公司客服部', '', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_data_3 = pd.read_excel(excelpath_sheet18, header=2)  # 报表十八
    df_data_3.fillna(0, inplace=True)
    amount_df_3_col = [
        '税前费用', '销项税(%6)', '税前费用', '销项税(%6)', '税前费用', '销项税(%6)', '交易金额',
        '交易金额', '交易金额'
    ]
    amount_df_3_row = [
        '检定费', '检定费', '水质检测费', '水质检测费', '查漏费', '查漏费', '维修费', '工程费', '换表费'
    ]  # 收费项目
    amount_df_3_all = []
    for index in range(len(amount_df_3_col)):
        # print(df_data_3.loc[df_data_3['收费项目'] == amount_df_3_row[index], amount_df_3_col[index]].tolist())
        # print([round(each, 2) for each in df_data_3.loc[df_data_3['收费项目'] ==
        # amount_df_3_row[index], amount_df_3_col[index]].tolist()])
        amount_df_3_all.extend([
            round(each, 2) for each in
            df_data_3.loc[df_data_3['收费项目'] == amount_df_3_row[index],
                          amount_df_3_col[index]].tolist()
        ])
    # print(amount_df_3_all)
    df_3 = df.copy(deep=True)
    df_3['原币金额'] = amount_df_3_all
    df_3['方向'] = '0'  # 贷方
    df_3['贷方金额'] = amount_df_3_all
    df_3['摘要'] = [
        '收到检定费 6%', '收到检定费 6% 销项税', '收到水质检测费 6%', '收到水质检测费 6% 销项税', '收到查漏费 6%',
        '收到查漏费 6% 销项税', '收到维修费', '收到工程费', '收到换表费'
    ]
    df_3['科目'] = [
        '6051.005', '2221.001.002.004', '6051.001.001', '2221.001.002.004',
        '6051.007', '2221.001.002.004', '6051.001.002', '6051.001.002', ''
    ]
    df_3['科目名称'] = [
        '其他业务收入_水表检定收入', '应交税费_应交增值税_销项税额_销项税额6%', '其他业务收入_外接业务收入_水质检测收入',
        '应交税费_应交增值税_销项税额_销项税额6%', '其他业务收入_其他收入', '应交税费_应交增值税_销项税额_销项税额6%',
        '其他业务收入_外接业务收入_给水安装工程收入', '其他业务收入_外接业务收入_给水安装工程收入', ''
    ]
    df_3['核算项目1'] = ['', '', '', '', '', '', '工程项目', '工程项目', '']
    df_3['编码1'] = ['', '', '', '', '', '', '2.1.00003', '2.1.00003', '']
    df_sum = df_sum.append(df_3, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00001'
    print(df_sum)
    df_sum.to_excel(save_dir + "\\" + f'{type1}_shuaka.xlsx', index=False)
    return df_sum


def genExcel_chengqu_saoma(excelpath_sheet16, excelpath_sheet3,
                           excelpath_sheet19, save_dir, type1):
    # 竹苑微信扫码、支付宝扫码--->城区
    # save_dir 末端不加反斜杠
    # excelpath_sheet16, excelpath_sheet3, excelpath_sheet19:报表十六、报表三、报表十九
    # type1 城区营业厅  str     zhuyuan or liantang
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet16)  # 报表十六
    # df_data_1_1 = df_data_1.loc[df_data_1[df_data_1.iloc[:, 0].name].isna()]
    # 筛选去掉“合计”行
    # print(df_data_1_1)
    df_1 = df.copy(deep=True)
    df_1['原币金额'] = df_data_1['交易金额']
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = df_data_1['交易金额']
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '收各单位水费（微信、支付宝扫码）', '1002', '银行存款', '银行账户', '2.03.001',
        "建行松苑办'44001780308051308979"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_data_2 = pd.read_excel(excelpath_sheet3, sheet_name='营业厅收费汇总报表_扫码')
    print(df_data_2)
    df_2 = df.copy(deep=True)
    amount_df_2_firstSubtrahend = [
        round(abs(each), 2) for each in df_data_1.loc[:, '手续费'].tolist()
    ]
    # 第一行被减数报表十六“交易金额”（手续费）列，只有一行
    amount_df_2_all = [
        round(each, 2) for each in df_data_2.loc[
            df_data_2['序号'] == '合计',
            ['水费', '污水费', '垃圾处理费', '违约金', '税额', '预收款收支']].values.tolist()[0]
    ]
    df_2['原币金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = [amount_df_2_all[0] - amount_df_2_firstSubtrahend[0]
                    ] + amount_df_2_all[1:]
    df_2['摘要'] = [
        '收各单位水费（信用卡）', '代收污水费', '代收垃圾费', '收水费违约金', '水费违约金销项税', '收到预收水费'
    ]
    df_2['科目'] = [
        '1122.001', '2241.003.001', '2241.003.002', '6301.003', '2221.016.002',
        '2203.004'
    ]
    df_2['科目名称'] = [
        '应收账款_自来水', '其他应付款_外部单位往来款_污水费', '其他应付款_外部单位往来款_垃圾费', '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%', '预收账款_水费'
    ]
    df_2['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
    df_2['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.22.01', '', ''
    ]
    df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '城区分公司客服部', '', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_data_3 = pd.read_excel(excelpath_sheet19, header=2)  # 报表十九
    df_data_3.fillna(0, inplace=True)
    amount_df_3_col = [
        '税前费用', '销项税(%6)', '税前费用', '销项税(%6)', '税前费用', '销项税(%6)', '交易金额',
        '交易金额', '交易金额'
    ]
    amount_df_3_row = [
        '检定费', '检定费', '水质检测费', '水质检测费', '查漏费', '查漏费', '维修费', '工程费', '换表费'
    ]  # 收费项目
    amount_df_3_all = []
    for index in range(len(amount_df_3_col)):
        # print(df_data_3.loc[df_data_3['收费项目'] == amount_df_3_row[index], amount_df_3_col[index]].tolist())
        # print([round(each, 2) for each in df_data_3.loc[df_data_3['收费项目'] ==
        # amount_df_3_row[index], amount_df_3_col[index]].tolist()])
        amount_df_3_all.extend([
            round(each, 2) for each in
            df_data_3.loc[df_data_3['收费项目'] == amount_df_3_row[index],
                          amount_df_3_col[index]].tolist()
        ])
    # print(amount_df_3_all)
    df_3 = df.copy(deep=True)
    df_3['原币金额'] = amount_df_3_all
    df_3['方向'] = '0'  # 贷方
    df_3['贷方金额'] = amount_df_3_all
    df_3['摘要'] = [
        '收到检定费 6%', '收到检定费 6% 销项税', '收到水质检测费 6%', '收到水质检测费 6% 销项税', '收到查漏费 6%',
        '收到查漏费 6% 销项税', '收到维修费', '收到工程费', '收到换表费'
    ]
    df_3['科目'] = [
        '6051.005', '2221.001.002.004', '6051.001.001', '2221.001.002.004',
        '6051.007', '2221.001.002.004', '6051.001.002', '6051.001.002', ''
    ]
    df_3['科目名称'] = [
        '其他业务收入_水表检定收入', '应交税费_应交增值税_销项税额_销项税额6%', '其他业务收入_外接业务收入_水质检测收入',
        '应交税费_应交增值税_销项税额_销项税额6%', '其他业务收入_其他收入', '应交税费_应交增值税_销项税额_销项税额6%',
        '其他业务收入_外接业务收入_给水安装工程收入', '其他业务收入_外接业务收入_给水安装工程收入', ''
    ]
    df_3['核算项目1'] = ['', '', '', '', '', '', '工程项目', '工程项目', '']
    df_3['编码1'] = ['', '', '', '', '', '', '2.1.00003', '2.1.00003', '']
    df_sum = df_sum.append(df_3, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00002'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + f'{type1}_saoma.xlsx', index=False)
    return df_sum


def genExcel_chengqu_xianjinshuifei(excelpath_sheet1, save_dir, type1):
    # 竹苑现金水费--->城区
    # save_dir 末端不加反斜杠
    # excelpath_sheet1:报表一
    # type1 城区营业厅  str     zhuyuan or liantang
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet1,
                              sheet_name='营业厅收费汇总报表_现金')  # 报表一
    df_data_1_1 = df_data_1.loc[df_data_1['序号'] == '合计']
    df_1 = df.copy(deep=True)
    df_1['原币金额'] = df_data_1_1['合计']
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = df_data_1_1['合计']
    df_1['摘要'] = f'收到{last_month}月水费'
    df_1['科目'] = '1001'
    df_1['科目名称'] = '库存现金'
    df_sum = df_sum.append(df_1, ignore_index=True)

    amount_df_2_all = [
        round(each, 2) for each in df_data_1.loc[
            df_data_1['序号'] == '合计',
            ['水费', '违约金', '税额', '污水费', '垃圾处理费', '预收款收支']].values.tolist()[0]
    ]
    df_2 = df.copy(deep=True)
    df_2['原币金额'] = amount_df_2_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = amount_df_2_all
    df_2['摘要'] = [
        f'收到{last_month}月水费',
        '水费违约金',
        '水费违约金销项税金',
        '代收污水处理费',
        '代收垃圾处理费',
        '收水费预收款',
    ]
    df_2['科目'] = [
        '1122.001',
        '6301.003',
        '2221.016.002',
        '2241.003.001',
        '2241.003.002',
        '2203.004',
    ]
    df_2['科目名称'] = [
        '应收账款_自来水',
        '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%',
        '其他应付款_外部单位往来款_污水费',
        '其他应付款_外部单位往来款_垃圾费',
        '预收账款_水费',
    ]
    df_2['核算项目1'] = ['', '行政组织', '', '供应商', '供应商', '']
    df_2['编码1'] = [
        '', '2.01.01.01.01.22.01', '', 'G2.21.000326', 'G2.21.000319', ''
    ]
    df_2['名称1'] = ['', '城区分公司客服部', '', '中山市建设局', '中山市环境卫生管理处', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00003'
    print(df_sum)
    df_sum.to_excel(save_dir + "\\" + f'{type1}_xianjinshuifei.xlsx',
                    index=False)
    return df_sum


def genExcel_chengqu_yinhanghuazhang(excelspath_sheet28_of_dir, save_dir):
    # 城区-银行划账--->城区
    # save_dir 末端不加反斜杠
    # excelspath_sheet28_of_dir:报表二十八s所在的文件路径

    if os.path.exists(save_dir + '\\chengqu_yinhanghuazhang') == False:
        os.mkdir(save_dir + '\\chengqu_yinhanghuazhang')
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
                                     '实收金额'].tolist()
        df_1['方向'] = '1'  # 借方
        df_1['借方金额'] = df_data_1.loc[df_data_1['项目'] == '总金额',
                                     '实收金额'].tolist()
        df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
            '收到工商银行代收水费', '1002', '银行存款', '银行账户', '2.01.003',
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
            '收到工商银行代收水费', '代收污水费', '代收垃圾费', '水费违约金', '水费违约金销项税', '收到预收水费'
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
            '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.22.01', '', ''
        ]
        df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '城区分公司客服部', '', '']
        df_sum = df_sum.append(df_2, ignore_index=True)

        df_sum = constant_value(df_sum)
        # df_sum['凭证号'] = ''
        print(df_sum)
        df_sum.to_excel(save_dir + '\\chengqu_yinhanghuazhang' + '\\' +
                        f'chengqu_yinhanghuazhang{acountID}.xlsx',
                        index=False)
        # return df_sum


def genExcel_banfu_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-板芙--->支付宝-板芙
    # save_dir 末端不加反斜杠
    # excelpath_sheet20:报表二十
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '板芙', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[[
        '摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1'
    ]] = '代板芙分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001', "中国民生银行中山分行营业部'691990019"
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代板芙分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000016', "中山公用水务有限公司三乡分公司"
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00005'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'banfu_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_dongfeng_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-东凤--->城区
    # save_dir 末端不加反斜杠
    # excelpath_sheet20:报表二十
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '东凤', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代东凤分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代东凤分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000017', '中山公用水务有限公司东凤分公司'
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00006'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'dongfeng_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_dongsheng_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-东升--->城区
    # save_dir 末端不加反斜杠
    # excelpath_sheet20:报表二十
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '东升', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代东升分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代东升分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000018', '中山公用水务有限公司东升分公司'
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00007'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'dongsheng_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_guzhen_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-古镇--->城区
    # excelpath_sheet20:报表二十
    # save_dir 末端不加反斜杠
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '古镇', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代古镇分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代古镇分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000015', '中山公用水务有限公司古镇分公司'
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00008'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'guzhen_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_hannanqu_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-含南区--->城区
    # excelpath_sheet20:报表二十
    # save_dir 末端不加反斜杠
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2) for each in df_data_1.loc[(df_data_1['区域'] == '南区')
                                                 | (df_data_1['区域'] == '城区')
                                                 | (df_data_1['区域'] == '五桂山'),
                                                 '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    # 城区
    amount_df_chengqu_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '城区', '水费'].tolist()
    ]
    amount_df_chengqu_2_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '城区', '污水费'].tolist()
    ]
    amount_df_chengqu_3_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '城区', '垃圾费'].tolist()
    ]
    amount_df_chengqu_4_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '城区', '违约金'].tolist()
    ]
    amount_df_chengqu_6_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '城区', '重收金额'].tolist()
    ]
    sum_amount_chengqu_1_all = round(sum(amount_df_chengqu_1_all), 2)
    sum_amount_chengqu_2_all = round(sum(amount_df_chengqu_2_all), 2)
    sum_amount_chengqu_3_all = round(sum(amount_df_chengqu_3_all), 2)
    sum_amount_chengqu_4_all = round(sum(amount_df_chengqu_4_all), 2) - round(
        sum(amount_df_chengqu_4_all), 2) * 0.06
    sum_amount_chengqu_5_all = round(sum(amount_df_chengqu_4_all), 2) * 0.06
    sum_amount_chengqu_6_all = round(sum(amount_df_chengqu_6_all), 2)
    sum_amount_chengqu_all = [
        sum_amount_chengqu_1_all, sum_amount_chengqu_2_all,
        sum_amount_chengqu_3_all, sum_amount_chengqu_4_all,
        sum_amount_chengqu_5_all, sum_amount_chengqu_6_all
    ]
    df_2['原币金额'] = sum_amount_chengqu_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_chengqu_all
    df_2['摘要'] = [
        '收水费（支付宝）',
        '代收污水费',
        '代收垃圾费',
        '收水费违约金',
        '水费违约金销项税',
        '收到预收水费',
    ]
    df_2['科目'] = [
        '1122.001',
        '2241.003.001',
        '2241.003.002',
        '6301.003',
        '2221.016.002',
        '2203.004',
    ]
    df_2['科目名称'] = [
        '应收账款_自来水',
        '其他应付款_外部单位往来款_污水费',
        '其他应付款_外部单位往来款_垃圾费',
        '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%',
        '预收账款_水费',
    ]
    df_2['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
    df_2['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.22.01', '', ''
    ]
    df_2['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '城区分公司客服部', '', '']
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_3 = df.copy(deep=True)
    # 五桂山
    amount_df_wuguishan_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '五桂山', '水费'].tolist()
    ]
    amount_df_wuguishan_2_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '五桂山', '污水费'].tolist()
    ]
    amount_df_wuguishan_3_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '五桂山', '垃圾费'].tolist()
    ]
    amount_df_wuguishan_4_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '五桂山', '违约金'].tolist()
    ]
    amount_df_wuguishan_6_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '五桂山', '重收金额'].tolist()
    ]
    sum_amount_wuguishan_1_all = round(sum(amount_df_wuguishan_1_all), 2)
    sum_amount_wuguishan_2_all = round(sum(amount_df_wuguishan_2_all), 2)
    sum_amount_wuguishan_3_all = round(sum(amount_df_wuguishan_3_all), 2)
    sum_amount_wuguishan_4_all = round(
        sum(amount_df_wuguishan_4_all),
        2) - round(sum(amount_df_wuguishan_4_all), 2) * 0.06
    sum_amount_wuguishan_5_all = round(sum(amount_df_wuguishan_4_all),
                                       2) * 0.06
    sum_amount_wuguishan_6_all = round(sum(amount_df_wuguishan_6_all), 2)
    sum_amount_wuguishan_all = [
        sum_amount_wuguishan_1_all, sum_amount_wuguishan_2_all,
        sum_amount_wuguishan_3_all, sum_amount_wuguishan_4_all,
        sum_amount_wuguishan_5_all, sum_amount_wuguishan_6_all
    ]
    df_3['原币金额'] = sum_amount_wuguishan_all
    df_3['方向'] = '0'  # 贷方
    df_3['贷方金额'] = sum_amount_wuguishan_all
    df_3['摘要'] = [
        '收水费（支付宝） 五桂山',
        '代收污水费 五桂山',
        '代收垃圾费 五桂山',
        '收水费违约金 五桂山',
        '水费违约金销项税 五桂山',
        '收到预收水费 五桂山',
    ]
    df_3['科目'] = [
        '1122.001',
        '2241.003.001',
        '2241.003.002',
        '6301.003',
        '2221.016.002',
        '2203.004',
    ]
    df_3['科目名称'] = [
        '应收账款_自来水',
        '其他应付款_外部单位往来款_污水费',
        '其他应付款_外部单位往来款_垃圾费',
        '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%',
        '预收账款_水费',
    ]
    df_3['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
    df_3['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.22.01', '', ''
    ]
    df_3['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '城区分公司客服部', '', '']
    df_sum = df_sum.append(df_3, ignore_index=True)

    df_4 = df.copy(deep=True)
    # 南区
    amount_df_nanqu_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '南区', '水费'].tolist()
    ]
    amount_df_nanqu_2_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '南区', '污水费'].tolist()
    ]
    amount_df_nanqu_3_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '南区', '垃圾费'].tolist()
    ]
    amount_df_nanqu_4_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '南区', '违约金'].tolist()
    ]
    amount_df_nanqu_6_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '南区', '重收金额'].tolist()
    ]
    sum_amount_nanqu_1_all = round(sum(amount_df_nanqu_1_all), 2)
    sum_amount_nanqu_2_all = round(sum(amount_df_nanqu_2_all), 2)
    sum_amount_nanqu_3_all = round(sum(amount_df_nanqu_3_all), 2)
    sum_amount_nanqu_4_all = round(sum(amount_df_nanqu_4_all), 2) - round(
        sum(amount_df_nanqu_4_all), 2) * 0.06
    sum_amount_nanqu_5_all = round(sum(amount_df_nanqu_4_all), 2) * 0.06
    sum_amount_nanqu_6_all = round(sum(amount_df_nanqu_6_all), 2)
    sum_amount_nanqu_all = [
        sum_amount_nanqu_1_all, sum_amount_nanqu_2_all, sum_amount_nanqu_3_all,
        sum_amount_nanqu_4_all, sum_amount_nanqu_5_all, sum_amount_nanqu_6_all
    ]
    df_4['原币金额'] = sum_amount_nanqu_all
    df_4['方向'] = '0'  # 贷方
    df_4['贷方金额'] = sum_amount_nanqu_all
    df_4['摘要'] = [
        '收水费（支付宝） 南区',
        '代收污水费 南区',
        '代收垃圾费 南区',
        '收水费违约金 南区',
        '水费违约金销项税 南区',
        '收到预收水费 南区',
    ]
    df_4['科目'] = [
        '1122.001',
        '2241.003.001',
        '2241.003.002',
        '6301.003',
        '2221.016.002',
        '2203.004',
    ]
    df_4['科目名称'] = [
        '应收账款_自来水',
        '其他应付款_外部单位往来款_污水费',
        '其他应付款_外部单位往来款_垃圾费',
        '营业外收入_违约金收入',
        '应交税费_简易计税_简易计税3%',
        '预收账款_水费',
    ]
    df_4['核算项目1'] = ['', '供应商', '供应商', '行政组织', '', '']
    df_4['编码1'] = [
        '', 'G2.21.000326', 'G2.21.000319', '2.01.01.01.01.25', '', ''
    ]
    df_4['名称1'] = ['', '中山市建设局', '中山市环境卫生管理处', '南区营业厅', '', '']
    df_sum = df_sum.append(df_4, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = ''
    print(df_sum)
    df_sum.to_excel(save_dir + "\\" + 'hannanqu_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_sanxiang_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-三乡--->城区
    # excelpath_sheet20:报表二十
    # save_dir 末端不加反斜杠
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '三乡', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代三乡分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代三乡分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000016', '中山公用水务有限公司三乡分公司'
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00010'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'sanxiang_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_shenwan_zhifubao(excelpath_sheet20, save_dir):
    # 支付宝-神湾--->城区
    # excelpath_sheet20:报表二十
    # save_dir 末端不加反斜杠
    # 把定值全表深拷贝一份最为初始
    df_sum = df.copy(deep=True)

    df_data_1 = pd.read_excel(excelpath_sheet20)  # 报表二十
    df_1 = df.copy(deep=True)
    amount_df_1_all = [
        round(each, 2)
        for each in df_data_1.loc[df_data_1['区域'] == '神湾', '到账总金额'].tolist()
    ]
    df_1['原币金额'] = amount_df_1_all
    df_1['方向'] = '1'  # 借方
    df_1['借方金额'] = amount_df_1_all
    df_1[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代神湾分公司收水费（支付宝）', '1002', '银行存款', '银行账户', '2.16.001',
        "中国民生银行中山分行营业部'691990019"
    ]
    df_sum = df_sum.append(df_1, ignore_index=True)

    df_2 = df.copy(deep=True)
    sum_amount_df_1_all = [round(sum(amount_df_1_all), 2)]
    df_2['原币金额'] = sum_amount_df_1_all
    df_2['方向'] = '0'  # 贷方
    df_2['贷方金额'] = sum_amount_df_1_all
    df_2[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
        '代神湾分公司收水费（支付宝）', '1221.002.013.001', '其他应收款_关联单位往来款_代收款_水费', '客户',
        'KA.10.000016', '中山公用水务有限公司三乡分公司'
    ]
    df_sum = df_sum.append(df_2, ignore_index=True)

    df_sum = constant_value(df_sum)
    # df_sum['凭证号'] = '00011'
    print(df_sum)
    df_sum.to_excel(save_dir + '\\' + 'shenwan_zhifubao.xlsx', index=False)
    return df_sum


def genExcel_chengqu_check(excelpath_sheet5, save_dir, type1):
    # 竹苑支票--->城区
    # excelpath_sheet5:报表五
    # save_dir 末端不加反斜杠
    # type1 城区营业厅  str     zhuyuan or liantang

    if os.path.exists(save_dir + f'\\{type1}_check') == False:
        os.mkdir(save_dir + f'\\{type1}_check')
    df_data_1 = pd.read_excel(excelpath_sheet5,
                              sheet_name='营业厅收费日报_支票_汇总',
                              header=None)  # 报表五
    # print(df_data_1.iloc[:, 0].tolist())
    header_rows = []
    sum_rows = []
    for index in range(len(df_data_1.iloc[:, 0].tolist())):
        col1_content = df_data_1.iloc[:, 0].tolist()[index]
        if col1_content == '户号':
            header_rows.append(index)
        if col1_content == '合计':
            sum_rows.append(index)
    # print(header_rows)
    # print(sum_rows)
    for header in header_rows:
        # 把定值全表深拷贝一份最为初始
        df_sum = df.copy(deep=True)
        df_data_1_eachitem = pd.read_excel(excelpath_sheet5,
                                           sheet_name='营业厅收费日报_支票_汇总',
                                           header=header)
        df_data_1_eachitem = df_data_1_eachitem.iloc[
            0:sum_rows[header_rows.index(header)] - header, ]
        # print(df_data_1_eachitem)
        customName = df_data_1_eachitem.iloc[0, ]['户名']
        amount1_eachitem = [
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '金额'].tolist()
        ]
        amount2_eachitem = [
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '实收金额'].tolist()
        ]
        wushuifei_eachitem = [
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '污水费'].tolist()
        ]
        lajichulifei_eachitem = [
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '垃圾处理费'].tolist()
        ]
        yushoukuanshouzhi_eachitem = [
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '预收款收支'].tolist()
        ]
        weiyuejin_eachitem = [[
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '滞纳金'].tolist()
        ][0] * (1 - 0.06)]
        shuie_eachitem = [[
            round(i, 2)
            for i in df_data_1_eachitem.loc[df_data_1_eachitem['户号'] == '合计',
                                            '滞纳金'].tolist()
        ][0] * 0.06]
        if amount1_eachitem[0] < 0:  # 合计金额小于零，跳过，不需要记账。
            continue
        else:
            # print(customName)
            # print(amount1_eachitem)  # amount1_list# 金额
            # print(amount2_eachitem)  # amount2_list# 实收金额
            # print(wushuifei_eachitem)
            # print(lajichulifei_eachitem)
            # print(yushoukuanshouzhi_eachitem)
            # print(weiyuejin_eachitem)
            # print(shuie_eachitem)

            df_1 = df.copy(deep=True)
            df_1['原币金额'] = amount1_eachitem
            df_1['方向'] = '1'  # 借方
            df_1['借方金额'] = amount1_eachitem
            df_1['摘要'] = f'收到{customName}水费'
            df_1[['科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
                '1002', '银行存款', '银行账户', '2.01.003',
                "工商行香山支行'2011002109022109510"
            ]
            df_sum = df_sum.append(df_1, ignore_index=True)

            df_2 = df.copy(deep=True)
            df_2['原币金额'] = amount2_eachitem
            df_2['方向'] = '0'  # 贷方
            df_2['贷方金额'] = amount2_eachitem
            df_2['摘要'] = f'收到{customName}水费'
            df_2[['科目', '科目名称', '核算项目1', '编码1',
                  '名称1']] = ['1122.001', '应收账款_自来水', '', '', ""]
            df_sum = df_sum.append(df_2, ignore_index=True)

            # 代收污水处理费
            df_3 = df.copy(deep=True)
            df_3['原币金额'] = wushuifei_eachitem
            df_3['方向'] = '0'  # 贷方
            df_3['贷方金额'] = wushuifei_eachitem
            df_3[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
                '代收污水处理费', '2241.003.001', '其他应付款_外部单位往来款_污水费', '供应商',
                'G2.21.000326', "中山市建设局"
            ]
            df_sum = df_sum.append(df_3, ignore_index=True)

            # 代收垃圾处理费
            df_4 = df.copy(deep=True)
            df_4['原币金额'] = lajichulifei_eachitem
            df_4['方向'] = '0'  # 贷方
            df_4['贷方金额'] = lajichulifei_eachitem
            df_4[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
                '代收垃圾处理费', '2241.003.002', '其他应付款_外部单位往来款_垃圾费', '供应商',
                'G2.21.000319', "中山市环境卫生管理处"
            ]
            df_sum = df_sum.append(df_4, ignore_index=True)

            # 收到预收水费
            df_5 = df.copy(deep=True)
            df_5['原币金额'] = yushoukuanshouzhi_eachitem
            df_5['方向'] = '0'  # 贷方
            df_5['贷方金额'] = yushoukuanshouzhi_eachitem
            df_5[['摘要', '科目', '科目名称', '核算项目1', '编码1',
                  '名称1']] = ['收到预收水费', '2203.004', '预收账款_水费', '', '', ""]
            df_sum = df_sum.append(df_5, ignore_index=True)

            # 收水费违约金
            df_6 = df.copy(deep=True)
            df_6['原币金额'] = weiyuejin_eachitem
            df_6['方向'] = '0'  # 贷方
            df_6['贷方金额'] = weiyuejin_eachitem
            df_6[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
                '收水费违约金', '6301.003', '营业外收入_违约金收入', '行政组织',
                '2.01.01.01.01.22.01', "城区分公司客服部"
            ]
            df_sum = df_sum.append(df_6, ignore_index=True)

            # 水费违约金销项税
            df_7 = df.copy(deep=True)
            df_7['原币金额'] = shuie_eachitem
            df_7['方向'] = '0'  # 贷方
            df_7['贷方金额'] = shuie_eachitem
            df_7[['摘要', '科目', '科目名称', '核算项目1', '编码1', '名称1']] = [
                '水费违约金销项税', '2221.016.002', '应交税费_简易计税_简易计税3%', '', '', ""
            ]
            df_sum = df_sum.append(df_7, ignore_index=True)

            df_sum = constant_value(df_sum)
            # df_sum['凭证号'] = '00012'
            print(df_sum)
            df_sum.to_excel(save_dir + f'\\{type1}_check' + '\\' +
                            f'{type1}_check_{customName}.xlsx',
                            index=False)
    # return df_sum


if __name__ == "__main__":
    pass
    genExcel_chengqu_shuaka(
        excelpath_sheet15=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\扫码-刷卡\2021-10-13\竹苑-刷卡\竹苑-刷卡.xlsx',
        excelpath_sheet2=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\竹苑\营业厅收费汇总报表\db_营业厅收费汇总报表.xlsx',
        excelpath_sheet18=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\新建文件夹\Robot_Files\刷卡汇总\2021-10-13\竹苑-刷卡\刷卡汇总.xlsx',
        save_dir=r'.\pingzheng',
        type1='zhuyuan')
    # todo 可用于莲塘

    genExcel_chengqu_saoma(
        excelpath_sheet16=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\扫码-刷卡\2021-10-13\竹苑-扫码\竹苑-扫码.xlsx',
        excelpath_sheet3=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\竹苑\营业厅收费汇总报表\db_营业厅收费汇总报表.xlsx',
        excelpath_sheet19=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\新建文件夹\Robot_Files\扫码汇总\2021-10-13\竹苑-扫码\扫码汇总.xlsx',
        save_dir=r'.\pingzheng',
        type1='zhuyuan')
    # todo 可用于莲塘

    genExcel_chengqu_xianjinshuifei(
        excelpath_sheet1=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\竹苑\营业厅收费汇总报表\db_营业厅收费汇总报表.xlsx',
        save_dir=r'.\pingzheng',
        type1='zhuyuan')
    # todo 可用于莲塘

    genExcel_chengqu_yinhanghuazhang(
        excelspath_sheet28_of_dir=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\客服大厅\划帐情况汇总',
        save_dir=r'.\pingzheng')

    genExcel_banfu_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_dongfeng_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_dongsheng_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_guzhen_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_hannanqu_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_sanxiang_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_shenwan_zhifubao(
        excelpath_sheet20=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\支付宝\远程销账汇总_支付宝美宜佳\db_远程销账汇总_支付宝_不合并.xlsx',
        save_dir=r'.\pingzheng')

    genExcel_chengqu_check(
        excelpath_sheet5=
        r'F:\zhongshan_shuiwu_RPA\20211013\voucher\data\2021-10-13\竹苑\营业厅收费日报_支票\db_营业厅收费日报_支票.xlsx',
        save_dir=r'.\pingzheng',
        type1='zhuyuan')
    # todo 可用于莲塘
