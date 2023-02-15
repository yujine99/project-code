# _*_ coding: utf-8 _*_
# @Time   : 2022/4/27 下午 04:40
# @Author : Carolyn_yu
# @FileName : 人力筛选统计.py
# @Software : PyCharm
# @Blog     :
# _*_ coding: utf-8 _*_
# @Time   : 2022/4/27 下午 02:29
# @Author : Carolyn_yu
# @FileName : Python对Excel多条件筛选统计.py
# @Software : PyCharm
# @Blog     :
import pandas as pd
import numpy as np

df = pd.read_excel('20220301 2.xlsx', 'Sheet1')
div = np.unique(df['支援部門代碼'])
print(div)
d = {}
divis = []
m1 = []
m2 = []
m3 = []
m4 = []
m5 = []
m6 = []
m7 = []

for division in div:
    print(division)
    C1 = df.loc[(df['支援部門代碼'] == division) & (df['區分'].isin([1, 2])) & (df['班別'].str.contains('白班')) & (
        ~df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C1)
    print(len(C1))
    C2 = df.loc[(df['支援部門代碼'] == division) & (df['區分'].isin([1, 2])) & (df['班別'].str.contains('白班')) & (
        df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C2)
    print(len(C2))
    C3 = df.loc[(df['支援部門代碼'] == division) & (df['區分'].isin([1, 2])) & (~df['班別'].str.contains('白班')) & (
        ~df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C3)
    print(len(C3))
    C4 = df.loc[(df['支援部門代碼'] == division) & (df['區分'].isin([1, 2])) & (~df['班別'].str.contains('白班')) & (
        df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C4)
    print(len(C4))
    C5 = df.loc[(df['支援部門代碼'] == division) & (~df['區分'].isin([1, 2])) & (df['班別'].str.contains('白班')) & (
        ~df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C5)
    print(len(C5))
    C6 = df.loc[(df['支援部門代碼'] == division) & (~df['區分'].isin([1, 2])) & (df['班別'].str.contains('白班')) & (
        df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C6)
    print(len(C6))
    C7 = df.loc[(df['支援部門代碼'] == division) & (~df['區分'].isin([1, 2])) & (~df['班別'].str.contains('白班')) & (
        ~df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C7)
    print(len(C7))
    C8 = df.loc[(df['支援部門代碼'] == division) & (~df['區分'].isin([1, 2])) & (~df['班別'].str.contains('白班')) & (
        df['異常原因'].isin(['請假'])), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
    print(C8)
    print(len(C8))

# for division in div:
#     for zhiden in np.unique(df['區分']):
#         C0 = df.loc[df['支援部門代碼'] == division]
#         C1 = df.loc[(df['支援部門代碼'] == division) & (df['區分'] == ['1|2']), ['支援部門代碼', '區分', '班別', '異常原因', '姓名']]
#         print(C1)
#         d = {zhiden: len(C1)}
#         print(d)

# C0 = df.loc[df["省份"] == Province]
#     C1 = df.loc[(df["省份"] == Province) & (df['组合'].str.contains('SCR')) & (df['组合'].str.contains('SNCR')) & (
#         df['组合'].str.contains('SFGD|DFGD')) & (df['组合'].str.contains('WFGD'))]
#     C2 = df.loc[(df["省份"] == Province) & (df['组合'].str.contains('SCR')) & (df['组合'].str.contains('SNCR')) & (
#         df['组合'].str.contains('SFGD|DFGD')) & (~df['组合'].str.contains('WFGD'))]
#     C3 = df.loc[(df["省份"] == Province) & (df['组合'].str.contains('SNCR')) & (df['组合'].str.contains('SFGD|DFGD')) & (
#         df['组合'].str.contains('WFGD')) & (~df['组合'].str.contains('SCR'))]
#     C4 = df.loc[(df["省份"] == Province) & (df['组合'].str.contains('SNCR')) & (df['组合'].str.contains('SFGD|DFGD')) & (
#         ~df['组合'].str.contains('SCR|WFGD'))]
#     C5 = df.loc[
#         (df["省份"] == Province) & (df['组合'].str.contains('SFGD|DFGD')) & (~df['组合'].str.contains('SNCR|SCR|WFGD'))]
#     C6 = df.loc[(df["省份"] == Province) & (df['组合'].str.contains('SFGD')) & (df['组合'].str.contains('WFGD')) & (
#         ~df['组合'].str.contains('SNCR|SCR|DFGD'))]
#     # ~是不包含 |或的关系
#     L0 = len(C0)
#     L1 = len(C1)
#     L2 = len(C2)
#     L3 = len(C3)
#     L4 = len(C4)
#     L5 = len(C5)
#     L6 = len(C6)
#     L7 = L0 - (L1 + L2 + L3 + L4 + L5 + L6)
#     pro.append(Province)
#     m1.append(L1)
#     m2.append(L2)
#     m3.append(L3)
#     m4.append(L4)
#     m5.append(L5)
#     m6.append(L6)
#     m7.append(L7)
# d = {'省份': pro, 'SNCR+SCR+SFGD/DFGD+WFGD': m1, 'SNCR+SCR+SFGD/DFGD': m2, 'SNCR+SFGD/DFGD+WFGD': m3,
#      'SNCR+SFGD/DFGD': m4, 'SFGD/DFGD': m5, 'SFGD+WFGD': m6, '其他': m7}
# result = pd.DataFrame(d)
# writer = pd.ExcelWriter('组合设施01.xlsx')
# result.to_excel(writer, 'Sheet1')
# writer.save()
# print('Over')
