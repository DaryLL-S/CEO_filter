import re

import numpy as np
import pandas as pd

# 读取数据
df = pd.read_excel('./data/CEO2-学生待检查版本.xlsx')
company_array = np.array(df['stkcd'].drop_duplicates())
company_list = company_array.tolist()

# 对日期数据进行处理，生成年、月、日三个新列，以便后续判断
df['year'] = pd.DatetimeIndex(df['Reptdt']).year
df['month'] = pd.DatetimeIndex(df['Reptdt']).month
df['day'] = pd.DatetimeIndex(df['Reptdt']).day

true_df = pd.DataFrame(columns=df.columns)
false_df = pd.DataFrame(columns=df.columns)

# 具体思路：任务目标是在同一个公司同一年里确定position最高的数据，其余数据均剔除，即仅需对每个公司每一年进行单独分析
# 按公司划分
for i in company_list:
    data_1 = df.loc[df['stkcd'] == i]
    year_array = np.array(data_1['year'].drop_duplicates())
    year_list = year_array.tolist()

    truelist = np.array([h for h in range(len(data_1))])
    falselist = []

    # 按年度划分
    for j in year_list:
        data_2 = data_1.loc[data_1['year'] == j]
        namelist = []
        CEO_list = []
        chairman_list = []
        Manager_list = []
        else_list = []
        for k in range(len(data_2)):
            row = data_2.iloc[k].values.tolist()
            print(row[0], row[1], row[3], row[4], row[5], row[9], row[44])

            # 剔除同公司同一年的重复数据
            if row[1] not in namelist:
                namelist.append(row[1])
            else:
                falselist.append(namelist.index(row[1]))

            # 按优先级匹配position，先找优先级高的，找不到再找低优先级的
            if re.search('首席执行官|执行官|行政总裁|CEO', row[3]) is not None:
                CEO_list.append(k)
            elif re.search('总裁', row[3]) is not None:
                chairman_list.append(k)
            elif re.search('总经理', row[3]) is not None:
                Manager_list.append(k)
            else:
                else_list.append(k)

        # 确定在当前年份里，同一家公司按时间往前还有无其他数据
        length = 0
        while year_list.index(j) > 0:
            j = year_list[(year_list.index(j) - 1)]
            length += len(data_1.loc[data_1['year'] == j])

        # 当有较高级的职务存在时，剔除其他低级职位及不相关职位（包含所有报告期不在12月31号的数据）
        if CEO_list:
            if chairman_list:
                for x in chairman_list:
                    falselist.append(x + length)
            if Manager_list:
                for x in Manager_list:
                    falselist.append(x + length)
            if else_list:
                for x in else_list:
                    falselist.append(x + length)
            for x in CEO_list:
                row = data_2.iloc[x].values.tolist()
                if not (row[9] == 12 and row[44] == 31):
                    falselist.append(x + length)
        elif chairman_list:
            if Manager_list:
                for x in Manager_list:
                    falselist.append(x + length)
            if else_list:
                for x in else_list:
                    falselist.append(x + length)
            for x in chairman_list:
                row = data_2.iloc[x].values.tolist()
                if not (row[9] == 12 and row[44] == 31):
                    falselist.append(x + length)
        elif Manager_list:
            if else_list:
                for x in else_list:
                    falselist.append(x + length)
            for x in Manager_list:
                row = data_2.iloc[x].values.tolist()
                if not (row[9] == 12 and row[44] == 31):
                    falselist.append(x + length)
        elif else_list:
            for x in else_list:
                falselist.append(x + length)

        print(CEO_list, chairman_list, Manager_list, else_list)  # 输出每一年各职位的列表

    print(falselist)  # 输出当前公司中要剔除的数据索引

    for g in range(len(truelist)):
        if g not in falselist:
            true_df = true_df.append(data_1.iloc[g], ignore_index=False)
        else:
            false_df = false_df.append(data_1.iloc[g], ignore_index=False)

# 保存数据
write = pd.ExcelWriter("./result/CEO3.xlsx")
true_df.to_excel(write, sheet_name='保留', index=False)
false_df.to_excel(write, sheet_name='删除', index=False)
write.save()
