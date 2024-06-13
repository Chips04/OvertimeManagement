import pandas as pd
from datetime import timedelta
from 已完成.functions import *

df = pd.read_excel('/home/sf107/桌面/值班表格处理/值班明细表_20240116145922.xlsx')
df2 = pd.read_excel('/home/sf107/桌面/值班表格处理/值班分组表_20240218141954.xlsx')

# 处理合并的表头
columns = ['组名', '值班领导', '值班领导电话', '组员', '组员姓名', '电话']
df2.columns = columns  # 重新设置列名
df2 = df2.drop([0])   # 删除第一行数据

# 原输出表最后一行
lastGroup = df.at[df.__len__() - 1, '值班组名']  # 原表最后一个值班组
lastDate = df.at[df.__len__() - 1, '值班日期日期格式']  # 原表最后一个值班日期日期格式


# 确定组名第一个的排序
groupsSeries = df2['组名']
groupsSeries = groupsSeries.dropna()  # 去掉空值
groupsArr = groupsSeries.tolist()

# 处理list排序
list1 = groupsArr[:groupsArr.index(lastGroup) + 1]
list2 = groupsArr[-(len(groupsArr) - groupsArr.index(lastGroup) - 1):]
groupsArrNew = groupsArr if groupsArr.index(lastGroup) == 15 else list2 + list1

column_names = df.columns.tolist()  # 列名list

df2 = df2.fillna(method='ffill')  # 填充表格空值

weekday_str = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
# moreDuty 生成最终输出表格
# date 排班到几号 日期格式为："2024-08-31"
def md(date, dataframe):
    dfLen = len(dataframe)
    dutyDate = lastDate
    formatDate = date + " 00:00:00"
    timestamp = pd.to_datetime(formatDate)
    delta = timedelta(days=1)
    while dutyDate < timestamp:
        for groupName in groupsArrNew:
            if dutyDate >= timestamp:
                break
            groupLeaderHasDone = False
            dutyDate = dutyDate + delta
            for index, row in df2.iterrows():
                if row['组名'] == groupName:
                    # 组员行
                    newRow2 = pd.Series(
                        [
                            row['组名'],
                            row['组员姓名'],
                            '值班组员',
                            row['电话'],
                            str(int(row['电话'])),
                            dutyDate.strftime('%Y-%m-%d'),
                            dutyDate,
                            weekday_str[dutyDate.weekday()],
                            row['值班领导'],
                            row['值班领导电话'],
                            28039061,
                            0
                        ],
                        index=column_names
                    )
                    dataframe = dataframe._append(newRow2, ignore_index=True)
                    if groupLeaderHasDone == False:
                        # 值班组长行
                        newRow = pd.Series(
                            [
                                row['组名'],
                                row['值班领导'],
                                '值班领导',
                                row['值班领导电话'],
                                str(int(row['值班领导电话'])),
                                dutyDate.strftime('%Y-%m-%d'),
                                dutyDate,
                                weekday_str[dutyDate.weekday()],
                                row['值班领导'],
                                row['值班领导电话'],
                                28039061,
                                0
                            ],
                            index=column_names
                        )
                        dataframe = dataframe._append(newRow, ignore_index=True)
                        groupLeaderHasDone = True
    dataframe = dataframe.drop(dataframe.index[:dfLen], axis=0)  # 删除原来数据
    df_reset = dataframe.reset_index()  # 重置index
    # 替换重名
    rn(df_reset, '值班人员')
    # 输出表格
    df_reset.to_excel(('/home/sf107/桌面/值班表格处理/值班输出表格.xlsx'), index=False)


# 主函数 第一个是排班到几号日期
md("2024-12-31", df)

