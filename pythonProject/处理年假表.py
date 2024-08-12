import openpyxl
import pandas as pd
import os
import fnmatch
import numpy as np
import json

print('正在处理……')
# 读取txt文件内容
# with open('variables.txt', 'r', encoding='utf-8') as file:
#     content = file.read().strip()
# 读取文件内容，并手动处理可能的 UTF-8 BOM
with open('variables.txt', 'rb') as file:  # 使用二进制模式打开文件
    content = file.read()
    # 检查并去除 UTF-8 BOM（如果有的话）
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]  # 跳过 BOM
    content = content.decode('utf-8').strip()  # 解码为 UTF-8 并去除空白字符

# 尝试将内容转换为字典，注意这里假设了txt文件的内容是合法的JSON格式
try:
    data = json.loads(content)  # 将字符串转换为字典
except json.JSONDecodeError as e:
    print(f"无法解析txt文件内容为JSON: {e}")
    exit(1)

# 从字典中提取变量
directory = data['path']
group = data['group']

# 获取数据源dataframe
df_r = pd.DataFrame()

pattern2 = '社区工作人员花名册*.xlsx'
# 获取匹配的文件列表
matches2 = []
for filename2 in os.listdir(directory):
    if fnmatch.fnmatch(filename2, pattern2):
        matches2.append(os.path.join(directory, filename2))

# 打开第一个匹配的文件并读取内容
if matches2:
    df_r = pd.read_excel(matches2[0])
else:
    print('No matching files found2.')

# 获取数据源dataframe
df3 = pd.DataFrame()
# 指定目录和通配符
pattern3 = '人员管理-请休假*.xlsx'
# 获取匹配的文件列表
matches3 = []
for filename3 in os.listdir(directory):
    if fnmatch.fnmatch(filename3, pattern3):
        matches3.append(os.path.join(directory, filename3))

# 打开第一个匹配的文件并读取内容
if matches3:
    df3 = pd.read_excel(matches3[0])
else:
    print('No matching files found.')

wb_origin8 = openpyxl.load_workbook(directory + '模板/年假统计表模板.xlsx')  # ！！！
ws_origin8 = wb_origin8['2024']

# 选组
df = df_r[df_r['加班费组别（加班费使用）'] == group].copy()
# df = df_r.copy()

df['参加工作时间'] = df['入职时间'].apply(lambda x: f"{x.year}.{x.strftime('%m')}" if pd.notnull(x) else None)
# df['参加工作时间'] = df['入职时间'].apply(lambda x: f'{x.year}.{x.month:02d}')
df['假期类别'] = '年假'
names_to_remove = ['陈学荣', '苏伟如', '戴南真', '王大森']
mask = ~df['工作人员姓名'].isin(names_to_remove)
df = df[mask]
df = df.reset_index(drop=True)  # 重置索引
df['序号'] = df.index + 1  # 将索引值加1作为序号列
df['姓名'] = df['工作人员姓名']
merged_df = df[['序号', '姓名', '参加工作时间', '假期类别', '上一年剩余年假', '今年可休年假总天数']].copy()
# merged_df = pd.merge(df, df_r[['工作人员姓名', '上一年剩余年假', '今年可休年假总天数']], left_on='姓名', right_on='工作人员姓名', how='left')

merged_df['2023剩余'] = merged_df['上一年剩余年假']
merged_df['可休假总天数(2024)'] = merged_df['今年可休年假总天数']

BBB = 5

# 删除'旧列'
del merged_df['上一年剩余年假']
del merged_df['今年可休年假总天数']
# del merged_df['工作人员姓名']
# 设置显示的列数
pd.set_option('display.max_columns', None)  # 显示所有列
# 设置每列的最大宽度
pd.set_option('display.max_colwidth', 100)  # 例如，设置为100个字符
# merged_df['可休假总天数'] = f'=O{}+P{}'
merged_df['可休假总天数'] = merged_df.index.map(lambda x: '=O' + str(x + BBB) + "+P" + str(x + BBB))
merged_df['可休假总天数（剩余假期）'] = merged_df.index.map(lambda x: '=Q' + str(x + BBB) + "-R" + str(x + BBB))
new_columns = [f'NewCol{i+1}' for i in range(10)]
existing_columns = merged_df.columns.tolist()
columns_with_new = existing_columns[:4] + new_columns + existing_columns[4:]
merged_df = merged_df.reindex(columns=columns_with_new)
merged_df.insert(loc=merged_df.columns.get_loc('可休假总天数') + 1, column='已休假总天数', value=np.nan)

# 请休假表格处理
# 只保留年假的行
df3 = df3[df3['休假类别'] == '年休假']
df3 = df3.sort_values(by=['申请人员姓名', '请假开始日期', '请假开始时段'])
df3 = df3.reset_index(drop=True)
# print(df3)
merged_df = merged_df.rename(
    columns={merged_df.columns[4]: '休假1', merged_df.columns[5]: '休假2', merged_df.columns[6]: '休假3', merged_df.columns[7]: '休假4',
             merged_df.columns[8]: '休假5', merged_df.columns[9]: '休假6', merged_df.columns[10]: '休假7',
             merged_df.columns[11]: '休假8', merged_df.columns[12]: '休假9', merged_df.columns[13]: '休假10'})
# merged_df.to_excel('/home/sf107/桌面/22222222222222.xlsx', index=False)
# print(merged_df, flush=True)
# print(merged_df.columns)

for index, row in df3.iterrows():
    # 查找姓名等于“申请人员姓名”的行
    rows_name = merged_df[merged_df['姓名'] == row['申请人员姓名']]

    if not rows_name.empty:
        # print(rows_name)
        # 获取第一行的索引
        first_row_index = rows_name.index[0]
        for i in range(10):
            # 检查'休假1'到'休假10'的列中是否有NaN值
            if pd.isna(rows_name.loc[first_row_index, f'休假{i + 1}']):
                # 在merged_df中直接修改值
                merged_df.loc[first_row_index, f'休假{i + 1}'] = row['起止日期'].replace("2024-", "").replace("-", ".").replace("至", "-")
                break
        if pd.isna(merged_df.loc[first_row_index, '已休假总天数']):
            merged_df.loc[first_row_index, '已休假总天数'] = 0
        merged_df.loc[first_row_index, '已休假总天数'] += row['本次使用年假（天）']
    else:
        print(f"没有找到姓名等于{row['申请人员姓名']}的行")

# print(merged_df)
# 填充数据
CCC = 5
for iii in range(len(merged_df)):
    for jjj in range(len(merged_df.columns)):
        cell_value = merged_df.iat[iii, jjj]
        if cell_value != '':
            ws_origin8.cell(iii + CCC, jjj + 1, value=cell_value)


wb_origin8.save(f"{directory}处理结果/年假处理表.xlsx")  # ！！！
print(f"生成---年假处理表.xlsx")
# subprocess.run(['python', '处理年假表.py'])
print('完成！')
