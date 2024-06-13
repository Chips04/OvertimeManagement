import pandas as pd
from 已完成.functions import *

df = pd.read_excel('/home/sf107/桌面/替换重名完成输出文件.xlsx')  # ！！！！！！！
# 替换重名
rn(df, '值班人员')  # ！！！！！！！
# 输出表格
df.to_excel('/home/sf107/桌面/替换重名完成输出文件.xlsx', index=False)
