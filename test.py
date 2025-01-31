import pandas as pd
import pyautogui
import datetime

# 获取日期，格式为mmdd
date = datetime.datetime.now().strftime('%m%d')

# 存储文档路径
file_path = '/Users/bensen/Desktop/2025年排名.xlsx'

# 读取Excel文件
df = pd.read_excel(file_path)

# 获取玩家名称
player_names = df['名称'].tolist()

# 设置一个字典，包括所有玩家，默认得分均为0
score = {key:0 for key in player_names}

# 弹窗输入玩家的当日战绩
warning_title = "请输入 " + date + " 的战绩情况"

for i in player_names:
    score[i] = int(pyautogui.prompt(text='请输入' + i + '本日的战绩', title=warning_title, default=''))

personal_score = []

for i in player_names:
    personal_score.append(score[i])

# 在最右边增加一列，名为当天日期，值为每个人的得分
df[date] = personal_score

df.to_excel(file_path, index=False)

# 计算得分合计（假设得分在D列到Z列）
df['合计'] = df.iloc[:, 3:26].sum(axis=1)  # 3:26表示D列到Z列的索引范围

# 根据得分合计从高到低排序，并重置索引
df_sorted = df.sort_values(by='合计', ascending=False).reset_index(drop=True)

# 更新排名（A列）
df_sorted['排名'] = df_sorted.index + 1

# 将更新后的数据保存回Excel文件
df_sorted.to_excel(file_path, index=False)
