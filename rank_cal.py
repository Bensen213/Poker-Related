import pandas as pd
import pyautogui
import datetime
import openpyxl


def get_sheet_name():
    """
    根据当前日期确定对应的季度工作表名称
    :return: 工作表名称
    """
    date = datetime.datetime.now().strftime('%m%d')
    month = int(date[:2])
    if month <= 3:
        return "Q1"
    elif month <= 6:
        return "Q2"
    elif month <= 9:
        return "Q3"
    else:
        return "Q4"


def get_player_scores(player_names, date):
    """
    通过弹窗获取每个玩家当天的战绩
    :param player_names: 玩家名称列表
    :param date: 日期
    :return: 玩家得分字典
    """
    warning_title = "请输入 " + date + " 的战绩情况"
    score = {key: 0 for key in player_names}
    for player in player_names:
        try:
            score[player] = int(pyautogui.prompt(text='请输入' + player + '本日的战绩', title=warning_title, default=''))
        except ValueError:
            print(f"输入的 {player} 的战绩不是有效的整数，将使用默认值 0。")
    return score


def process_excel(file_path):
    """
    处理 Excel 文件，更新玩家战绩和排名
    :param file_path: Excel 文件路径
    """
    try:
        # 获取日期，格式为 mmdd
        date = datetime.datetime.now().strftime('%m%d')
        # 获取对应的工作表名称
        sheet_name = get_sheet_name()
        # 读取 Excel 文件
        wb = openpyxl.load_workbook(file_path)
        try:
            ws = wb[sheet_name]
            data = ws.values
            columns = next(data)
            df = pd.DataFrame(data, columns=columns)
        except KeyError:
            print(f"Excel 文件中不存在名为 {sheet_name} 的工作表，请检查文件。")
            return
        # 获取玩家名称
        player_names = df['名称'].tolist()
        # 获取玩家得分
        scores = get_player_scores(player_names, date)
        personal_scores = [scores[player] for player in player_names]
        # 在最右边增加一列，名为当天日期，值为每个人的得分
        df[date] = personal_scores
        # 计算得分合计（假设得分在 D 列到 Z 列）
        df['合计'] = df.iloc[:, 3:26].sum(axis=1)
        # 根据得分合计从高到低排序，并重置索引
        df_sorted = df.sort_values(by='合计', ascending=False).reset_index(drop=True)
        # 更新排名（A 列）
        df_sorted['排名'] = df_sorted.index + 1
        # 清除原工作表内容
        ws.delete_rows(1, ws.max_row)
        # 写入表头
        ws.append(df_sorted.columns.tolist())
        # 写入数据
        for row in df_sorted.values.tolist():
            ws.append(row)
        # 保存修改后的 Excel 文件
        wb.save(file_path)
        print("数据更新成功！")
    except FileNotFoundError:
        print(f"文件 {file_path} 未找到，请检查文件路径。")
    except Exception as e:
        print(f"发生错误: {e}")


if __name__ == "__main__":
    file_path = '/Users/bensen/Desktop/2025年排名.xlsx'
    process_excel(file_path)