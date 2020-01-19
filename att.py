from argparse import ArgumentParser
from datetime import timedelta
from math import ceil
from os import path, listdir
from typing import List, Optional

import numpy as np
import pandas as pd
import openpyxl as xl
from openpyxl.styles import Border, Alignment, Side
from openpyxl.styles.colors import BLACK

import conf

"""
使用 Pandas 对多个涉及考勤数据的Excel表格进行处理
1. 原始打卡记录表
   数据格式：部门名称        人员编号   姓名       日期          最早打卡时间              最晚打卡时间
   数据示例：国家工程实验室   K01962    王丽梅	   2019-12-04	2019-12-04 08:20:53	    2019-12-04 18:04:51
2. 原始考勤异常记录表
   数据格式：序号  工号     姓名   部门                         职位     异常类型  开始日期    异常时数  异常情况说明/事由 流程状态
   数据示例：65	 K01962	 王丽梅	新智认知/国家工程实验室/办公室	行政专员	 漏打卡	  2019/12/6	 0	     漏打卡	         进行中
3. 原始请假记录表
   数据格式：序号  员工编号    员工姓名    假别    开始日期      结束日期        缺勤时长    开始时间(上午/下午)
   数据示例：1     610054982  王丽梅     9700   2019-12-12    2019-12-12    1.00       上午
4. 生成明细记录表
   数据格式：部门名称        人员编号   姓名   日期          上班       下班        上班  下班  星期  出勤天数  有异常流程   异常/没有记录
   数据示例：国家工程实验室   K01962	王丽梅  2019-12-04	08:20:53   18:04:51   8	    18	 3	  2		   12/3事假	
5. 上报考勤情况表
   数据格式：序号  姓名      迟到/早退       考勤异常（未有打卡记录）        备注
   数据示例：5	 王丽梅	   12/4-9:01      12/12下班卡，12/16全天

通过对前3张表格的处理，生成第4、5张表格
"""

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


def make_workdays(year, month: int, max_date: np.datetime64 = None) -> List["Timestamp"]:
    """生成指定月份的工作日期列表
    :param year: 指定年份，格式为2020
    :param month: 指定月份，格式为1-12
    :param max_date: 可以指定处理到那一天
    :return list: 月份列表['2020-01-03', ...]
    """
    result = []
    str_day = f'{year}-{month}'
    days_of_month = pd.Period(str_day).daysinmonth
    workdays = pd.date_range(str_day, periods=days_of_month)
    for _day in workdays:
        str_day = _day.strftime("%Y-%m-%d")
        if (str_day not in conf.WORKDAYS) and ((_day.dayofweek >= 5) or (str_day in conf.HOLIDAYS)):
            continue
        if (max_date is not None) and (_day > max_date):
            continue
        result.append(_day)
    return result


def get_next_workday(currdate: pd.Timestamp, days: int) -> pd.Timestamp:
    """下一个工作日
    :param currdate: 当前工作日
    :param days: 第几个工作日
    :return: 工作日
    """
    def next_workday(date: pd.Timestamp) -> pd.Timestamp:
        curr = date
        while True:
            curr = curr + timedelta(days=1)
            if curr.dayofweek < 5:
                break
        return curr

    date = currdate
    for i in range(days):
        date = next_workday(date)
    return date


def findfile_byname(file_path: str, find_str: str) -> Optional[str]:
    """根据文件名查找字符串，从指定目录中模糊查询出文件名，返回第一个符合条件的文件名
    :param file_path: 查找的目录
    :param find_str: 模糊查询的文件名
    :return: None 或者 文件名
    """
    try:
        files = listdir(file_path)
        for filename in files:
            if find_str in filename:
                return path.join(path.abspath(file_path), filename)
        return None
    except FileNotFoundError:
        return None


def read_att_info(filepath: str) -> Optional[pd.DataFrame]:
    """从文件中读取打卡信息
    :param filepath: 原始打卡数据Excel文件存放的目录（文件名称：打卡记录_20191227103406.xls）
    :return: None 或者 (打卡信息DataFrame (name, date, onduty, offduty)， 人员DataFrame)
    """
    # 读取原始打卡记录
    filename = findfile_byname(filepath, '打卡记录_')
    if filename is None:
        return None
    df = pd.read_excel(filename)
    # 将第一行去掉，只保留最后四列
    # 姓名    日期      最早打卡时间      最晚打卡时间
    df = df.iloc[1:, 2:]
    df.columns = ['name', 'date', 'onduty', 'offduty']
    # 对数据进行类型转换
    df[['date', 'onduty', 'offduty']] = df[['date', 'onduty', 'offduty']].astype(np.datetime64)
    return df


def read_abnormal_info(filepath: str) -> Optional[pd.DataFrame]:
    """从文件中读取考勤异常数据
    :param filepath: 原始打卡数据Excel文件存放的目录（文件名称：考勤异常数据_20191227.xls）
    :return: None 或者 考勤异常数据DataFrame（name, date type nums）
    """
    # 读取考勤异常记录
    filename = findfile_byname(filepath, '考勤异常数据_')
    if filename is None:
        return None
    df = pd.read_excel(filename)
    df = pd.DataFrame({'name': df['姓名'],
                       'date': df['开始日期'],
                       'type': df['异常类型'],
                       'time': df['异常时数']})
    # 对数据进行类型转换
    df['time'] = df['time'].astype(np.float64)
    # 筛选出 time > 8 的记录
    # 例如：XXX    2019-12-12      培训      16
    # 添加：XXX    2019-12-13      培训
    df_1 = df[df['time'] > 8]
    for i in range(len(df_1)):
        row = df_1.iloc[i]
        time = row['time']
        for j in range(ceil(time / 8) - 1):
            value = row.copy()
            value['date'] = get_next_workday(value['date'], j + 1)
            value['time'] = np.NaN
            df = df.append(value, ignore_index=True)
    # 去掉重复的记录
    df = df.drop_duplicates()
    return df


def read_offwork_info(filepath: str) -> Optional[pd.DataFrame]:
    """从文件中读取请假异常数据
    :param filepath: 原始请假数据Excel文件存放的目录（文件名：考勤汇总表-请假(1).xls）
    :return: None 或者 请假数据DateFrame (name, type, startdate, enddate, times)
    """
    # 读取请假记录
    filename = findfile_byname(filepath, '请假')
    if filename is None:
        return None
    df = pd.read_excel(filename, sheet_name=1)
    df = df.iloc[1:, 2:-1]
    df.columns = ['name', 'type', 'date', 'enddate', 'time']
    df.drop(['enddate'], axis=1, inplace=True)
    # 对数据进行类型转换
    df['date'] = df['date'].astype(np.datetime64)
    df['time'] = df['time'].astype(np.float64)
    # 单位转换为小时 * 8
    df['time'] = df['time'] * 8
    # 将请假类型转换为中文描述
    df['type'] = df.apply(lambda row: conf.HOLIDAY_TYPE[str(row.type)], axis=1)
    return df


def general_blank_dataframe(filepath: str, enddate: str, df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """生成空白DF
    :param filepath: 文件存放的目录 201912
    :param enddate: 考勤统计截止日期
    :param df: 原始打卡记录
    :return: 空白列头文件 (name, date)
    """
    try:
        max_date = pd.Timestamp(enddate)
    except:
        max_date = df['date'].max()
    names = df['name'].drop_duplicates()
    days = make_workdays(max_date.year, max_date.month, max_date)
    df_blank = pd.DataFrame(columns=['name', 'date'])
    for name in names:
        for day in days:
            df_blank = df_blank.append({'name': name, 'date': day}, ignore_index=True)
    df_blank['date'] = df_blank['date'].astype(np.datetime64)
    return df_blank


def general_final_info(filepath: str, enddate: str) -> Optional[pd.DataFrame]:
    """生成最终的上报文件
    :param filepath: 输出文件的目录
    :param enddate: 考勤统计截止日期
    :return: 最终
    """
    def _general_chidao(row):
        """生成迟到早退列的数据
        :param row: map的每一行数据
        :return:None
        """
        if row.type is not np.NaN:
            return np.NaN
        if row.onduty == row.offduty:
            return np.NaN
        if (row.onduty.hour > 9) or (row.onduty.hour == 9 and row.onduty.minute > 0):
            return row.onduty.strftime('%m/%d-%H:%M')
        if row.offduty.hour < 18:
            return row.offduty.strftime('%m/%d-%H:%M')
        return np.NaN

    def _general_abnormal(row):
        """生成考勤异常的数据
        :param row: map的每一行数据
        :return:None
        """
        if row.type is not np.NaN:
            return np.NaN
        # 没有刷过卡
        if (row.onduty is pd.NaT) or (row.offduty is pd.NaT):
            return row.date.strftime('%m/%d') + '全天卡'
        # 上班/下班刷卡
        if row.onduty == row.offduty:
            if row.onduty.hour < 12:
                return row.onduty.strftime('%m/%d') + '下班卡'
            else:
                return row.onduty.strftime('%m/%d') + '上班卡'
        return np.NaN

    df_att = read_att_info(filepath)
    df_abnormal = read_abnormal_info(filepath)
    df_offwork = read_offwork_info(filepath)
    # 生成空表，由人员和本月的工作时间组成 (name, date)
    df_final = general_blank_dataframe(filepath, enddate, df_att)
    # 合并考勤异常和请假
    df_abn = pd.concat([df_abnormal, df_offwork], sort=False)
    # 连接考勤表 (name, date, onduty, offduty)
    df_final = pd.merge(df_final, df_att, how='left', on=['name', 'date'])
    # 连接考勤异常表(name, date, onduty, offduty, type, time)
    df_final = pd.merge(df_final, df_abn, how='left', on=['name', 'date'])
    # 生成迟到早退信息
    df_final['chidao'] = df_final.apply(_general_chidao, axis=1)
    # 生成考勤异常
    df_final['abn'] = df_final.apply(_general_abnormal, axis=1)
    return df_final


def set_excel_style(filename: str) -> None:
    """设置汇总表格文件中单元格的样式
    :param filename: excel文件名称
    :return:
    """
    cells_format = [
        {'B': 'yyyy-mm-dd', 'C': 'hh:mm:ss', 'D': 'hh:mm:ss'},
        {}
    ]
    cells_width = [
        [('B', 13), ('C', 13), ('D', 13), ('G', 12), ('H', 25)],
        [('B', 20), ('C', 25)]
    ]
    try:
        # #### 定义样式
        # 设置边框
        thin = Side(border_style="thin", color=BLACK)
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        alignment = Alignment(wrapText=True, horizontal='center')
        wb = xl.load_workbook(filename)
        # 遍历所有worksheet
        for ws in wb:
            # 获得当前worksheet的索引
            ws_idx = wb.index(ws)
            for row in ws.rows:
                for cell in row:
                    # 给所有单元格设置边框
                    cell.border = border
                    cell.alignment = alignment
                    key = cell.column_letter
                    if key in cells_format[ws_idx]:
                        cell.number_format = cells_format[ws_idx][key]

            # 对列宽进行设置
            for idx, wid in cells_width[ws_idx]:
                ws.column_dimensions[idx].width = wid
            # 对单元格数值格式进行设置，主要是日期时间类型

        wb.save(filename)
        wb.close()
    except:
        print(f'设置{filename}格式时，出现异常！')


def write_to_excel(df_data: pd.DataFrame, filename: str) -> None:
    """将数据写入Excel表中
    :param data: 要写入的数据集
    :param filename: 文件名
    :return: None
    """
    def _general_stat(df: pd.DataFrame) -> pd.Series:
        stat1, stat2 = [], []
        for i in range(len(df)):
            if df.iloc[i]['chidao'] not in (np.NaN, None):
                stat1.append(df.iloc[i]['chidao'])
            if df.iloc[i]['abn'] not in (np.NaN, None):
                stat2.append(df.iloc[i]['abn'])
        return pd.Series({'chidao': '\n'.join(stat1), 'abn': ';'.join(stat2)})

    df_group = pd.DataFrame({'name': df_data['name'], 'chidao': df_data['chidao'], 'abn': df_data['abn']})
    df_group = df_group.groupby('name').apply(_general_stat)
    # 列头更换为中文
    df_data_header = ['姓名', '日期', '上班打卡', '下班打卡', '异常类型', '异常时数', '迟到/早退', '考勤异常（未有打卡记录）']
    df_group_header = ['迟到/早退', '考勤异常（未有打卡记录）']
    # 写入文件中
    with pd.ExcelWriter(filename) as writer:
        df_data.to_excel(writer, sheet_name='详情', header=df_data_header, index=False)
        df_group.to_excel(writer, sheet_name='汇总', header=df_group_header, index_label='姓名')
        writer.save()
    set_excel_style(filename)


def main():
    parser = ArgumentParser(description='生成每月考勤汇总记录.')
    parser.add_argument('path', help="请输入存放考勤记录的目录", default='.')
    parser.add_argument('enddate', help="请输入考勤统计的截止日期，如：20200122")
    args = parser.parse_args()
    df = general_final_info(args.path, args.enddate)
    write_to_excel(df, f'实验室打卡记录汇总-{args.enddate}.xlsx')


if __name__ == '__main__':
    main()
