# calculations.py
# IAA休闲游戏成本收益核心计算模块

from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import os
import pickle
import io
from openpyxl import Workbook
from scipy.optimize import curve_fit
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side

"""   根据前端传入的参数，计算所有IAA游戏相关的成本收益指标 """
def calculate_metrics(params: dict) -> dict:
    """

    Args:
        params (dict): 一个包含所有前端输入的字典。结构如下：
            {
              "project_name": str,                    # 项目名称
              "investment_periods": [                 # 资源投入时间段列表
                { 
                  "start": str,                       # 起始时间 (YYYY-MM-DD)
                  "end": str,                         # 结束时间 (YYYY-MM-DD)
                  "cost_type": str,                   # 日消耗类型: "c" 或 "linear"
                  "cost_value": float,                # 定值日消耗 (万元)，当cost_type为"fixed"时使用
                  "cost_start": float,                # 线性增长初始值 (万元)，当cost_type为"linear"时使用
                  "cost_end": float,                  # 线性增长最终值 (万元)，当cost_type为"linear"时使用
                  "dnu": float,                       # DNU (万人)
                  "team_size": int,                   # 团队规模 (人)
                  "labor_cost": float,                # 用工成本 (万/人/天)
                  "other_cost": float                 # 其他运营成本 (万元/天)
                }
              ],
              "roi_data": {                           # ROI数据
                "type": str,                          # 输入类型: "manual" 或 "excel"
                "points": [                           # 数据点列表
                  {"day": int, "value": float},       # day: 天数, value: ROI值(%)
                  ...
                ]
              },
              "retention_data": {                     # 留存率数据
                "type": str,                          # 输入类型: "manual" 或 "excel"
                "points": [                           # 数据点列表
                  {"day": int, "value": float},       # day: 天数, value: 留存率(%)
                  ...
                ]
              },
              "repayment_months": int,                # 现金回款月数 (1或2)
              "target_dau": int                       # 目标DAU (万人)，用于计算达成目标DAU指标
            }

    Returns:
        dict: 一个包含所有计算结果的字典，用于前端展示。结构如下：
            {
                "key_metrics": {                      # 关键指标
                    "max_cash_demand_1m": float,      # 1个月回款最大现金需求 (万元)
                    "max_cash_demand_2m": float,      # 2个月回款最大现金需求 (万元)
                    "dynamic_profit_breakeven_day": int,          # 动态会计利润打正天数
                    "cumulative_profit_breakeven_day": int,       # 累计会计利润打正天数
                    "cumulative_cash_flow_1m_breakeven_day": int, # 1个月回款累计现金流打正天数
                    "cumulative_cash_flow_2m_breakeven_day": int, # 2个月回款累计现金流打正天数
                    "day_to_10m_dau": int,            # 达到1000万DAU的天数
                    "day_to_2m_dau": int,             # 达到200万DAU的天数
                    "day_to_target_dau": int          # 达到目标DAU的天数 (-1表示未达成)
                },
                "charts": {                           # 图表数据
                    "dau_quarterly": {                # DAU季度折线图数据
                        "labels": ["26Q1", "26Q2", ...],   # 季度标签
                        "data": [150000, 710000, ...]      # 季度末DAU数据 (人)
                    },
                    "finance_quarterly": {            # 收入成本季度图表数据
                        "labels": ["26Q1", "26Q2", ...],   # 季度标签
                        "income": [111, 724, ...],         # 当期收入 (万元)
                        "cost": [-361, -1262, ...],        # 当期成本 (万元，负数)
                        "cumulative_profit": [-250, -788, ...]  # 累计利润 (万元)
                    }
                },
                "quarterly_table_data": [             # 季度汇总表格数据
                    {
                        "quarter": "26Q1",            # 季度标识
                        "end_date": "2026/4/1",       # 季度结束日期
                        "cumulative_days": 59,        # 累计天数
                        "cumulative_revenue": 111,    # 累计收入 (万元)
                        "cumulative_cost": -361,      # 累计成本 (万元)
                        "ua_cost": -342,              # UA成本 (万元)
                        "personnel_cost": -15,        # 人员成本 (万元)
                        "other_cost": -3,             # 其他成本 (万元)
                        "cumulative_profit": -250,    # 累计利润 (万元)
                        "cumulative_cash_flow_1m": -352,   # 1个月回款累计现金流 (万元)
                        "cumulative_cash_flow_2m": -361,   # 2个月回款累计现金流 (万元)
                        "current_revenue": 111,       # 当期收入 (万元)
                        "current_cost": -361,         # 当期成本 (万元)
                        "current_ua_cost": -342,      # 当期UA成本 (万元)
                        "current_personnel_cost": -15,     # 当期人员成本 (万元)
                        "current_other_cost": -3,     # 当期其他成本 (万元)
                        "current_profit": -250,       # 当期利润 (万元)
                        "current_cash_demand_1m": 352,     # 当期1个月回款现金需求 (万元)
                        "current_cash_demand_2m": 361,     # 当期2个月回款现金需求 (万元)
                        "dau": 15                     # 季度末DAU (万人)
                    },
                    # ... 其他季度的数据
                ],
                "daily_data_csv_string": "Date,DNU,DAU,Revenue,Cost,...\n2026-01-01,10000,10000,..."  # 用于CSV下载的字符串
            }
    """
    
    """ 1. 读取基础参数 """
    project_name = params.get('project_name')  # 项目名称，默认IAA
    repayment_months = params.get('repayment_months', 1)  # 现金回款月数，默认1个月
    repayment_flag = True if repayment_months > 2 else False
    target_dau = params.get('target_dau', 500)  # 目标DAU（万人），默认500万

    
    """ 2. 解析投资时间段 investment_periods """
    investment_periods_raw = params.get('investment_periods', [])
    investment_periods = []
    total_true_investment_days = 0  # 所有时间段的总天数
    for period in investment_periods_raw:
        start_date = datetime.strptime(period['start'], '%Y-%m-%d')
        end_date = datetime.strptime(period['end'], '%Y-%m-%d')
        period_days = (end_date - start_date).days + 1  # 包含起始和结束当天
        total_true_investment_days += period_days
        parsed_period = {
            'start': start_date,                                      # 起始时间转为日期
            'end': end_date,                                          # 结束时间转为日期
            'days': period_days,                                      # 该时间段天数
            'cost_type': period.get('cost_type', 'fixed'),            # 日消耗类型
            'cost_value': period.get('cost_value', 0),                # 定值日消耗（万元）
            'cost_start': period.get('cost_start', 0),                # 线性变化初始值（万元）
            'cost_end': period.get('cost_end', 0),                    # 线性变化最终值（万元）
            'dnu': period.get('dnu', 0),                              # DNU（万人）
            'team_size': period.get('team_size', 0),                  # 团队规模（人）
            'labor_cost': period.get('labor_cost', 0),                # 用工成本（万/人/天）
            'other_cost': period.get('other_cost', 0)                 # 其他运营成本（万元/天）
        }
        investment_periods.append(parsed_period)  # 将解析后的时间段添加到列表中

    # 计算需要的总天数（包含回款延迟天数，用于现金流计算）
    total_investment_days = total_true_investment_days + min(repayment_months, 2) * 30

    """ 3. 解析 ROI 数据 """
    roi_data_raw = params.get('roi_data', {'type': 'manual', 'points': []})
    roi_points = roi_data_raw.get('points', [])
    roi_input_type = roi_data_raw.get('type')
    
    # 采用Excel导入
    if roi_input_type == 'excel':
        roi_raw = [p['value'] / 100 for p in roi_points]
        # 超过则截断
        if len(roi_raw) >= total_investment_days:
            roi_vector = roi_raw[:total_investment_days]
        # 不足则使用已有数据进行外推
        else:
            known_days = list(range(len(roi_raw)))
            extrapolated = fit_roi_curve_advanced(known_days, roi_raw, total_investment_days, "roi")
            roi_vector = roi_raw + extrapolated[len(roi_raw):]
    # 采用手动输入，则直接用曲线拟合的形式填充所有数据
    else:
        known_days = [p['day'] - 1 for p in roi_points]  # 转换为0-based索引
        known_values = [p['value'] / 100 for p in roi_points]
            
        # 按天数排序
        known_days, known_values = zip(*sorted(zip(known_days, known_values)))
            
        # 使用曲线拟合生成完整数据
        roi_vector = fit_roi_curve_advanced(known_days, known_values, total_investment_days, "roi")
    
    """ 4. 解析留存率数据 """
    retention_data_raw = params.get('retention_data', {'type': 'manual', 'points': []})
    retention_points = retention_data_raw.get('points', [])
    retention_input_type = retention_data_raw.get('type')

    # 采用Excel导入
    if retention_input_type == 'excel':
        retention_raw = [p['value'] / 100 for p in retention_points]
        # 超过则截断
        if len(retention_raw) >= total_investment_days:
            retention_vector = retention_raw[:total_investment_days]
        # 不足则使用已有数据进行外推
        else:
            known_days = list(range(len(retention_raw)))
            extrapolated = fit_roi_curve_advanced(known_days, retention_raw, total_investment_days, "retention")
            retention_vector = retention_raw + extrapolated[len(retention_raw):]
    # 采用手动输入，则直接用曲线拟合的形式填充所有数据
    else:
        known_days = [p['day'] for p in retention_points]  # 转换为0-based索引
        known_values = [p['value'] / 100 for p in retention_points]

        # 按天数排序
        known_days, known_values = zip(*sorted(zip(known_days, known_values)))

        # 使用曲线拟合生成完整数据
        retention_vector = fit_roi_curve_advanced(known_days, known_values, total_investment_days, "retention")

    """ 5. 将所有原始数据转化为dataframe """
    # 创建一个行数为总天数、列数为43的DataFrame
    df = pd.DataFrame(np.zeros((total_investment_days, 43)))
    
    # 第0列：天数索引，从0开始直到总天数-1
    df.iloc[:, 0] = list(range(total_investment_days))
    
    # 遍历所有投资时间段，填充每日数据
    current_day = 0  # 当前天数索引
    for period in investment_periods:
        period_days = period['days']  # 该时间段的天数
        
        # 第1列：日消耗
        if period['cost_type'] == 'fixed':
            # 定值日消耗
            df.iloc[current_day:current_day + period_days, 1] = period['cost_value']
        else:
            # 线性变化日消耗
            cost_start = period['cost_start']
            cost_end = period['cost_end']
            # 生成从cost_start到cost_end的线性序列
            linear_costs = np.linspace(cost_start, cost_end, period_days)
            df.iloc[current_day:current_day + period_days, 1] = linear_costs
        
        # 第4列：团队规模 * 用工成本
        labor_total = period['team_size'] * period['labor_cost']
        df.iloc[current_day:current_day + period_days, 4] = labor_total
        
        # 第5列：其他运营成本
        df.iloc[current_day:current_day + period_days, 5] = period['other_cost']

        # 第20列：DNU
        df.iloc[current_day:current_day + period_days, 20] = period['dnu']
        
        current_day += period_days
    
    # 第9列：ROI的值
    df.iloc[:, 9] = roi_vector

    # 第41列：每日ROI增长
    df.iloc[:, 41] = df.iloc[:, 9].diff()
    df.iloc[0, 41] = df.iloc[0, 9]  # 第一天的ROI增长为第一天的ROI值
    
    # 第42列：留存率的值
    df.iloc[:, 42] = retention_vector

    """ 6. 根据原始数据计算每日数据 """
    # 第2列：每日收入，第i天的收入 = Σ(j=0 to i) [第j天的UA支出 × 第j天用户在第(i-j)天的ROI_DIFF]
    ua_cost_array = df.iloc[:, 1].values
    roi_array = df.iloc[:, 41].values
    n = len(ua_cost_array)
    daily_revenue_array = np.zeros(n)
    
    for i in range(n):
        revenue_sum = 0
        for j in range(i + 1):
            revenue_sum += ua_cost_array[j] * roi_array[i - j]
        daily_revenue_array[i] = revenue_sum
    
    df.iloc[:, 2] = daily_revenue_array

    # 第3列：账面毛利
    df.iloc[:, 3] = df.iloc[:, 2] - df.iloc[:, 1]

    # 第6列：其他成本合计
    df.iloc[:, 6] = df.iloc[:, 4] + df.iloc[:, 5]

    # 第7列：总成本
    df.iloc[:, 7] = df.iloc[:, 1] + df.iloc[:, 6]

    # 第8列：账面净利
    df.iloc[:, 8] = df.iloc[:, 3] - df.iloc[:, 6]

    # 第10列：累积消耗
    df.iloc[:, 10] = df.iloc[:, 1].cumsum()

    # 第11列：累积收入
    df.iloc[:, 11] = df.iloc[:, 2].cumsum()

    # 第12列：总毛利
    df.iloc[:, 12] = df.iloc[:, 3].cumsum()

    # 第13列：累积用工成本
    df.iloc[:, 13] = df.iloc[:, 4].cumsum()

    # 第14列：累积其他运营成本
    df.iloc[:, 14] = df.iloc[:, 5].cumsum()

    # 第15列：累积其他成本
    df.iloc[:, 15] = df.iloc[:, 6].cumsum()

    # 第16列：累积总成本
    df.iloc[:, 16] = df.iloc[:, 15] + df.iloc[:, 10]

    # 第17列：总净利
    df.iloc[:, 17] = df.iloc[:, 12] - df.iloc[:, 15]

    # 第18列：总净利2
    df.iloc[:, 18] = df.iloc[:, 11] - df.iloc[:, 16]

    # 第19列：CPI
    df.iloc[:, 19] = df.iloc[:, 1] / df.iloc[:, 20]

    # 第21列：DAU
    # DAU_t = Σ(DNU_i × r_{t-i})，i从0到t
    dnu_array = df.iloc[:, 20].values
    retention_array = df.iloc[:, 42].values
    n = len(dnu_array)
    dau_array = np.zeros(n)
    
    for t in range(n):
        dau_sum = 0
        for i in range(t + 1):
            retention_index = t - i
            dau_sum += dnu_array[i] * retention_array[retention_index]
        dau_array[t] = dau_sum
    
    df.iloc[:, 21] = dau_array

    # 第22列：ARPU
    df.iloc[:, 22] = df.iloc[:, 2] / df.iloc[:, 21]

    # 第23列：滞后一个月的累积现金收入
    df.iloc[:, 23] = df.iloc[:, 11].shift(30).fillna(0)

    # 第24列：滞后一个月的累积现金缺口
    df.iloc[:, 24] = df.iloc[:, 23] - df.iloc[:, 16]

    # 第25列：滞后两个月的累积现金收入
    df.iloc[:, 25] = df.iloc[:, 11].shift(60).fillna(0)

    # 第26列：滞后两个月的累积现金缺口
    df.iloc[:, 26] = df.iloc[:, 25] - df.iloc[:, 16]

    # 第27列：累积回本判断，总净利大于0标注0，小于标注1
    df.iloc[:, 27] = df.iloc[:, 17].apply(lambda x: 0 if x > 0 else 1)

    # 第28列：累积回本倒序求和
    df.iloc[:, 28] = df.iloc[:, 27].iloc[::-1].cumsum().iloc[::-1]

    # 第29列：当期利润回正，总利润大于0标注0，小于标注1
    df.iloc[:, 29] = df.iloc[:, 8].apply(lambda x: 0 if x > 0 else 1)

    # 第30列：当期利润回正倒序求和
    df.iloc[:, 30] = df.iloc[:, 29].iloc[::-1].cumsum().iloc[::-1]

    # 第31列：累计现金流打正（滞后1个月），现金流大于0标注0，小于标注1
    df.iloc[:, 31] = df.iloc[:, 24].apply(lambda x: 0 if x > 0 else 1)

    # 第32列：累计现金流打正（滞后1个月）倒序求和
    df.iloc[:, 32] = df.iloc[:, 31][::-1].cumsum()[::-1].values

    # 第33列：累计现金流打正（滞后2个月），现金流大于0标注0，小于标注1
    df.iloc[:, 33] = df.iloc[:, 26].apply(lambda x: 0 if x > 0 else 1)

    # 第34列：累计现金流打正（滞后2个月）倒序求和
    df.iloc[:, 34] = df.iloc[:, 33][::-1].cumsum()[::-1].values

    # 第35列：达成1000万DAU，达成标注0，未达成标注1
    df.iloc[:, 35] = df.iloc[:, 21].apply(lambda x: 0 if x > 1000 else 1)

    # 第36列：达成1000万DAU倒序求和
    df.iloc[:, 36] = df.iloc[:, 35][::-1].cumsum()[::-1].values

    # 第37列：达成200万DAU，达成标注0，未达成标注1
    df.iloc[:, 37] = df.iloc[:, 21].apply(lambda x: 0 if x > 200 else 1)

    # 第38列：达成200万DAU倒序求和
    df.iloc[:, 38] = df.iloc[:, 37][::-1].cumsum()[::-1].values

    # 第39列：达成xx万DAU，达成标注0，未达成标注1
    df.iloc[:, 39] = df.iloc[:, 21].apply(lambda x: 0 if x > target_dau else 1)

    # 第40列：达成xx万DAU倒序求和
    df.iloc[:, 40] = df.iloc[:, 39][::-1].cumsum()[::-1].values

    # 考虑现金回款月数大于2的情况
    df_extra = pd.DataFrame(np.zeros((total_investment_days, 4)))
    if repayment_flag:
        # 第0列：滞后n个月的累积现金收入
        df_extra.iloc[:, 0] = df.iloc[:, 11].shift(repayment_months * 30).fillna(0)

        # 第1列：滞后n个月的累积现金缺口
        df_extra.iloc[:, 1] = df_extra.iloc[:, 0] - df.iloc[:, 16]

        # 第2列：累计现金流打正（滞后n个月），现金流大于0标注0，小于标注1
        df_extra.iloc[:, 2] = df_extra.iloc[:, 1].apply(lambda x: 0 if x > 0 else 1)

        # 第3列：累计现金流打正（滞后n个月）倒序求和
        df_extra.iloc[:, 3] = df_extra.iloc[:, 2][::-1].cumsum()[::-1].values
    
    """ 7. 计算关键指标 key_metrics """
    # 最大现金需求
    max_cash_demand_1m = abs(df.iloc[:, 24].min()) if df.iloc[:, 24].min() < 0 else 0
    max_cash_demand_2m = abs(df.iloc[:, 26].min()) if df.iloc[:, 26].min() < 0 else 0
    if repayment_flag:
        max_cash_demand_nm = abs(df_extra.iloc[:, 1].min()) if df_extra.iloc[:, 1].min() < 0 else 0
    else:
        max_cash_demand_nm = 0

    # 动态会计利润打正天数：找到第30列第一个为0的行索引
    dynamic_profit_breakeven_day = -1
    for i in range(len(df)):
        if df.iloc[i, 30] == 0:
            dynamic_profit_breakeven_day = int(df.iloc[i, 0])
            break
    
    # 累计会计利润打正天数：找到第28列第一个为0的行索引
    cumulative_profit_breakeven_day = -1
    for i in range(len(df)):
        if df.iloc[i, 28] == 0:
            cumulative_profit_breakeven_day = int(df.iloc[i, 0])
            break
    
    # 1个月回款累计现金流打正天数：找到第32列第一个为0的行索引
    cumulative_cash_flow_1m_breakeven_day = -1
    for i in range(len(df)):
        if df.iloc[i, 32] == 0:
            cumulative_cash_flow_1m_breakeven_day = int(df.iloc[i, 0])
            break
    
    # 2个月回款累计现金流打正天数：找到第34列第一个为0的行索引
    cumulative_cash_flow_2m_breakeven_day = -1
    for i in range(len(df)):
        if df.iloc[i, 34] == 0:
            cumulative_cash_flow_2m_breakeven_day = int(df.iloc[i, 0])
            break

    # n个月回款累计现金流打正天数：找到第df_extra 的3列第一个为0的行索引
    cumulative_cash_flow_nm_breakeven_day = -1
    for i in range(len(df_extra)):
        if df_extra.iloc[i, 3] == 0:
            cumulative_cash_flow_nm_breakeven_day = int(df_extra.iloc[i, 0])
            break
    
    # 达成1000万DAU天数：找到第36列第一个为0的行索引
    day_to_10m_dau = -1
    for i in range(len(df)):
        if df.iloc[i, 36] == 0:
            day_to_10m_dau = int(df.iloc[i, 0])
            break
    
    # 达成200万DAU天数：找到第38列第一个为0的行索引
    day_to_2m_dau = -1
    for i in range(len(df)):
        if df.iloc[i, 38] == 0:
            day_to_2m_dau = int(df.iloc[i, 0])
            break
    
    # 达成目标DAU天数：找到第40列第一个为0的行索引
    day_to_target_dau = -1
    for i in range(len(df)):
        if df.iloc[i, 40] == 0:
            day_to_target_dau = int(df.iloc[i, 0])
            break
    
    """ 8. 按季度聚合数据，生成图表和表格数据 """
    # 获取第一个投资时间段的起始日期作为基准日期
    base_date = investment_periods[0]['start']

    # 为df添加日期列
    df['date'] = pd.to_datetime([base_date + timedelta(days=int(d))
                                 for d in df.iloc[:, 0]])

    # 使用pandas内置功能计算季度
    df['quarter'] = df['date'].dt.to_period('Q').astype(str).str.replace('20', '')

    # 只取真实投资天数范围内的数据用于季度聚合和表格绘制
    df_display = df.iloc[:total_true_investment_days].copy()

    # 获取有序的季度列表
    quarters = df_display['quarter'].unique().tolist()

    # 图表数据
    dau_quarterly_labels = []
    dau_quarterly_data = []
    finance_quarterly_labels = []
    finance_quarterly_income = []
    finance_quarterly_cost = []
    finance_quarterly_cumulative_profit = []
    
    # 季度表格数据
    quarterly_table_data = []
    
    prev_cumulative_revenue = 0
    prev_cumulative_cost = 0
    prev_cumulative_ua_cost = 0
    prev_cumulative_personnel_cost = 0
    prev_cumulative_profit = 0
    prev_cumulative_other_cost = 0
    prev_cumulative_cash_flow_1m = 0
    prev_cumulative_cash_flow_2m = 0
    prev_cumulative_cash_flow_nm = 0
    
    for quarter in quarters:
        quarter_df = df_display[df_display['quarter'] == quarter]
        if len(quarter_df) == 0:
            continue
        
        quarter_end_idx = quarter_df.index[-1]
        quarter_end_date = quarter_df['date'].iloc[-1]

        cumulative_days = int(df_display.loc[quarter_end_idx, 0]) + 1  # 累积天数
        cumulative_revenue = round(df_display.loc[quarter_end_idx, 11], 2)  # 累积收入
        cumulative_total_cost = -round(df_display.loc[quarter_end_idx, 16], 2)  # 累积成本
        cumulative_ua_cost = -round(df_display.loc[quarter_end_idx, 10], 2)  # 累积UA成本
        cumulative_personnel_cost = -round(df_display.loc[quarter_end_idx, 13], 2)  # 累积人员成本
        cumulative_other_cost = -round(df_display.loc[quarter_end_idx, 14], 2)  # 累积其他成本
        cumulative_cash_flow_1m = round(df_display.loc[quarter_end_idx, 24], 2) if not pd.isna(df_display.loc[quarter_end_idx, 24]) else 0  # 累积净现金流（一个月回款）
        cumulative_cash_flow_2m = round(df_display.loc[quarter_end_idx, 26], 2) if not pd.isna(df_display.loc[quarter_end_idx, 26]) else 0  # 累积净现金流（两个月回款）
        
        # n个月回款累计现金流
        if repayment_flag:
            cumulative_cash_flow_nm = round(df_extra.loc[quarter_end_idx, 1], 2) if not pd.isna(df_extra.loc[quarter_end_idx, 1]) else 0
        else:
            cumulative_cash_flow_nm = 0

        current_revenue = round(cumulative_revenue - prev_cumulative_revenue, 2) # 当期收入 = 累计收入 - 上期累计收入
        current_cost = round(cumulative_total_cost - prev_cumulative_cost, 2)  # 当期成本 = 累积成本 - 上期累积成本
        current_ua_cost = round(cumulative_ua_cost - prev_cumulative_ua_cost, 2)  # 当期UA成本 = 累积UA成本 - 上期累积UA成本
        current_personnel_cost = round(cumulative_personnel_cost - prev_cumulative_personnel_cost, 2)  # 当期人员成本 = 累积人员成本 - 上期累积人员成本
        current_other_cost = round(cumulative_other_cost - prev_cumulative_other_cost, 2)  # 当期其他成本 = 累积其他成本 - 上期累积其他成本
        current_profit = round(current_revenue + current_cost, 2)  # 当期利润 = 当期收入 + 当期总成本

        cumulative_profit = round(prev_cumulative_profit + current_profit, 2)  # 累计利润 = 上期累积利润 + 当期利润
        
        # 当期资金需求 = -（累积净现金流 - 上期累积净现金流）小于0则置0
        current_cash_demand_1m = max(0, -round(cumulative_cash_flow_1m - prev_cumulative_cash_flow_1m, 2))
        current_cash_demand_2m = max(0, -round(cumulative_cash_flow_2m - prev_cumulative_cash_flow_2m, 2))
        
        # n个月回款当期资金需求
        if repayment_flag:
            current_cash_demand_nm = max(0, -round(cumulative_cash_flow_nm - prev_cumulative_cash_flow_nm, 2))
        else:
            current_cash_demand_nm = 0

        dau = round(df_display.loc[quarter_end_idx, 21], 2)  # DAU
        
        # 图表数据
        dau_quarterly_labels.append(quarter)
        dau_quarterly_data.append(round(dau, 2))
        finance_quarterly_labels.append(quarter)
        finance_quarterly_income.append(round(current_revenue, 2))
        finance_quarterly_cost.append(round(current_cost, 2))
        finance_quarterly_cumulative_profit.append(round(cumulative_profit, 2))
        
        # 表格数据
        quarter_data = {
            'quarter': quarter,
            'end_date': quarter_end_date.strftime('%Y/%m/%d'),
            'cumulative_days': cumulative_days,
            'cumulative_revenue': round(cumulative_revenue, 2),
            'cumulative_cost': round(cumulative_total_cost, 2),
            'ua_cost': round(cumulative_ua_cost, 2),
            'personnel_cost': round(cumulative_personnel_cost, 2),
            'other_cost': round(cumulative_other_cost, 2),
            'cumulative_profit': round(cumulative_profit, 2),
            'cumulative_cash_flow_1m': round(cumulative_cash_flow_1m, 2),
            'cumulative_cash_flow_2m': round(cumulative_cash_flow_2m, 2),
            'current_revenue': round(current_revenue, 2),
            'current_cost': round(current_cost, 2),
            'current_ua_cost': round(current_ua_cost, 2),
            'current_personnel_cost': round(current_personnel_cost, 2),
            'current_other_cost': round(current_other_cost, 2),
            'current_profit': round(current_profit, 2),
            'current_cash_demand_1m': round(current_cash_demand_1m, 2),
            'current_cash_demand_2m': round(current_cash_demand_2m, 2),
            'dau': round(dau, 2)
        }

        if repayment_flag:
            quarter_data['cumulative_cash_flow_nm'] = round(cumulative_cash_flow_nm, 2)
            quarter_data['current_cash_demand_nm'] = round(current_cash_demand_nm, 2)
        
        quarterly_table_data.append(quarter_data)
        
        # 更新前一季度的累计值
        prev_cumulative_revenue = cumulative_revenue
        prev_cumulative_cost = cumulative_total_cost
        prev_cumulative_ua_cost = cumulative_ua_cost
        prev_cumulative_personnel_cost = cumulative_personnel_cost
        prev_cumulative_other_cost = cumulative_other_cost
        prev_cumulative_profit = cumulative_profit
        prev_cumulative_cash_flow_1m = cumulative_cash_flow_1m
        prev_cumulative_cash_flow_2m = cumulative_cash_flow_2m
        prev_cumulative_cash_flow_nm = cumulative_cash_flow_nm

    
    """ 9. 保存DataFrame到本地文件 """
    # 创建data目录（如果不存在）
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    
    # 以项目名称_data.pkl命名保存DataFrame
    safe_project_name = project_name.replace('/', '_').replace('\\', '_').replace(':', '_') if project_name else 'default'
    data_file_path = os.path.join(data_dir, f"{safe_project_name}_data.pkl")
    
    # 保存DataFrame到pickle文件
    with open(data_file_path, 'wb') as f:
        save_data = {
            'df': df,
            'df_extra': df_extra if repayment_flag else None,
            'repayment_months': repayment_months,
            'repayment_flag': repayment_flag
        }
        pickle.dump(save_data, f)
    
    """ 10. 构建返回结果 """
    key_metrics = {
        "max_cash_demand_1m": round(max_cash_demand_1m, 2),
        "max_cash_demand_2m": round(max_cash_demand_2m, 2),
        "dynamic_profit_breakeven_day": dynamic_profit_breakeven_day,
        "cumulative_profit_breakeven_day": cumulative_profit_breakeven_day,
        "cumulative_cash_flow_1m_breakeven_day": cumulative_cash_flow_1m_breakeven_day,
        "cumulative_cash_flow_2m_breakeven_day": cumulative_cash_flow_2m_breakeven_day,
        "day_to_10m_dau": day_to_10m_dau,
        "day_to_2m_dau": day_to_2m_dau,
        "day_to_target_dau": day_to_target_dau,
        "target_dau": target_dau,
        "repayment_months": repayment_months  # 返回回款月数，供前端判断显示
    }

    if repayment_flag:
        key_metrics["max_cash_demand_nm"] = round(max_cash_demand_nm, 2)
        key_metrics["cumulative_cash_flow_nm_breakeven_day"] = cumulative_cash_flow_nm_breakeven_day
    
    results = {
        "key_metrics": key_metrics,
        "charts": {
            "dau_quarterly": {
                "labels": dau_quarterly_labels,
                "data": dau_quarterly_data
            },
            "finance_quarterly": {
                "labels": finance_quarterly_labels,
                "income": finance_quarterly_income,
                "cost": finance_quarterly_cost,
                "cumulative_profit": finance_quarterly_cumulative_profit
            }
        },
        "quarterly_table_data": quarterly_table_data,
        "data_file_saved": True,  # 标记数据已保存
        "project_name": safe_project_name  # 返回处理后的项目名称
    }
    
    return results

""" 使用曲线拟合生成生成完整的ROI和留存率向量 """
def fit_roi_curve_advanced(known_days, known_values, target_days, type):
    # 基础边界处理
    if len(known_days) == 0:
        return [0.0] * target_days

    # 将输入转为 numpy 数组方便计算
    x_data = np.array(known_days)
    y_data = np.array(known_values)

    # 1. 定义拟合函数
    if type == "roi":
        # ROI使用 y = a * x^b + c
        def fit_func(x, a, b, c):
            return a * np.power(x, b) + c
        initial_guess = [y_data[0], 0.5, 0.0]
    else:
        # 留存率使用 y = a * (x + 1)^-b + c
        def fit_func(x, a, b, c):
            return a * np.power(x + 1, -b) + c
        initial_guess = [1.0, 0.5, 0.0]

    # 2. 进行拟合
    try:
        # 添加初始参数猜测和边界条件
        if type == "roi":
            p_opt, _ = curve_fit(
                fit_func,
                x_data,
                y_data,
                p0=initial_guess,
                bounds=([0, 0, -np.inf], [np.inf, 2, np.inf]),
                maxfev=10000
            )
        else:
            # 留存率需要更严格的边界
            p_opt, _ = curve_fit(
                fit_func,
                x_data,
                y_data,
                p0=initial_guess,
                bounds=([0, 0, 0], [1.0, np.inf, 1.0]),
                maxfev=5000
            )
    except Exception as e:
        # 拟合失败回退到线性处理
        print(f"拟合失败: {e}, 使用最后已知值填充")
        return [float(y_data[-1])] * target_days

    # 3. 生成完整数据
    all_days = np.arange(target_days)
    roi_fitted = fit_func(all_days, *p_opt)

    # 4. 确保所有值在合理范围内
    if type == 'roi':
        for i in range(1, len(roi_fitted)):
            roi_fitted[i] = max(roi_fitted[i], roi_fitted[i - 1])
        roi_fitted = np.maximum(roi_fitted, 0)
    else:
        roi_fitted = np.clip(roi_fitted, 0, 1.0)

    return roi_fitted.tolist()

""" 读取保存的DataFrame文件，转换为Excel文件返回 """
def export_daily_excel(project_name: str) -> bytes:
    """
    Args:
        project_name (str): 项目名称，用于定位保存的数据文件

    Returns:
        bytes: Excel文件的二进制数据，用于导出下载
    """
    # 处理项目名称，与保存时保持一致
    safe_project_name = project_name.replace('/', '_').replace('\\', '_').replace(':', '_') if project_name else 'default'
    
    # 构建数据文件路径
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    data_file_path = os.path.join(data_dir, f"{safe_project_name}_data.pkl")
    
    # 检查文件是否存在
    if not os.path.exists(data_file_path):
        raise FileNotFoundError(f"数据文件不存在: {data_file_path}")
    
    # 读取DataFrame
    with open(data_file_path, 'rb') as f:
        loaded_data = pickle.load(f)
    

    df = loaded_data['df']
    df_extra = loaded_data.get('df_extra')
    repayment_months = loaded_data.get('repayment_months', 1)
    repayment_flag = loaded_data.get('repayment_flag', False)
    
    # 创建Excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "每日数据"
    
    # 定义样式
    header_font = Font(bold=True)
    center_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # ==================== 第1行：大分类表头 ====================
    # 列A(1)为空，B-I(2-9)每日，J-S(10-19)累计，T-W(20-23)用户量，X-AA(24-27)现金流，AB-AO(28-41)目标达成周期计算，AP-AQ(42-43)为空
    # 如果n>2，则增加n个月回款相关列
    
    # 计算基础列数
    base_columns = 43
    # 当n>2时增加的6列：滞后n个月现金收入、累计现金缺口、现金流打正标记、倒序求和
    extra_columns = 6 if repayment_flag else 0
    total_columns = base_columns + extra_columns
    
    # 第1行：大分类表头（合并单元格）
    ws.cell(row=1, column=1, value='')  # A列为空
    ws.cell(row=1, column=2, value='每日')  # B列开始
    ws.cell(row=1, column=10, value='累计')  # J列开始
    ws.cell(row=1, column=20, value='用户量')  # T列开始
    ws.cell(row=1, column=24, value='现金流')  # X列开始
    
    # 根据是否有扩展列调整目标达成周期计算的起始列
    if repayment_flag:
        ws.cell(row=1, column=30, value='目标达成周期计算')  # AB列开始
    else:
        ws.cell(row=1, column=28, value='目标达成周期计算')  # AB列开始
    
    # 合并大分类表头单元格（第1行）
    ws.merge_cells('B1:I1')   # 每日: B-I (列2-9)
    ws.merge_cells('J1:S1')   # 累计: J-S (列10-19)
    ws.merge_cells('T1:W1')   # 用户量: T-W (列20-23)
    
    if repayment_flag:
        # 现金流包含额外列：X-AC (列24-29)
        ws.merge_cells('X1:AC1')  # 现金流
        # 目标达成周期计算：AD-AU (列30-47)
        ws.merge_cells('AD1:AU1')
    else:
        ws.merge_cells('X1:AA1')  # 现金流: X-AA (列24-27)
        ws.merge_cells('AB1:AO1') # 目标达成周期计算: AB-AO (列28-41)
    
    # ==================== 第2行：子分类表头 ====================
    # 第2行：子分类表头
    row2_values = ['', '', '', '', '', '', '', '', '',  # A-I (1-9): 空/每日区域
                   '', '', '', '', '', '', '', '', '', '',  # J-S (10-19): 累计区域
                   '', '', '', '',  # T-W (20-23): 用户量区域
                   '滞后1个月', '', '滞后2个月', '']  # X-AA (24-27): 现金流区域
    
    # 当n>2时，添加滞后n个月的子分类
    if repayment_flag:
        row2_values.extend([f'滞后{repayment_months}个月', ''])  # AB-AC (28-29)
    
    # 目标达成周期计算部分
    row2_values.extend([
        '累积利润回本', '', '当期利润回本', '',  # 累积利润回本、当期利润回本
        '累计现金流打正（滞后1个月）', '',  # 累计现金流打正（滞后1个月）
        '累计现金流打正（滞后2个月）', ''  # 累计现金流打正（滞后2个月）
    ])
    
    # 当n>2时，添加累计现金流打正（滞后n个月）
    if repayment_flag:
        row2_values.extend([f'累计现金流打正（滞后{repayment_months}个月）', ''])
    
    # DAU目标部分
    row2_values.extend([
        '1000万DAU', '', '200万DAU', '', '目标DAU', '',  # DAU目标
        '', ''  # ROI of the day, Retention
    ])
    
    for col, value in enumerate(row2_values, 1):
        ws.cell(row=2, column=col, value=value)
    
    # 合并子分类表头单元格（第2行）
    ws.merge_cells('X2:Y2')   # 滞后一个月
    ws.merge_cells('Z2:AA2')  # 滞后两个月
    
    if repayment_flag:
        # n>2时的合并
        ws.merge_cells('AB2:AC2')  # 滞后n个月
        ws.merge_cells('AD2:AE2')  # 累积利润回本
        ws.merge_cells('AF2:AG2')  # 当期利润回本
        ws.merge_cells('AH2:AI2')  # 累计现金流打正（滞后1个月）
        ws.merge_cells('AJ2:AK2')  # 累计现金流打正（滞后2个月）
        ws.merge_cells('AL2:AM2')  # 累计现金流打正（滞后n个月）
        ws.merge_cells('AN2:AO2')  # 1000万DAU
        ws.merge_cells('AP2:AQ2')  # 200万DAU
        ws.merge_cells('AR2:AS2')  # 目标DAU
    else:
        ws.merge_cells('AB2:AC2')  # 累积利润回本
        ws.merge_cells('AD2:AE2')  # 当期利润回本
        ws.merge_cells('AF2:AG2')  # 累计现金流打正（滞后1个月）
        ws.merge_cells('AH2:AI2')  # 累计现金流打正（滞后2个月）
        ws.merge_cells('AJ2:AK2')  # 1000万DAU
        ws.merge_cells('AL2:AM2')  # 200万DAU
        ws.merge_cells('AN2:AO2')  # 目标DAU
    
    # ==================== 第3行：具体列名表头 ====================
    row3_values = ['days', '预期日均消耗', '当日收入', '账面毛利', '总用工成本', '其他运营成本', '其他成本合计', '总成本', '账面净利',
                   '账面ROI', '累计消耗', '累计收入', '总毛利', '累计用工成本', '累计其他运营成本', '累计其他成本', '累计总成本', '总净利', '总净利2',
                   'CPI', 'DNU (万)', 'DAU (万)', 'ARPU',
                   '累计现金收入', '累计现金缺口', '累计现金收入', '累计现金缺口']
    
    # 当n>2时，添加滞后n个月的列名
    if repayment_flag:
        row3_values.extend(['累计现金收入', '累计现金缺口'])
    
    # 目标达成周期计算部分
    row3_values.extend([
        '正累计收益=0，负数=1', '倒序求和', '正当期收益=0，负数=1', '倒序求和', 
        '现金流为正=0，为负=1', '倒序求和', '现金流为正=0，为负=1', '倒序求和'
    ])
    
    # 当n>2时，添加累计现金流打正（滞后n个月）的6列
    if repayment_flag:
        row3_values.extend(['现金流为正=0，为负=1', '倒序求和'])
    
    # DAU目标部分
    row3_values.extend([
        '1000万以上DAU=0，以下=1', '倒序求和', '200万以上DAU=0，以下=1', '倒序求和', 
        'xx万以上DAU=0，以下=1', '倒序求和', 'ROI of the day', 'Retention'
    ])
    
    for col, value in enumerate(row3_values, 1):
        ws.cell(row=3, column=col, value=value)
    
    # ==================== 数据行：从第4行开始 ====================
    for i in range(len(df)):
        row_num = i + 4  # 从第4行开始写入数据
        row_data = [
            int(df.iloc[i, 0]),              # 0: days
            round(df.iloc[i, 1], 2),         # 1: 预期日均消耗
            round(df.iloc[i, 2], 2),         # 2: 当日收入
            round(df.iloc[i, 3], 2),         # 3: 账面毛利
            round(df.iloc[i, 4], 2),         # 4: 总用工成本
            round(df.iloc[i, 5], 2),         # 5: 其他运营成本
            round(df.iloc[i, 6], 2),         # 6: 其他成本合计
            round(df.iloc[i, 7], 2),         # 7: 总成本
            round(df.iloc[i, 8], 2),         # 8: 账面净利
            round(df.iloc[i, 9], 4),         # 9: 账面ROI
            round(df.iloc[i, 10], 2),        # 10: 累计消耗
            round(df.iloc[i, 11], 2),        # 11: 累计收入
            round(df.iloc[i, 12], 2),        # 12: 总毛利
            round(df.iloc[i, 13], 2),        # 13: 累计用工成本
            round(df.iloc[i, 14], 2),        # 14: 累计其他运营成本
            round(df.iloc[i, 15], 2),        # 15: 累计其他成本
            round(df.iloc[i, 16], 2),        # 16: 累计总成本
            round(df.iloc[i, 17], 2),        # 17: 总净利
            round(df.iloc[i, 18], 2),        # 18: 总净利2
            round(df.iloc[i, 19], 4),        # 19: CPI
            round(df.iloc[i, 20], 2),        # 20: DNU (万)
            round(df.iloc[i, 21], 2),        # 21: DAU (万)
            round(df.iloc[i, 22], 4) if not pd.isna(df.iloc[i, 22]) else 0,  # 22: ARPU
            round(df.iloc[i, 23], 2) if not pd.isna(df.iloc[i, 23]) else 0,  # 23: 累计现金收入（滞后1个月）
            round(df.iloc[i, 24], 2) if not pd.isna(df.iloc[i, 24]) else 0,  # 24: 累计现金缺口（滞后1个月）
            round(df.iloc[i, 25], 2) if not pd.isna(df.iloc[i, 25]) else 0,  # 25: 累计现金收入（滞后2个月）
            round(df.iloc[i, 26], 2) if not pd.isna(df.iloc[i, 26]) else 0   # 26: 累计现金缺口（滞后2个月）
        ]
        
        # 当n>2时，添加滞后n个月的现金流数据
        if repayment_flag and df_extra is not None:
            row_data.extend([
                round(df_extra.iloc[i, 0], 2) if not pd.isna(df_extra.iloc[i, 0]) else 0,  # 累计现金收入（滞后n个月）
                round(df_extra.iloc[i, 1], 2) if not pd.isna(df_extra.iloc[i, 1]) else 0   # 累计现金缺口（滞后n个月）
            ])
        
        # 目标达成周期计算数据
        row_data.extend([
            int(df.iloc[i, 27]),             # 27: 正累计收益=0，负数=1
            int(df.iloc[i, 28]),             # 28: 倒序求和
            int(df.iloc[i, 29]),             # 29: 正当期收益=0，负数=1
            int(df.iloc[i, 30]),             # 30: 倒序求和
            int(df.iloc[i, 31]),             # 31: 现金流为正=0，为负=1（滞后1个月）
            int(df.iloc[i, 32]),             # 32: 倒序求和
            int(df.iloc[i, 33]),             # 33: 现金流为正=0，为负=1（滞后2个月）
            int(df.iloc[i, 34])              # 34: 倒序求和
        ])
        
        # 当n>2时，添加累计现金流打正（滞后n个月）数据
        if repayment_flag and df_extra is not None:
            row_data.extend([
                int(df_extra.iloc[i, 2]),    # 现金流为正=0，为负=1（滞后n个月）
                int(df_extra.iloc[i, 3])     # 倒序求和
            ])
        
        # DAU目标和ROI/Retention数据
        row_data.extend([
            int(df.iloc[i, 35]),             # 35: 1000万以上DAU=0，以下=1
            int(df.iloc[i, 36]),             # 36: 倒序求和
            int(df.iloc[i, 37]),             # 37: 200万以上DAU=0，以下=1
            int(df.iloc[i, 38]),             # 38: 倒序求和
            int(df.iloc[i, 39]),             # 39: xx万以上DAU=0，以下=1
            int(df.iloc[i, 40]),             # 40: 倒序求和
            round(df.iloc[i, 41], 4),        # 41: ROI of the day
            round(df.iloc[i, 42], 4)         # 42: Retention
        ])
        
        for col, value in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col, value=value)
    
    # ==================== 应用样式 ====================
    # 设置表头样式
    for row in range(1, 4):
        for col in range(1, total_columns + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = thin_border
    
    # 设置数据行边框
    for row in range(4, len(df) + 4):
        for col in range(1, total_columns + 1):
            ws.cell(row=row, column=col).border = thin_border
    
    # 调整列宽
    for col in range(1, total_columns + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        total_rows = len(df) + 4
        start_row = max(1, total_rows - 100)
        for row in range(start_row, total_rows):
            cell_value = str(ws.cell(row=row, column=col).value or '')
            if len(cell_value) > max_length:
                max_length = len(cell_value)
        adjusted_width = min(max_length + 2, 20)  # 最大宽度20
        ws.column_dimensions[column_letter].width = max(adjusted_width, 8)
    
    # 保存到BytesIO对象
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()