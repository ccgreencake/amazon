import numpy as np
from scipy import stats
from scipy.stats import linregress
import matplotlib.pyplot as plt

# 假设的销量数据
# days = np.array([3, 7, 15, 30, 60])  # 时间段（天）
# sales = np.array([293,748,1840,3638,10768])  # 累计销量

# 假设的累积销量数据
cumulative_sales = [293,748,1840,3638,10768]
# 对应的时间段（天）
periods = [3, 7, 15, 30, 60]

# 估算每日销量和时间点
daily_sales = []
days = []
for i in range(len(cumulative_sales)):
    if i == 0:
        daily_sales.append(cumulative_sales[i] / periods[i])
        days.append(periods[i])
    else:
        daily_increase = (cumulative_sales[i] - cumulative_sales[i-1]) / (periods[i] - periods[i-1])
        daily_sales.append(daily_increase)
        days.append(periods[i])

# 将列表转换为numpy数组
days = np.array(days)
daily_sales = np.array(daily_sales)

# 进行线性回归
slope, intercept, r_value, p_value, std_err = linregress(days, daily_sales)

# 绘制原始数据点
plt.scatter(days, daily_sales, color='blue', label='估算的每日销量')

# 绘制拟合的直线
plt.plot(days, intercept + slope * days, 'r-', label='拟合直线')
# 指定字体为支持中文的字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 显示图表
plt.show()