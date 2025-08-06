import pandas as pd
import numpy as np

# ==== 参数配置（根据实际情况填写） ====
filename = "Result.xlsx"
recover_Z_day = 66
recover_Z_km = 66000
recover_L_km = 250000

# ==== 1. 读取 Excel 排班结果 ====
df = pd.read_excel(filename, sheet_name="甘特图", header=0)
days = df.columns[1:].tolist()

# ==== 2. 提取检修数据 ====
df_z = df[df['任务'] == 'Z']
df_l = df[df['任务'] == 'L']

# 将检修数据转换为天->车组的映射
z_maint = {day: str(df_z[day].values[0]).split(',') if pd.notna(df_z[day].values[0]) else [] for day in days}
l_maint = {day: str(df_l[day].values[0]).split(',') if pd.notna(df_l[day].values[0]) else [] for day in days}

# ==== 3. 提取车组每天执行的交路 ====
# 过滤掉Z检修和L检修行
df_schedule = df[~df['任务'].isin(['Z', 'L'])].copy()
route_days = days
vehicles = set()

# 构建一个DataFrame：列是日期，行为车组，值是交路
schedule_data = []
for day in days:
    for _, row in df_schedule.iterrows():
        route = row['任务']
        v = row[day]
        if pd.notna(v) and v != '':
            vehicles.add(v)
            schedule_data.append([v, day, route])

df_task = pd.DataFrame(schedule_data, columns=['车组', '日期', '交路'])

# ==== 4. 加载原始里程数据 ====
mileage_df = pd.read_excel("Data.xlsx", sheet_name="车组里程修时信息")
initial_z_day = dict(zip(mileage_df['车组号'], mileage_df['Z剩余天数']))
initial_z_km = dict(zip(mileage_df['车组号'], mileage_df['Z剩余里程']))
initial_l_km = dict(zip(mileage_df['车组号'], mileage_df['L剩余里程']))

route_df = pd.read_excel("Data.xlsx", sheet_name="待排交路信息")
route_distance = dict(zip(route_df['交路'], route_df['distance']))
route_rid = dict(zip(route_df['交路'], route_df['R_ID']))

# ==== 目标 1：过修程度指标 ====
z_pre_days = []
z_pre_kms = []

for day in days:
    for v in z_maint[day]:
        v = v.strip()
        if v == '':
            continue
        z_pre_days.append(initial_z_day[v])
        z_pre_kms.append(initial_z_km[v])

z_overhaul_index = ((sum(z_pre_days) / recover_Z_day) + (sum(z_pre_kms) / recover_Z_km)) / 2

l_pre_kms = []
for day in days:
    for v in l_maint[day]:
        v = v.strip()
        if v == '':
            continue
        l_pre_kms.append(initial_l_km[v])

l_overhaul_index = sum(l_pre_kms) / recover_L_km

# ==== 目标 2：换车次数指标 ====
from collections import defaultdict

r_id_day_route = defaultdict(lambda: defaultdict(str))  # rid -> day -> (车组)
for _, row in df_task.iterrows():
    rid = route_rid[row['交路']]
    r_id_day_route[rid][row['日期']] = row['车组']

change_ratios = []
for rid, assign in r_id_day_route.items():
    sequence = [assign.get(day, '') for day in days]
    change_count = sum(1 for i in range(1, len(sequence)) if sequence[i] != sequence[i - 1] and sequence[i] and sequence[i - 1])
    max_change = len([1 for i in range(1, len(sequence)) if sequence[i] and sequence[i - 1]])
    if max_change > 0:
        change_ratios.append(change_count / max_change)

swap_ratio_index = sum(change_ratios) / len(change_ratios)

# ==== 目标 3：检修均衡指标 ====
z_counts = [len(z_maint[day]) for day in days]
l_counts = [len(l_maint[day]) for day in days]

z_balance_index = np.var(z_counts)
l_balance_index = np.var(l_counts)

# ==== 打印结果 ====
print(f"Z检修过修指标 ≈ {z_overhaul_index:.4f}")
print(f"L检修过修指标 ≈ {l_overhaul_index:.4f}")
print(f"换车次数指标 ≈ {swap_ratio_index:.4f}")
print(f"Z检修均衡指标（方差）≈ {z_balance_index:.4f}")
print(f"L检修均衡指标（方差）≈ {l_balance_index:.4f}")
