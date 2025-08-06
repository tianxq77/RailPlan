import pandas as pd
import coptpy as cp

# ========== 读取数据 ==========
data_file = "Data.xlsx"
group_info = pd.read_excel(data_file, sheet_name="车组里程修时信息", dtype={'车组号': str})
route_info = pd.read_excel(data_file, sheet_name="待排交路信息", dtype={'交路': str})
initial_schedule = pd.read_excel(data_file, sheet_name="Day1检修上线情况", dtype={'交路': str})
repair_capacity = pd.read_excel(data_file, sheet_name="班组检修能力")
candidate_routes = pd.read_excel(data_file, sheet_name="候选交路", header=None)
recovery_info = pd.read_excel(data_file, sheet_name="车组修后恢复信息")

# ========== 模型创建 ==========
env = cp.Envr()
model = env.createModel("TrainMaintenanceScheduling")

# ========== 集合定义 ==========
days = [col for col in route_info.columns if col.startswith('day')]  # 提取所有天数字段
num_days = len(days)  # 总天数
days = range(num_days)
trains = group_info['车组号'].unique()
routes = route_info['交路'].unique()
repair_types = ['Z', 'L']

# ========== 决策变量定义 ==========
#车组 $i$ 是否在第 $t$ 天执行交路 $j$
x = {(t, d, r): model.addVar(vtype=cp.COPT.BINARY, name=f"x_{t}_{d}_{r}")
     for t in trains for d in days for r in routes}
#车组 $i$ 是否在第 $t$ 天进行 Z \L检修
z = {(t, d, k): model.addVar(vtype=cp.COPT.BINARY, name=f"z_{t}_{d}_{k}")
     for t in trains for d in days for k in repair_types}


# ========== 约束 ==========
# 每车每天最多执行1个任务
for t in trains:
    for d in days:
        model.addConstr(
            cp.quicksum(x[t, d, r] for r in routes if (t, d, r) in x) +
            cp.quicksum(z[t, d, k] for k in repair_types) <= 1,
            name=f"UniqueTask_{t}_{d}"
        )

# 每条交路每天必须被执行
for d in days:
    for r in routes:
        if route_info.loc[route_info['交路'] == r, f'day{d+1}'].values[0] == 1:
            model.addConstr(
                cp.quicksum(x[t, d, r] for t in trains if (t, d, r) in x) == 1,
                name=f"RouteAssigned_{r}_{d}"
            )

# 每日维修能力限制
for d in days:
    for k in repair_types:
        capacity = repair_capacity.loc[repair_capacity['maintlevel'] == k, f'day{d+1}'].values[0]
        model.addConstr(
            cp.quicksum(z[t, d, k] for t in trains) <= capacity,
            name=f"RepairCap_{k}_{d}"
        )

# 限制车组只能执行候选交路
valid_pairs = set()
for _, row in candidate_routes.iterrows():
    train = str(row[0])
    for col in candidate_routes.columns[1:]:
        if pd.notna(row[col]):# 检查是否为空值
            valid_pairs.add((train, row[col]))
for t, d, r in x:
    if (t, r) not in valid_pairs:
        model.addConstr(x[t, d, r] == 0, name=f"InvalidRoute_{t}_{d}_{r}")

# ========== 目标函数 ==========
repair_obj = cp.quicksum(z[t, d, k] for t in trains for d in days for k in repair_types)
model.setObjective(repair_obj, sense=cp.COPT.MINIMIZE)

# ========== 求解 ==========
model.solve()

# ========== 输出结果 ==========
if model.status == cp.COPT.OPTIMAL:
    print("Optimal solution found.\n--- Route Assignments ---")
    for key, var in x.items():
        if var.x > 0.5:
            print(f"Train {key[0]} on Day {key[1]+1} executes Route {key[2]}")
    print("\n--- Maintenance Assignments ---")
    for key, var in z.items():
        if var.x > 0.5:
            print(f"Train {key[0]} on Day {key[1]+1} does Repair {key[2]}")
else:
    print("No optimal solution found.")