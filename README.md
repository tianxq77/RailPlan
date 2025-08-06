# RailPlan
铁路运检计划编排
## 模型构建与求解

### 决策变量


$x_{ijt} \in \{0,1\}$：车组 $i$ 是否在第 $t$ 天执行交路 $j$
$z_{it} \in \{0,1\}$：车组 $i$ 是否在第 $t$ 天进行 Z 检修
$l_{it} \in \{0,1\}$：车组 $i$ 是否在第 $t$ 天进行 L 检修


### 约束条件

1. **每日检修能力限制**：

   $$
   \sum_i z_{it} \leq \text{Z能力}_t, \quad \sum_i l_{it} \leq \text{L能力}_t
   $$

2. **每个交路每天必须有一个车组执行**：

   $$
   \sum_i x_{ijt} = \text{是否需要交路}_{jt}
   $$

3. **每个车组每天最多执行一项任务**：

   $$
   \sum_j x_{ijt} + z_{it} + l_{it} \leq 1
   $$

4. **连续交路安排**：
   若交路 $j_1, j_2$从上到下 属于同一个 R\_ID，要求：

   $$
   x_{ij_1,t} = x_{ij_2,t+1}
   $$

5. **剩余天数和里程限制**：
   车组的Z剩余天数、Z剩余里程和L剩余里程不能为负
    $$
    \text{Z剩余天数} \geq 0,\quad \text{Z/L剩余里程} \geq 0
    $$

6. **候选交路限制**：

   $$
   x_{ijt} = 0 \quad \text{若 } j \notin \text{候选交路}_i
   $$



###  目标函数：

设计为问题一指标的加权组合：

$$
\min \left( \alpha \cdot \text{过修程度指标} + \beta \cdot \text{换车次数指标} + \gamma \cdot \text{检修均衡指标
} \right)
$$

## 源代码说明
- calc_excel()：计算三项评价指标
- export_to_excel()：导出排班结果
- main()：主程序入口，包含模型构建与求解


