import pandas as pd
import numpy as np
import random
import re
# ==========================================
# 1. 数据加载与预处理
# ==========================================
# 确保文件路径正确
df_pop = pd.read_excel(r"Total Population by Age Group in the West Bank and Gaza.xlsx")
df_life = pd.read_excel(r"Life Tables.xlsx")
df_pal = pd.read_excel(r"Total Cancer Deaths by Age Group – Overall.xlsx")
df_wb = pd.read_excel(r"Total Cancer Deaths by Age Group – West Bank.xlsx")
# 统一列名
for df in [df_pop, df_life, df_pal, df_wb]:
    df.columns = [c.strip() for c in df.columns]
# 清洗 AgeGroup
def clean_age_group(s):
    s = str(s)
    return s.replace(' ', '').replace('-', ' - ').replace('  ', ' ')
for df in [df_pop, df_pal, df_wb]:
    df['AgeGroup'] = df['AgeGroup'].astype(str).apply(clean_age_group)
# 归一化 AgeKey
def normalize_age(x):
    x = str(x).replace(' ', '').replace('(', '').replace(')', '').replace('Years', '').replace('Year', '')
    if 'Less' in x: return '0'
    if 'Above' in x: return '80+'
    return x
df_pop['AgeKey'] = df_pop['AgeGroup'].apply(normalize_age)
df_pal['AgeKey'] = df_pal['AgeGroup'].apply(normalize_age)
df_wb['AgeKey'] = df_wb['AgeGroup'].apply(normalize_age)
# 生命表字典
le_dict = dict(zip(df_life['Age'], df_life['Life Expectancy']))
age_key_to_le_idx = {
    '0': 0, '1-4': 2, '5-19': 12, '20-29': 25, '30-39': 35,
    '40-49': 45, '50-59': 55, '60-69': 65, '70-79': 75, '80+': 85
}
# ==========================================
# 2. 参数定义 (RR模式: 基线为1)
# ==========================================
# U: 漏报率 (Underreporting), 仅作用于西岸数据以还原Total, 以及间接影响加沙Total的基线计算
U_PARAMS = {"min": 0.15, "mode": 0.225, "max": 0.3}
# W: 战争风险比 (War Hazard Ratio), 基线为1
W_PARAMS = {"val": 1.32, "lower": 1.30, "upper": 1.35}
# N: 修正系数 (Multiplier), 基线为1
N_PARAMS = {"val": 1, "lower": 1, "upper": 1}
# ==========================================
# 3. 辅助函数: Log-Normal 抽样
# ==========================================
def get_log_normal_sample(val, lower, upper):
    """
    对 RR/HR 进行 Log-Normal 抽样
    """
    log_val = np.log(val)
    log_lower = np.log(lower)
    log_upper = np.log(upper)
    # 95% CI 宽度约为 3.92 个标准误
    se = (log_upper - log_lower) / 3.92
    log_sim = np.random.normal(log_val, se)
    return np.exp(log_sim)
# ==========================================
# 4. 模拟函数 (严格按照逻辑矩阵执行)
# ==========================================
def calculate_yll_simulation(sex, n_sims=10000):
    # --- 1. 准备人口数据 ---
    pop_col = sex if sex != 'Both sex' else 'Total'

    # 获取各地区人口 (用于计算Rate)
    pop_gaza_sub = df_pop[df_pop['NAME'] == 'Gaza Strip']
    pop_gaza = sum(dict(zip(pop_gaza_sub['AgeKey'], pop_gaza_sub[pop_col])).values())

    pop_wb_sub = df_pop[df_pop['NAME'] == 'West Bank']
    pop_wb = sum(dict(zip(pop_wb_sub['AgeKey'], pop_wb_sub[pop_col])).values())

    pop_pal = pop_gaza + pop_wb

    # --- 2. 准备死亡数据 ---
    if sex == 'Both sex':
        pal_sub = df_pal[df_pal['Cause'] == 'All cancers']
        wb_sub = df_wb[df_wb['Cause'] == 'All cancers']
    else:
        pal_sub = df_pal[(df_pal['Sex'] == sex) & (df_pal['Cause'] == 'All cancers')]
        wb_sub = df_wb[(df_wb['Sex'] == sex) & (df_wb['Cause'] == 'All cancers')]

    # --- 3. 构建基础数据列表 ---
    # 每一行包含：AgeKey, LE, 西岸原始值(Registry), 巴勒斯坦原始值(Total Source)
    data_list = []
    for idx, p_row in pal_sub.iterrows():
        age_key = p_row['AgeKey']
        row_sex = p_row['Sex']
        le_idx = age_key_to_le_idx.get(age_key, 85)
        le_val = le_dict.get(le_idx, le_dict.get(85, 5.0))

        # 匹配西岸数据
        w_row = wb_sub[(wb_sub['AgeKey'] == age_key) & (wb_sub['Sex'] == row_sex)]
        if w_row.empty: continue

        # 西岸原始数据 (Registry)
        d_raw_wb = w_row.iloc[0]['Deaths']

        # 巴勒斯坦原始数据 (Total Burden) - 包含不确定性区间
        d_source_pal_v = p_row['Deaths']
        d_source_pal_l = p_row['Deaths_lower']
        d_source_pal_u = p_row['Deaths_upper']

        data_list.append({
            'd_raw_wb': d_raw_wb,
            'd_source_pal_dist': (d_source_pal_l, d_source_pal_u, d_source_pal_v),
            'le': le_val
        })

    # --- 4. 循环模拟 ---
    # 结果容器结构: Results[Region][Mode][Scenario] = List of sums
    results = {
        'West Bank': {
            'Registry': {'traj': [], 'war': []},
            'Total': {'traj': [], 'war': []}
        },
        'Gaza Strip': {
            'Registry': {'traj': [], 'war': []},  # 将为空 (N/A)
            'Total': {'traj': [], 'war': []}
        },
        'Palestine': {
            'Registry': {'traj': [], 'war': []},  # 将为空 (N/A)
            'Total': {'traj': [], 'war': []}
        }
    }

    for _ in range(n_sims):
        # 参数抽样
        u_sim = random.triangular(U_PARAMS['min'], U_PARAMS['max'], U_PARAMS['mode'])
        w_sim = get_log_normal_sample(W_PARAMS['val'], W_PARAMS['lower'], W_PARAMS['upper'])
        n_sim = get_log_normal_sample(N_PARAMS['val'], N_PARAMS['lower'], N_PARAMS['upper'])

        # 战争综合风险比 (RR)
        rr_war = w_sim * n_sim

        # 单次模拟累加器
        sim_acc = {
            'wb_reg_traj': 0, 'wb_reg_war': 0,
            'wb_tot_traj': 0, 'wb_tot_war': 0,

            'gaza_tot_traj': 0, 'gaza_tot_war': 0,

            'pal_tot_traj': 0, 'pal_tot_war': 0
        }

        for item in data_list:
            le = item['le']

            # 1. 基础数值获取
            val_wb_raw = item['d_raw_wb']  # Registry
            val_pal_source = random.triangular(*item['d_source_pal_dist'])  # Total Source

            # 2. 计算 West Bank (西岸)
            # ------------------------------------------------
            # Mode: Registry
            # Trajectory = Wartime = Raw
            wb_reg = val_wb_raw

            # Mode: Total
            # Trajectory = Wartime = Raw * (1 + U)
            wb_tot = val_wb_raw * (1 + u_sim)

            # 3. 计算 Gaza Strip (加沙)
            # ------------------------------------------------
            # Mode: Registry -> N/A (无法计算)

            # Mode: Total
            # Trajectory = Pal(Source) - WB(Total)
            # 逻辑：加沙基线是总负担减去西岸总负担后的残差
            gaza_tot_traj = max(0, val_pal_source - wb_tot)

            # Wartime = Trajectory * RR
            # 逻辑：在基线基础上施加战争风险
            gaza_tot_war = gaza_tot_traj * rr_war

            # 4. 计算 Palestine (巴勒斯坦)
            # ------------------------------------------------
            # Mode: Registry -> N/A (无法计算)

            # Mode: Total
            # Trajectory = WB(Total) + Gaza(Total Trajectory)
            pal_tot_traj = wb_tot + gaza_tot_traj

            # Wartime = WB(Total) + Gaza(Total Wartime)
            pal_tot_war = wb_tot + gaza_tot_war

            # 5. 累加 YLL (Deaths * Life Expectancy)
            sim_acc['wb_reg_traj'] += wb_reg * le
            sim_acc['wb_reg_war'] += wb_reg * le  # 西岸战争无变化

            sim_acc['wb_tot_traj'] += wb_tot * le
            sim_acc['wb_tot_war'] += wb_tot * le  # 西岸战争无变化

            sim_acc['gaza_tot_traj'] += gaza_tot_traj * le
            sim_acc['gaza_tot_war'] += gaza_tot_war * le

            sim_acc['pal_tot_traj'] += pal_tot_traj * le
            sim_acc['pal_tot_war'] += pal_tot_war * le

        # 存入结果列表
        results['West Bank']['Registry']['traj'].append(sim_acc['wb_reg_traj'])
        results['West Bank']['Registry']['war'].append(sim_acc['wb_reg_war'])

        results['West Bank']['Total']['traj'].append(sim_acc['wb_tot_traj'])
        results['West Bank']['Total']['war'].append(sim_acc['wb_tot_war'])

        results['Gaza Strip']['Total']['traj'].append(sim_acc['gaza_tot_traj'])
        results['Gaza Strip']['Total']['war'].append(sim_acc['gaza_tot_war'])

        results['Palestine']['Total']['traj'].append(sim_acc['pal_tot_traj'])
        results['Palestine']['Total']['war'].append(sim_acc['pal_tot_war'])

    return results, {'West Bank': pop_wb, 'Gaza Strip': pop_gaza, 'Palestine': pop_pal}
# ==========================================
# 5. 执行与输出
# ==========================================
def get_stats(data_list, pop):
    if not data_list: return 0, 0, 0, 0, 0, 0  # 处理空列表
    mean = np.mean(data_list)
    lower = np.percentile(data_list, 2.5)
    upper = np.percentile(data_list, 97.5)
    r_mean = (mean / pop) * 100000
    r_lower = (lower / pop) * 100000
    r_upper = (upper / pop) * 100000
    return mean, lower, upper, r_mean, r_lower, r_upper


# 表头
header = f"{'Region':<12} | {'Sex':<10} | {'Mode':<10} | {'Scenario':<10} | {'YLL (Mean)':<13} | {'95% UI (YLL)':<18} | {'YLL Rate':<12} | {'95% UI (Rate)':<18}"
print(header)
print("-" * len(header))

regions = ['West Bank', 'Gaza Strip', 'Palestine']
sexes = ['Both sex', 'Male', 'Female']
# Modes 列表我们在循环内部根据 Region 动态决定

for sex in sexes:
    # 运行一次模拟，获取该性别下所有 Region 的数据
    res_dict, pop_dict = calculate_yll_simulation(sex)

    for region in regions:
        current_pop = pop_dict[region]

        # 决定该地区可以计算哪些 Mode
        if region == 'West Bank':
            available_modes = ['Registry', 'Total']
        else:
            available_modes = ['Total']  # Gaza 和 Palestine 只有 Total

        for mode in available_modes:
            # 获取数据
            data_traj = res_dict[region][mode]['traj']
            data_war = res_dict[region][mode]['war']

            # 统计 Trajectory
            tm, tl, tu, trm, trl, tru = get_stats(data_traj, current_pop)
            # 统计 Wartime
            wm, wl, wu, wrm, wrl, wru = get_stats(data_war, current_pop)

            print(
                f"{region:<12} | {sex:<10} | {mode:<10} | Trajectory | {tm:<13.1f} | {tl:.0f}-{tu:.0f}            | {trm:<12.1f} | {trl:.1f}-{tru:.1f}")
            print(
                f"{'':<12} | {'':<10} | {'':<10} | Wartime    | {wm:<13.1f} | {wl:.0f}-{wu:.0f}            | {wrm:<12.1f} | {wrl:.1f}-{wru:.1f}")

    print("-" * len(header))