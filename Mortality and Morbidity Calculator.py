import pandas as pd
import numpy as np
import random

# ==========================================
# 1. 常量与字典定义
# ==========================================
POPULATION_2023 = {
    "Gaza": {"Both sex": 2257051, "Male": 1143293, "Female": 1113758},
    "West Bank": {"Both sex": 3291406, "Male": 1675256, "Female": 1616150}
}

# 西岸漏报率参数（只用于 West Bank 的 Total 校正）
U_PARAMS = {"min": 0.15, "mode": 0.225, "max": 0.3}

# W值 (治疗中断影响) - RR
w_dict = {
    "Breast cancer": {"overall": {"val": 1.26, "lower": 1.22, "upper": 1.30}},
    "Leukemia": {"overall": {"val": 1.22, "lower": 1.05, "upper": 1.41}},
    "Tracheal, bronchial and lung cancer": {"overall": {"val": 1.44, "lower": 1.35, "upper": 1.54}},
    "Prostate cancer": {"overall": {"val": 1.02, "lower": 0.99, "upper": 1.05}},
    "Ovarian cancer": {"overall": {"val": 1.96, "lower": 1.88, "upper": 2.04}},
    "Pancreatic cancer": {"overall": {"val": 1.59, "lower": 1.26, "upper": 2.02}},
}

# N值（目前弃用 => N=1）
risk_dict = {}

# Region 名称映射（df 里可能是 "Gaza Strip"，人口字典里是 "Gaza"）
REGION_TO_POP_KEY = {
    "Gaza Strip": "Gaza",
    "Gaza": "Gaza",
    "West Bank": "West Bank"
}

# ==========================================
# 2. 辅助函数
# ==========================================
def calculate_w(cancer_type, sex, w_dict):
    if cancer_type in w_dict:
        if sex in w_dict[cancer_type]:
            return w_dict[cancer_type][sex]
        return w_dict[cancer_type].get("overall", {"val": 1.32, "lower": 1.30, "upper": 1.35})
    return {"val": 1.32, "lower": 1.30, "upper": 1.35}

def calculate_n(cancer_type, sex, risk_dict):
    for _, cancers in risk_dict.items():
        if cancer_type in cancers:
            if sex in cancers[cancer_type]:
                return cancers[cancer_type][sex]
            return cancers[cancer_type].get("overall", {"val": 1, "lower": 1, "upper": 1})
    return {"val": 1, "lower": 1, "upper": 1}

def get_stats(sim_list):
    mean_val = np.mean(sim_list)
    lower_ui = np.percentile(sim_list, 2.5)
    upper_ui = np.percentile(sim_list, 97.5)
    return mean_val, lower_ui, upper_ui

def get_log_normal_sample(val, lower, upper):
    log_mean = np.log(val)
    log_se = (np.log(upper) - np.log(lower)) / 3.92
    return np.random.lognormal(log_mean, log_se)

def safe_get_value_row(df_sub):
    """
    尝试从子集读取 Value/Lower/Upper。
    - 如果 Lower/Upper 存在且非空，则返回 (val, low, up)
    - 否则返回 (val, None, None)
    """
    if df_sub.empty:
        return None

    row = df_sub.iloc[0]
    val = row.get("Value", None)
    low = row.get("Lower", None)
    up = row.get("Upper", None)

    def _is_valid(x):
        return x is not None and not (isinstance(x, float) and np.isnan(x))

    if _is_valid(val) and _is_valid(low) and _is_valid(up):
        return float(val), float(low), float(up)
    if _is_valid(val):
        return float(val), None, None
    return None

def sample_triangular_or_point(val, low, up):
    """
    对“数值（Number 或 Rate）”进行抽样：
    - 如果有 low/up：用三角分布（mode=val）
    - 如果没有：当作确定值 val
    """
    if low is None or up is None:
        return val
    return random.triangular(low, up, val)

# ==========================================
# 3. 核心计算引擎（支持地区选择 + Death/Incidence + West Bank官方Rate优先）
# ==========================================
def run_simulation(
    cancer_type,
    sex,
    df,
    target_region,
    mode="Total",          # 'Total' or 'Registry'
    event_type="Death",    # 'Death' or 'Incidence'
    num_simulations=10000
):
    """
    输出 target_region 的 Trajectory（及仅在 Gaza Strip + Death 时输出 Wartime）。

    地区逻辑：
    - target_region == 'West Bank':
        Trajectory Number = West Bank Registry（Total 模式用 u 校正；Registry 不校正）
        Trajectory Rate   = 优先读 West Bank 的 Catalog=Rate（Note=Registry），再按 mode 校正；读不到才用 Number/人口算
        Wartime           = 不计算（返回空列表）
    - target_region == 'Gaza Strip':
        Trajectory Number = Palestine(note_type) - West Bank(Registry, Total时u校正)
        Trajectory Rate   = 用 Trajectory Number/人口算
        Wartime           = 仅当 event_type == 'Death' 才计算

    注意：u 参数只对 West Bank 使用（与你原 U_PARAMS 定义一致）
    """

    # 0) 人口
    pop_key = REGION_TO_POP_KEY.get(target_region)
    if pop_key is None or pop_key not in POPULATION_2023:
        return None, f"Unknown target_region='{target_region}' for population mapping"
    region_pop = POPULATION_2023[pop_key].get(sex)
    if region_pop is None:
        return None, f"No population for target_region='{target_region}', sex='{sex}'"

    # 1) Note 类型（只影响 Palestine 读取）
    note_type = "Total" if mode == "Total" else "Registry"

    # 2) 取 West Bank Registry 的 Number（死亡/发病数）——用于：
    #    - West Bank 自己的 Trajectory Number
    #    - Gaza Strip 的差分扣减项
    wb_num_sub = df[
        (df["Region"] == "West Bank") &
        (df["Note"] == "Registry") &
        (df["Site"] == cancer_type) &
        (df["Sex"] == sex) &
        (df["Type"] == event_type) &
        (df["Catalog"] == "Number")
    ]
    wb_num_row = safe_get_value_row(wb_num_sub)
    if wb_num_row is None:
        return None, f"No West Bank Registry NUMBER data for Type={event_type}"
    wb_num_val, wb_num_low, wb_num_up = wb_num_row

    # 2.1) 取 West Bank Registry 的 Rate（官方率）——仅用于“West Bank 的 Trajectory Rate”
    wb_rate_sub = df[
        (df["Region"] == "West Bank") &
        (df["Note"] == "Registry") &
        (df["Site"] == cancer_type) &
        (df["Sex"] == sex) &
        (df["Type"] == event_type) &
        (df["Catalog"] == "Rate")
    ]
    wb_rate_row = safe_get_value_row(wb_rate_sub)  # 可能 None
    has_wb_official_rate = wb_rate_row is not None

    # 3) 若目标是 Gaza Strip，需要 Palestine（用于差分）
    if target_region == "Gaza Strip":
        pal_sub = df[
            (df["Region"] == "Palestine") &
            (df["Note"] == note_type) &
            (df["Site"] == cancer_type) &
            (df["Sex"] == sex) &
            (df["Type"] == event_type) &
            (df["Catalog"] == "Number")
        ]
        pal_row = safe_get_value_row(pal_sub)
        if pal_row is None:
            return None, f"No Palestine {mode} NUMBER data for Type={event_type}"
        pal_val, pal_low, pal_up = pal_row

    # 4) 参数（仅 Gaza Strip + Death wartime 用）
    compute_wartime = (target_region == "Gaza Strip") and (event_type == "Death")
    if compute_wartime:
        w_params = calculate_w(cancer_type, sex, w_dict)
        n_params = calculate_n(cancer_type, sex, risk_dict)

    # 5) 模拟
    traj_nums, traj_rates = [], []
    wartime_nums, wartime_rates = [], []

    for _ in range(num_simulations):
        # --- A. West Bank Number 基础抽样 ---
        wb_num_sim_base = sample_triangular_or_point(wb_num_val, wb_num_low, wb_num_up)

        # --- B. Total 模式 West Bank 漏报校正（u）---
        if mode == "Total":
            u_sim = random.triangular(U_PARAMS["min"], U_PARAMS["max"], U_PARAMS["mode"])
            wb_num_sim = wb_num_sim_base / (1 - u_sim)
        else:
            u_sim = None
            wb_num_sim = wb_num_sim_base

        # --- C. Trajectory Number（按 target_region）---
        if target_region == "West Bank":
            region_traj_num = max(0, wb_num_sim)
        elif target_region == "Gaza Strip":
            pal_sim = sample_triangular_or_point(pal_val, pal_low, pal_up)
            region_traj_num = max(0, pal_sim - wb_num_sim)
        else:
            return None, f"Unsupported target_region='{target_region}'"

        # --- D. Trajectory Rate（West Bank 优先用官方 Rate）---
        if target_region == "West Bank" and has_wb_official_rate:
            wb_rate_val, wb_rate_low, wb_rate_up = wb_rate_row
            wb_rate_sim_base = sample_triangular_or_point(wb_rate_val, wb_rate_low, wb_rate_up)
            if mode == "Total":
                # rate 与 number 同比放大
                region_traj_rate = wb_rate_sim_base / (1 - u_sim)
            else:
                region_traj_rate = wb_rate_sim_base
        else:
            region_traj_rate = (region_traj_num / region_pop) * 100000

        traj_nums.append(region_traj_num)
        traj_rates.append(region_traj_rate)

        # --- E. Wartime：仅 Gaza Strip + Death ---
        if compute_wartime:
            w_sim = get_log_normal_sample(w_params["val"], w_params["lower"], w_params["upper"])
            n_sim = get_log_normal_sample(n_params["val"], n_params["lower"], n_params["upper"])
            region_wartime_num = region_traj_num * (w_sim * n_sim)

            wartime_nums.append(region_wartime_num)
            wartime_rates.append((region_wartime_num / region_pop) * 100000)

    return {
        "traj_nums": traj_nums,
        "traj_rates": traj_rates,
        "wartime_nums": wartime_nums,   # West Bank 时为空
        "wartime_rates": wartime_rates  # West Bank 时为空
    }, "Success"


# ==========================================
# 4. 执行与输出逻辑（带 Region 列）
# ==========================================
file_path = r"data.xlsx"
df = pd.read_excel(file_path)
df.columns = [c.strip() for c in df.columns]

# =======================
# 改这两个开关
# =======================
TARGET_REGION = "West Bank"     # "West Bank" 或 "Gaza Strip"
EVENT_TYPE = "Death"           # "Death" 或 "Incidence"
# =======================
# # 死亡
# target_cancers = [
#     "All sites",
#     "Tracheal, bronchial and lung cancer",
#     "Colorectal cancer",
#     "Breast cancer",
#     "Brain and central nervous system cancers",
#     "Leukemia",
#     "Stomach cancer",
#     "Pancreatic cancer",
#     "Liver cancer",
#     "Prostate cancer",
#     "Non-Hodgkin lymphoma"
# ]
# sexes = ["Both sex"]
#
# target_cancers = [
#     "All sites",
#     "Tracheal, bronchial and lung cancer",
#     "Colorectal cancer",
#     "Prostate cancer",
#     "Brain and central nervous system cancers",
#     "Leukemia",
#     "Pancreatic cancer",
#     "Stomach cancer",
#     "Liver cancer",
#     "Bladder cancer",
#     "Non-Hodgkin lymphoma"
# ]
# sexes = ["Male"]

target_cancers = [
    "All sites",
    "Breast cancer",
    "Colorectal cancer",
    "Tracheal, bronchial and lung cancer",
    "Brain and central nervous system cancers",
    "Leukemia",
    "Liver cancer",
    "Stomach cancer",
    "Pancreatic cancer",
    "Ovarian cancer",
    "Uterine cancer"
]
sexes = ["Female"]

# 发病
# target_cancers = [
#     "All sites",
#     "Breast cancer",
#     "Colorectal cancer",
#     "Tracheal, bronchial and lung cancer",
#     "Prostate cancer",
#     "Brain and central nervous system cancers",
#     "Leukemia",
#     "Uterine cancer",
#     "Non-melanoma skin cancer",
#     "Non-Hodgkin lymphoma",
#     "Bladder cancer"
# ]
# sexes = ["Both sex"]

# target_cancers = [
#     "All sites",
#     "Tracheal, bronchial and lung cancer",
#     "Colorectal cancer",
#     "Prostate cancer",
#     "Bladder cancer",
#     "Brain and central nervous system cancers",
#     "Leukemia",
#     "Non-melanoma skin cancer",
#     "Stomach cancer",
#     "Pancreatic cancer",
#     "Non-Hodgkin lymphoma"
# ]
# sexes = ["Male"]

# target_cancers = [
#     "All sites",
#     "Breast cancer",
#     "Colorectal cancer",
#     "Uterine cancer",
#     "Brain and central nervous system cancers",
#     "Leukemia",
#     "Tracheal, bronchial and lung cancer",
#     "Non-Hodgkin lymphoma",
#     "Ovarian cancer",
#     "Thyroid cancer",
#     "Non-melanoma skin cancer"
# ]
# sexes = ["Female"]
# 表头
# 统一表头（根据 EVENT_TYPE + 是否会有 Wartime）
# 只有 Gaza Strip + Death 才会有 Wartime
should_print_wartime = (TARGET_REGION == "Gaza Strip") and (EVENT_TYPE == "Death")

warnings = []  # 收集“Registry缺失”等非致命问题

# 表头
if EVENT_TYPE == "Death":
    print(
        f"{'Region':<11} | {'Type':<15} | {'Sex':<8} | {'Mode':<10} | {'Scenario':<10} | "
        f"{'Deaths (Mean)':<13} | {'95% UI (Deaths)':<18} | {'Rate (Mean)':<11} | {'95% UI (Rate)':<18}"
    )
    print("=" * 145)
else:
    print(
        f"{'Region':<11} | {'Type':<15} | {'Sex':<8} | {'Mode':<10} | {'Scenario':<10} | "
        f"{'Incidence (Mean)':<16} | {'95% UI (Incidence)':<21} | {'Rate (Mean)':<11} | {'95% UI (Rate)':<18}"
    )
    print("=" * 150)

for cancer in target_cancers:
    for sex in sexes:
        if (cancer == "Breast cancer" and sex == "Male"):
            continue

        # --- 1) Total 必跑 ---
        res_tot, msg_tot = run_simulation(
            cancer, sex, df,
            target_region=TARGET_REGION,
            mode="Total",
            event_type=EVENT_TYPE
        )

        if not res_tot:
            # Total 都没有：这条才算“致命缺失”，直接报错并跳过
            print(f"{TARGET_REGION:<11} | {cancer[:15]:<15} | {sex:<8} | Total      | ERROR      | {msg_tot}")
            continue

        # Total stats
        td_m, td_l, td_u = get_stats(res_tot["traj_nums"])
        tr_m, tr_l, tr_u = get_stats(res_tot["traj_rates"])

        # --- 2) Registry 尝试跑（缺失则跳过并记录 warning） ---
        res_reg, msg_reg = run_simulation(
            cancer, sex, df,
            target_region=TARGET_REGION,
            mode="Registry",
            event_type=EVENT_TYPE
        )

        # 若 Registry 失败：不影响 Total 输出，只在末尾警告
        has_registry = bool(res_reg)

        if EVENT_TYPE == "Death":
            # 打印 Total Trajectory
            print(
                f"{TARGET_REGION:<11} | {cancer[:15]:<15} | {sex:<8} | Total      | Trajectory | "
                f"{td_m:<13.1f} | {td_l:.0f}-{td_u:.0f}            | {tr_m:<11.1f} | {tr_l:.1f}-{tr_u:.1f}"
            )

            # 只有 Gaza Strip + Death 才打印 Total Wartime（且列表非空）
            if should_print_wartime and len(res_tot.get("wartime_nums", [])) > 0:
                wd_m, wd_l, wd_u = get_stats(res_tot["wartime_nums"])
                wr_m, wr_l, wr_u = get_stats(res_tot["wartime_rates"])
                print(
                    f"{'':<11} | {'':<15} | {'':<8} |            | Wartime    | "
                    f"{wd_m:<13.1f} | {wd_l:.0f}-{wd_u:.0f}            | {wr_m:<11.1f} | {wr_l:.1f}-{wr_u:.1f}"
                )

            # Registry 若存在则打印
            if has_registry:
                rd_m, rd_l, rd_u = get_stats(res_reg["traj_nums"])
                rr_m, rr_l, rr_u = get_stats(res_reg["traj_rates"])
                print(
                    f"{'':<11} | {'':<15} | {'':<8} | Registry   | Trajectory | "
                    f"{rd_m:<13.1f} | {rd_l:.0f}-{rd_u:.0f}            | {rr_m:<11.1f} | {rr_l:.1f}-{rr_u:.1f}"
                )

                if should_print_wartime and len(res_reg.get("wartime_nums", [])) > 0:
                    rwd_m, rwd_l, rwd_u = get_stats(res_reg["wartime_nums"])
                    rwr_m, rwr_l, rwr_u = get_stats(res_reg["wartime_rates"])
                    print(
                        f"{'':<11} | {'':<15} | {'':<8} |            | Wartime    | "
                        f"{rwd_m:<13.1f} | {rwd_l:.0f}-{rwd_u:.0f}            | {rwr_m:<11.1f} | {rwr_l:.1f}-{rwr_u:.1f}"
                    )
            else:
                warnings.append(
                    f"[Registry skipped] Region={TARGET_REGION}, Type={EVENT_TYPE}, Site={cancer}, Sex={sex}: {msg_reg}"
                )

            print("-" * 145)

        else:
            # Incidence：只打印 Trajectory（Total 必有）
            print(
                f"{TARGET_REGION:<11} | {cancer[:15]:<15} | {sex:<8} | Total      | Trajectory | "
                f"{td_m:<16.1f} | {td_l:.0f}-{td_u:.0f}                 | {tr_m:<11.1f} | {tr_l:.1f}-{tr_u:.1f}"
            )

            # Registry 若有则打印；没有就跳过+warning
            if has_registry:
                rd_m, rd_l, rd_u = get_stats(res_reg["traj_nums"])
                rr_m, rr_l, rr_u = get_stats(res_reg["traj_rates"])
                print(
                    f"{'':<11} | {'':<15} | {'':<8} | Registry   | Trajectory | "
                    f"{rd_m:<16.1f} | {rd_l:.0f}-{rd_u:.0f}                 | {rr_m:<11.1f} | {rr_l:.1f}-{rr_u:.1f}"
                )
            else:
                warnings.append(
                    f"[Registry skipped] Region={TARGET_REGION}, Type={EVENT_TYPE}, Site={cancer}, Sex={sex}: {msg_reg}"
                )

            print("-" * 150)

# =======================
# 末尾 Warning 汇总
# =======================
if len(warnings) > 0:
    print("\n" + "!" * 30 + " WARNINGS " + "!" * 30)
    print("以下组合缺少 Registry 数据，已自动跳过 Registry 计算，仅保留 Total：")
    for w in warnings:
        print(" - " + w)
    print("!" * 72)
