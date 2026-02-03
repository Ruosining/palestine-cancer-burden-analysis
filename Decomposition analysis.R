# ==============================================================================
# 巴勒斯坦地区死亡率分解分析 - 左右大边距优化版
# ==============================================================================

# 第一步：加载必要的R包
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(patchwork)
library(ggtext)
library(openxlsx)
library(stringr)
library(grid)  # 确保加载grid用于圆角矩形

# ==============================================================================
# 【优化版】圆角矩形面板边框定义
# ==============================================================================
element_roundrect <- function(fill = NA, colour = "black", linewidth = 0.5, 
                              linetype = 1, r = unit(3, "mm"), color = NULL) {
  if (!is.null(color)) colour <- color
  structure(
    list(fill = fill, colour = colour, linewidth = linewidth, 
         linetype = linetype, r = r),
    class = c("element_roundrect", "element_rect", "element")
  )
}

element_grob.element_roundrect <- function(element, x = 0.5, y = 0.5, 
                                           width = 1, height = 1, 
                                           vp = NULL, ...) {
  gp <- grid::gpar(
    col = element$colour,
    fill = element$fill,
    lwd = if (is.null(element$linewidth) || element$linewidth == 0) 
      NA else element$linewidth * .pt,  # 使用.ggplot2内部.pt常量
    lty = if (is.null(element$linetype)) 1 else element$linetype
  )
  
  grid::roundrectGrob(
    x = unit(0.5, "npc"),
    y = unit(0.5, "npc"),
    width = unit(1, "npc"),
    height = unit(1, "npc"),
    r = element$r,
    gp = gp,
    vp = vp,
    name = "panel.border.roundrect"
  )
}

# ==============================================================================
# 数据准备部分
# ==============================================================================

# 设置文件路径
file_2022_disease <- "D:/战争社论代码及需要提取的数据/巴勒斯坦地区/地区分解分析/2022-West Bank - Gaza Strip.xlsx"
file_2023_disease <- "D:/战争社论代码及需要提取的数据/巴勒斯坦地区/地区分解分析/2023-West Bank2.xlsx"
file_2022_pop <- "D:/战争社论代码及需要提取的数据/巴勒斯坦地区/地区分解分析/2022_XW_XG.xlsx"
file_2023_pop <- "D:/战争社论代码及需要提取的数据/巴勒斯坦地区/地区分解分析/2023_XW_XG.xlsx"
output_path <- "D:/战争社论代码及需要提取的数据/巴勒斯坦地区/地区分解分析"

# 创建输出目录
if (!dir.exists(output_path)) {
  dir.create(output_path, recursive = TRUE)
}

cat("=== 巴勒斯坦地区死亡率分解分析开始 ===\n")

# ==============================================================================
# 第一步：读取数据
# ==============================================================================

cat("===第一步:读取数据===\n")

# 读取数据
disease_2022 <- read_excel(file_2022_disease, sheet = 1, guess_max = 20000)
disease_2023 <- read_excel(file_2023_disease, sheet = 1, guess_max = 20000)
pop_2022 <- read_excel(file_2022_pop, sheet = 1, guess_max = 20000)
pop_2023 <- read_excel(file_2023_pop, sheet = 1, guess_max = 20000)

cat("疾病数据2022:", nrow(disease_2022), "行,", ncol(disease_2022), "列\n")
cat("疾病数据2023:", nrow(disease_2023), "行,", ncol(disease_2023), "列\n")
cat("人口数据2022:", nrow(pop_2022), "行,", ncol(pop_2022), "列\n")
cat("人口数据2023:", nrow(pop_2023), "行,", ncol(pop_2023), "列\n")

# ==============================================================================
# 第二步：处理人口数据
# ==============================================================================

cat("===第二步:处理人口数据===\n")

process_population_data <- function(data, year){
  data %>%
    select(NAME, YR, AgeGroup, Male, Female) %>%
    rename(
      Region = NAME,
      Year = YR
    ) %>%
    mutate(Data_Year = year)
}

pop_2022_processed <- process_population_data(pop_2022, 2022)
pop_2023_processed <- process_population_data(pop_2023, 2023)

# 合并人口数据
population_data <- bind_rows(pop_2022_processed, pop_2023_processed)

# 将人口数据转换为长格式
population_long <- population_data %>%
  pivot_longer(
    cols = c(Male, Female),
    names_to = "Sex",
    values_to = "Population"
  ) %>%
  select(Year, Region, AgeGroup, Sex, Population, Data_Year)

cat("人口数据处理完成，记录数:", nrow(population_long), "\n")

# ==============================================================================
# 第三步：处理疾病数据
# ==============================================================================

cat("===第三步:处理疾病数据===\n")

process_disease_data <- function(data, year){
  # 选择需要的列
  required_cols <- c("Year","Region","AgeGroup","Sex","Deaths")
  existing_cols <- intersect(required_cols, names(data))
  data <- data[,existing_cols] %>%
    mutate(Data_Year = year)
  return(data)
}

disease_2022_processed <- process_disease_data(disease_2022, 2022)
disease_2023_processed <- process_disease_data(disease_2023, 2023)

# 合并疾病数据
disease_data <- bind_rows(disease_2022_processed, disease_2023_processed)

cat("疾病数据处理完成，记录数:", nrow(disease_data), "\n")

# ==============================================================================
# 第四步：年龄组标准化处理
# ==============================================================================

cat("===第四步:年龄组标准化处理===\n")

# 定义年龄组标准化函数
standardize_age_group <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- NA_character_
  x <- str_trim(x)
  
  # 统一处理各种连字符和空格格式
  x <- str_replace_all(x, "[‐\\–\\—\\-]", "-")  # 统一各种连字符为简单连字符
  
  # 处理人口数据特有的格式：(1- 4) Years -> (1-4) Years
  x <- str_replace_all(x, "\\((\\d+)-\\s*(\\d+)\\)", "(\\1-\\2)")
  
  # 处理疾病数据特有的格式：(1 - 4) Years -> (1-4) Years  
  x <- str_replace_all(x, "\\((\\d+)\\s*-\\s*(\\d+)\\)", "(\\1-\\2)")
  
  # 统一处理其他常见格式
  x <- ifelse(grepl("^Less than 1 Year$", x), "Less than 1 Year", x)
  x <- ifelse(grepl("^Above 79 Years$", x), "Above 79 Years", x)
  
  # 统一括号和空格格式
  x <- str_replace_all(x, "\\s+", " ")  # 多个空格合并为一个
  x <- str_replace_all(x, "\\(\\s*(\\d+-\\d+)\\s*\\)", "(\\1)")  # 统一括号内空格
  
  return(str_trim(x))
}

# 应用标准化函数
population_long$AgeGroup <- standardize_age_group(population_long$AgeGroup)
disease_data$AgeGroup <- standardize_age_group(disease_data$AgeGroup)

# 检查标准化后的年龄组
cat("人口数据年龄组（标准化后）:\n")
print(sort(unique(na.omit(population_long$AgeGroup))))

cat("疾病数据年龄组（标准化后）:\n")
print(sort(unique(na.omit(disease_data$AgeGroup))))

# 定义标准分类（使用标准化后的格式）
standard_regions <- c("Gaza Strip","West Bank")
standard_sex <- c("Male","Female")
standard_years <- c(2022, 2023)
standard_age_groups <- c("Less than 1 Year","(1-4) Years","(5-19) Years","(20-29) Years",
                         "(30-39) Years","(40-49) Years","(50-59) Years","(60-69) Years",
                         "(70-79) Years","Above 79 Years")

cat("标准年龄组:\n")
print(standard_age_groups)

# 检查匹配情况
missing_pop_ages <- setdiff(unique(na.omit(population_long$AgeGroup)), standard_age_groups)
missing_death_ages <- setdiff(unique(na.omit(disease_data$AgeGroup)), standard_age_groups)

if(length(missing_pop_ages) > 0) {
  cat("警告:人口数据中存在未匹配的年龄组:", paste(missing_pop_ages, collapse = ", "), "\n")
}

if(length(missing_death_ages) > 0) {
  cat("警告:疾病数据中存在未匹配的年龄组:", paste(missing_death_ages, collapse = ", "), "\n")
}

# ==============================================================================
# 第五步：合并数据并计算死亡率
# ==============================================================================

cat("===第五步:合并数据并计算死亡率===\n")

# 合并数据-使用完整数据，包括死亡数为0的情况
combined_data <- disease_data %>%
  inner_join(population_long, by = c("Year","Region","AgeGroup","Sex","Data_Year")) %>%
  mutate(
    Rate = ifelse(Population > 0, (Deaths / Population), 0)
  ) %>%
  filter(Region %in% standard_regions,
         Sex %in% standard_sex,
         Year %in% standard_years,
         AgeGroup %in% standard_age_groups) %>%
  select(Year, Region, AgeGroup, Sex, Deaths, Population, Rate) %>%
  arrange(Year, Region, Sex, AgeGroup)

cat("数据合并完成，总记录数:", nrow(combined_data), "\n")

# 检查合并后的数据完整性
cat("合并数据摘要:\n")
cat("- 总记录数:", nrow(combined_data), "\n")
cat("- 地区:", paste(unique(combined_data$Region), collapse = ", "), "\n")
cat("- 性别:", paste(unique(combined_data$Sex), collapse = ", "), "\n")
cat("- 年份:", paste(unique(combined_data$Year), collapse = ", "), "\n")
cat("- 年龄组:", paste(unique(combined_data$AgeGroup), collapse = ", "), "\n")
cat("- 年龄组数量:", length(unique(combined_data$AgeGroup)), "\n")

# 检查零死亡记录
zero_death_count <- sum(combined_data$Deaths == 0, na.rm = TRUE)
cat("- 零死亡记录数:", zero_death_count, 
    "(", round(zero_death_count/nrow(combined_data)*100, 1), "%)\n")

# ==============================================================================
# 第六步：分解分析
# ==============================================================================

cat("===第六步:分解分析-使用所有数据===\n")

decomposition_analysis <- function(data, population_long){
  results <- data.frame(
    analysis_group = character(),
    overall_difference = numeric(),
    age_effect = numeric(),
    population_effect = numeric(),
    rate_effect = numeric(),
    age_percent = numeric(),
    population_percent = numeric(),
    rate_percent = numeric(),
    description = character(),
    stringsAsFactors = FALSE
  )
  
  analyses <- list(
    list(region = "West Bank", sex = "Male", label = "2022-2023_Male_WestBank"),
    list(region = "West Bank", sex = "Female", label = "2022-2023_Female_WestBank"),
    list(region = "Gaza Strip", sex = "Male", label = "2022-2023_Male_Gaza"),
    list(region = "Gaza Strip", sex = "Female", label = "2022-2023_Female_Gaza")
  )
  
  for(analysis in analyses){
    region <- analysis$region
    sex <- analysis$sex
    label <- analysis$label
    
    cat("正在进行", label, "的分解分析...\n")
    
    # 获取该地区该性别的完整人口数据(所有年龄组)
    full_pop_2022 <- population_long %>%
      filter(Year == 2022, Region == region, Sex == sex)
    full_pop_2023 <- population_long %>%
      filter(Year == 2023, Region == region, Sex == sex)
    
    if(nrow(full_pop_2022) == 0 | nrow(full_pop_2023) == 0){
      cat("警告:没有找到", region, sex, "的完整人口数据\n")
      next
    }
    
    # 计算该年该地区的总人口数(使用所有年龄组)
    total_pop_2022 <- sum(full_pop_2022$Population, na.rm = TRUE)
    total_pop_2023 <- sum(full_pop_2023$Population, na.rm = TRUE)
    
    if(total_pop_2022 == 0 | total_pop_2023 == 0){
      cat("警告:总人口为0\n")
      next
    }
    
    # 计算各年龄组的年龄构成(使用完整人口数据)
    full_pop_2022$age_proportion <- full_pop_2022$Population / total_pop_2022
    full_pop_2023$age_proportion <- full_pop_2023$Population / total_pop_2023
    
    # 获取该地区该性别的所有数据(包括死亡数为0的情况)
    region_data <- data %>%
      filter(Region == region, Sex == sex)
    
    if(nrow(region_data) == 0){
      cat("警告:没有找到", region, sex, "的数据\n")
      next
    }
    
    # 获取2022年和2023年数据
    data_2022 <- region_data %>% filter(Year == 2022)
    data_2023 <- region_data %>% filter(Year == 2023)
    
    if(nrow(data_2022) == 0 | nrow(data_2023) == 0){
      cat("警告:缺少某年数据\n")
      next
    }
    
    # 确保两年有相同的年龄组
    common_age_groups <- intersect(data_2022$AgeGroup, data_2023$AgeGroup)
    if(length(common_age_groups) == 0){
      cat("警告:没有共同的年龄组\n")
      next
    }
    
    # 筛选共同年龄组并按相同顺序排序
    data_2022 <- data_2022 %>% 
      filter(AgeGroup %in% common_age_groups) %>% 
      arrange(AgeGroup)
    
    data_2023 <- data_2023 %>% 
      filter(AgeGroup %in% common_age_groups) %>% 
      arrange(AgeGroup)
    
    full_pop_2022 <- full_pop_2022 %>% 
      filter(AgeGroup %in% common_age_groups) %>% 
      arrange(AgeGroup)
    
    full_pop_2023 <- full_pop_2023 %>% 
      filter(AgeGroup %in% common_age_groups) %>% 
      arrange(AgeGroup)
    
    # 验证所有向量的年龄组顺序完全一致
    if(!identical(data_2022$AgeGroup, full_pop_2022$AgeGroup) ||
       !identical(data_2023$AgeGroup, full_pop_2023$AgeGroup)){
      stop(paste("错误：", region, sex, "的年龄组顺序不匹配，数据错位！"))
    }
    
    # 验证向量长度一致性（防御性编程）
    n_groups <- length(common_age_groups)
    if(length(full_pop_2022$age_proportion) != n_groups ||
       length(data_2022$Rate) != n_groups){
      stop(paste("错误：", region, sex, "2022年向量长度不匹配！"))
    }
    
    # 从完整人口数据中提取对应年龄组的年龄构成
    a_2022 <- full_pop_2022$age_proportion
    a_2023 <- full_pop_2023$age_proportion
    
    # 提取死亡率（此时顺序已确保一致）
    r_2022 <- data_2022$Rate
    r_2023 <- data_2023$Rate
    
    # 总人口（标量）
    p_2022 <- total_pop_2022
    p_2023 <- total_pop_2023
    
    # 确保向量长度一致（二次保险）
    n_groups_check <- min(length(a_2022), length(a_2023), length(r_2022), length(r_2023))
    if(n_groups_check == 0){
      cat("错误:没有有效的年龄组数据\n")
      next
    }
    
    # 截断为相同长度（处理可能的边缘情况）
    a_2022 <- a_2022[1:n_groups_check]
    a_2023 <- a_2023[1:n_groups_check]
    r_2022 <- r_2022[1:n_groups_check]
    r_2023 <- r_2023[1:n_groups_check]
    
    # 使用正确的分解公式
    # 年龄效应
    age_effect <- (sum(a_2023 * p_2022 * r_2022, na.rm = TRUE) + 
                     sum(a_2023 * p_2023 * r_2023, na.rm = TRUE)) / 3 +
      (sum(a_2023 * p_2022 * r_2023, na.rm = TRUE) + 
         sum(a_2023 * p_2023 * r_2022, na.rm = TRUE)) / 6 -
      (sum(a_2022 * p_2022 * r_2022, na.rm = TRUE) + 
         sum(a_2022 * p_2023 * r_2023, na.rm = TRUE)) / 3 -
      (sum(a_2022 * p_2022 * r_2023, na.rm = TRUE) + 
         sum(a_2022 * p_2023 * r_2022, na.rm = TRUE)) / 6
    
    # 人口规模效应
    population_effect <- (sum(a_2022 * p_2023 * r_2022, na.rm = TRUE) + 
                            sum(a_2023 * p_2023 * r_2023, na.rm = TRUE)) / 3 +
      (sum(a_2022 * p_2023 * r_2023, na.rm = TRUE) + 
         sum(a_2023 * p_2023 * r_2022, na.rm = TRUE)) / 6 -
      (sum(a_2022 * p_2022 * r_2022, na.rm = TRUE) + 
         sum(a_2023 * p_2022 * r_2023, na.rm = TRUE)) / 3 -
      (sum(a_2022 * p_2022 * r_2023, na.rm = TRUE) + 
         sum(a_2023 * p_2022 * r_2022, na.rm = TRUE)) / 6
    
    # 死亡率效应
    rate_effect <- (sum(a_2022 * p_2022 * r_2023, na.rm = TRUE) + 
                      sum(a_2023 * p_2023 * r_2023, na.rm = TRUE)) / 3 +
      (sum(a_2022 * p_2023 * r_2023, na.rm = TRUE) + 
         sum(a_2023 * p_2022 * r_2023, na.rm = TRUE)) / 6 -
      (sum(a_2022 * p_2022 * r_2022, na.rm = TRUE) + 
         sum(a_2023 * p_2023 * r_2022, na.rm = TRUE)) / 3 -
      (sum(a_2022 * p_2023 * r_2022, na.rm = TRUE) + 
         sum(a_2023 * p_2022 * r_2022, na.rm = TRUE)) / 6
    
    # 四舍五入
    age_effect <- round(age_effect, 3)
    population_effect <- round(population_effect, 3)
    rate_effect <- round(rate_effect, 3)
    
    # 总效应 = A效应 + P效应 + R效应
    overall_difference <- round(age_effect + population_effect + rate_effect, 3)
    
    # 计算百分比贡献
    if(!is.na(overall_difference) && overall_difference != 0){
      age_percent <- round(age_effect / overall_difference * 100, 2)
      population_percent <- round(population_effect / overall_difference * 100, 2)
      rate_percent <- round(rate_effect / overall_difference * 100, 2)
    } else {
      age_percent <- population_percent <- rate_percent <- 0
    }
    
    # 保存结果
    result_row <- data.frame(
      analysis_group = label,
      overall_difference = overall_difference,
      age_effect = age_effect,
      population_effect = population_effect,
      rate_effect = rate_effect,
      age_percent = age_percent,
      population_percent = population_percent,
      rate_percent = rate_percent,
      description = paste(region, sex, "2022-2023死亡率变化分解"),
      stringsAsFactors = FALSE
    )
    
    results <- rbind(results, result_row)
  }
  
  return(results)
}

# 调用分解分析函数
decomposition_results <- decomposition_analysis(combined_data, population_long)

cat("分解分析完成\n")
cat("分解分析结果:\n")
print(decomposition_results)

# ==============================================================================
# 第六步：创建可视化图表 
# ==============================================================================

cat("===第六步:创建可视化图表（左右大边距对称版）===\n")

# 定义配色方案
fill_colors <- c(
  "Population Growth" = "#dde1ef",        # 浅蓝色填充
  "Population Ageing" = "#f4dfdb",        # 浅红色填充
  "Age-Specific Mortality Rates" = "#d6eddd"  # 浅绿色填充
)

border_colors <- c(
  "Population Growth" = "#4e79a7",       # 深蓝色描边
  "Population Ageing" = "#e15759",       # 深红色描边
  "Age-Specific Mortality Rates" = "#59a14f"  # 深绿色描边
)

# 创建单个分解图函数
create_single_decomposition_plot <- function(decomp_data, plot_title = "", show_x_axis = TRUE) {
  # 使用原始效应值
  plot_data <- decomp_data %>%
    select(analysis_group, age_effect, population_effect, rate_effect, overall_difference) %>%
    pivot_longer(cols = c(age_effect, population_effect, rate_effect),
                 names_to = "effect_type", values_to = "effect_value") %>%
    mutate(
      effect_type = case_when(
        effect_type == "age_effect" ~ "Population Ageing",
        effect_type == "population_effect" ~ "Population Growth", 
        effect_type == "rate_effect" ~ "Age-Specific Mortality Rates"
      ),
      effect_type = factor(effect_type, levels = c("Population Growth", "Population Ageing", "Age-Specific Mortality Rates"))
    )
  
  # X轴范围0-280，刻度只显示到240
  x_limits <- c(0, 280)
  x_breaks <- seq(0, 240, by = 40)
  
  p <- ggplot(plot_data, aes(x = effect_value, y = analysis_group)) +
    # 堆叠条形图
    geom_col(aes(fill = effect_type, color = effect_type), 
             position = "stack", width = 0.5, linewidth = 0.5) +
    # 总效应点
    geom_point(aes(x = overall_difference, y = analysis_group, 
                   shape = "Total Effect"), 
               color = "black", size = 2.5) +
    
    # 应用配色
    scale_fill_manual(values = fill_colors) +
    scale_color_manual(values = border_colors) +
    scale_shape_manual(values = c("Total Effect" = 20), name = "") +
    
    # 坐标轴设置
    labs(
      title = plot_title,
      x = if(show_x_axis) "Decomposition (%)" else "",
      y = "",
      fill = "Effect Type",
      color = "Effect Type"
    ) +
    scale_x_continuous(
      breaks = x_breaks, 
      limits = x_limits,
      expand = c(0.01, 0)  # 微小内边距确保不贴边
    ) +
    
    # 零效应参考线
    geom_vline(xintercept = 0, linetype = "dashed", color = "gray50", alpha = 0.7) +
    
    theme_classic() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold", margin = margin(b = 8)),
      axis.text.y = element_text(size = 11, face = "bold"),
      axis.text.x = element_text(size = 10),
      axis.title.x = element_text(size = 12, face = "bold", margin = margin(t = 8)),
      axis.title.y = element_blank(),
      axis.line = element_blank(),
      axis.ticks = element_line(color = "gray50", linewidth = 0.6),
      axis.ticks.length = unit(2, "mm"),
      # 使用圆角边框
      panel.border = element_roundrect(color = "gray50", fill = NA, linewidth = 0.6, r = unit(2, "mm")),
      panel.grid = element_blank(),
      legend.position = "none",
      
      # 左右边距大幅增加至40pt，上下保持紧凑
      plot.margin = margin(t = 5, r = 40, b = 5, l = 40, unit = "pt"),
      
      panel.background = element_rect(fill = "white", color = NA),
      plot.background = element_rect(fill = "white", color = NA)
    )
  
  return(p)
}

# 组合图形函数 
create_complete_combined_plot <- function(decomp_results) {
  male_data <- decomp_results %>% 
    filter(grepl("_Male", analysis_group)) %>%
    mutate(
      analysis_group = case_when(
        grepl("WestBank", analysis_group) ~ "West Bank",
        grepl("Gaza", analysis_group) ~ "Gaza Strip",
        TRUE ~ analysis_group
      )
    )
  
  female_data <- decomp_results %>% 
    filter(grepl("_Female", analysis_group)) %>%
    mutate(
      analysis_group = case_when(
        grepl("WestBank", analysis_group) ~ "West Bank",
        grepl("Gaza", analysis_group) ~ "Gaza Strip",
        TRUE ~ analysis_group
      )
    )
  
  # 男性图形（顶部）- 保持40pt左右边距
  male_plot <- create_single_decomposition_plot(male_data, "Male", show_x_axis = FALSE) +
    theme(
      axis.text.x = element_blank(),
      axis.ticks.x = element_blank(),
      axis.title.x = element_blank(),
      # 保持大边距，去除底部边距以便与女性图拼接
      plot.margin = margin(b = 0, t = 5, r = 40, l = 40, unit = "pt")
    )
  
  # 女性图形（底部）- 保持40pt左右边距
  female_plot <- create_single_decomposition_plot(female_data, "Female", show_x_axis = TRUE) +
    theme(
      # 保持大边距，去除顶部边距以便与男性图拼接
      plot.margin = margin(t = 0, b = 5, r = 40, l = 40, unit = "pt")
    )
  
  # 组合图形 - 使用patchwork拼接，不在组合层面覆盖边距
  combined_plots <- male_plot / female_plot + 
    plot_layout(heights = c(1, 1), guides = "collect") &
    theme(
      legend.position = "bottom",
      legend.box = "horizontal",
      legend.title = element_text(face = "bold", size = 10),
      legend.text = element_text(size = 9),
      legend.key.size = unit(0.4, "cm"),
      legend.margin = margin(t = 10, b = 5),
      # 不在组合层面设置plot.margin，保留单图的边距设置
    )
  
  return(combined_plots)
}

# 生成并保存图形
combined_plot <- create_complete_combined_plot(decomposition_results)

#增加输出图片宽度至12英寸，容纳更大的边距
output_file <- file.path(output_path, "分解分析图_左右大边距对称版.png")
ggsave(output_file, 
       combined_plot, 
       width = 12,      # 从11增至12英寸，确保大边距下内容不拥挤
       height = 6.5,        
       dpi = 600,           
       bg = "white",
       type = "cairo")

cat("图形已保存:", output_file, "\n")
cat("提示：左右边距已设置为40pt，图片宽度增至12英寸以确保最佳视觉效果\n")

# ==============================================================================
# 第七步：保存结果
# ==============================================================================

cat("===第七步:保存结果===\n")

write.csv(decomposition_results, 
          file.path(output_path, "巴勒斯坦地区死亡率分解分析结果_包含所有数据.csv"), 
          row.names = FALSE, fileEncoding = "UTF-8")

write.csv(combined_data, 
          file.path(output_path, "分析数据集_包含所有数据.csv"), 
          row.names = FALSE, fileEncoding = "UTF-8")

write.csv(population_long, 
          file.path(output_path, "人口数据_长格式.csv"), 
          row.names = FALSE, fileEncoding = "UTF-8")

cat("结果已保存\n")

# ==============================================================================
# 第八步：数据摘要
# ==============================================================================

cat("===第八步:数据摘要===\n")

# 计算数据摘要
data_summary <- combined_data %>%
  group_by(Year, Region, Sex) %>%
  summarise(
    年龄组数量 = n(),
    总死亡数 = sum(Deaths, na.rm = TRUE),
    总人口数 = sum(Population, na.rm = TRUE),
    死亡率范围 = paste(round(min(Rate, na.rm = TRUE), 2), "-", round(max(Rate, na.rm = TRUE), 2)),
    零死亡年龄组数 = sum(Deaths == 0, na.rm = TRUE),
    .groups = 'drop'
  )

cat("数据摘要:\n")
print(data_summary)

write.csv(data_summary,
          file.path(output_path, "数据摘要_包含所有数据.csv"),
          row.names = FALSE, fileEncoding = "UTF-8")

# ==============================================================================
# 第九步：生成报告
# ==============================================================================

cat("===第九步:生成报告===\n")

# 创建报告
report <- data.frame(
  项目 = c(
    "分析完成时间",
    "数据来源",
    "分析地区",
    "分析年份", 
    "分析性别",
    "年龄组数量",
    "总记录数",
    "包含零死亡记录数",
    "零死亡记录占比(%)",
    "完成的分解分析",
    "死亡率计算公式",
    "年龄构成计算",
    "总人口计算",
    "数据完整性",
    "分解分析方法",
    "输出文件",
    "图形边距设置"  # 新增
  ),
  结果 = c(
    as.character(Sys.time()),
    "2022-2023年疾病和人口数据",
    paste(standard_regions, collapse = ","),
    paste(standard_years, collapse = ","),
    paste(standard_sex, collapse = ","),
    length(standard_age_groups),
    nrow(combined_data),
    sum(combined_data$Deaths == 0, na.rm = TRUE),
    round(sum(combined_data$Deaths == 0, na.rm = TRUE) / nrow(combined_data) * 100, 2),
    nrow(decomposition_results),
    "死亡率 = (死亡数 / 人口数)",
    "年龄构成 = 年龄组人口 / 该年该地区总人口",
    "总人口 = 所有年龄组人口之和",
    "包含所有数据，尊重客观事实",
    "基于正确的分解公式：总变化 = 人口增长效应 + 人口老龄化效应 + 年龄别死亡率效应",
    "巴勒斯坦地区死亡率分解分析结果_包含所有数据.csv, 分析数据集_包含所有数据.csv, 人口数据_长格式.csv, 数据摘要_包含所有数据.csv, 分解分析图_左右大边距对称版.png",
    "左右对称边距40pt，图片宽度12英寸"
  )
)

write.csv(report,
          file.path(output_path, "分析报告_包含所有数据.csv"),
          row.names = FALSE, fileEncoding = "UTF-8")

cat("分析报告已保存\n")

cat("===分析完成===\n")
cat("注意:本次分析包含了所有数据，包括死亡数为0的情况，以尊重客观事实。\n")
cat("零死亡记录数:", sum(combined_data$Deaths == 0, na.rm = TRUE),
    "占总记录数的", round(sum(combined_data$Deaths == 0, na.rm = TRUE) / nrow(combined_data) * 100, 2), "%\n")

# 显示图形
if (interactive()) {
  print(combined_plot)
}

cat("输出文件已保存至:", output_path, "\n")
cat("=== 巴勒斯坦地区死亡率分解分析完成 ===\n")