#==================== 0. 环境配置与可重复性 ====================
timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
base_dir <- "D:/战争社论代码及需要提取的数据/巴列斯坦总体/ARIMA预测"
output_dir <- file.path(base_dir, paste0("run_", timestamp))

log_file <- file.path(output_dir, "forecast_log.txt")
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
sink(log_file, append = TRUE, split = TRUE)
cat("=== ARIMA-ETS-TBATS预测系统启动 ===\n", as.character(Sys.time()), "\n\n")

set.seed(20241215)
cat("✓ 随机种子已设置: set.seed(20241215)\n")

required_packages <- c("tidyverse", "readxl", "forecast", "tseries", "ggplot2", "zoo", "stringr", "systemfonts")
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    cat(sprintf("安装缺失包: %s\n", pkg))
    install.packages(pkg, dependencies = TRUE)
  }
  library(pkg, character.only = TRUE)
}
cat("✓ 所有依赖包加载成功\n")

get_safe_font <- function() {
  available <- tryCatch(systemfonts::system_fonts()$family, error = function(e) NULL)
  if (is.null(available)) return("sans")
  safe_fonts <- c("Times New Roman", "Liberation Serif", "DejaVu Serif", "serif")
  for (font in safe_fonts) {
    if (font %in% available) return(font)
  }
  return("sans")
}
safe_font <- get_safe_font()
cat(sprintf("✓ 字体设置: %s\n", safe_font))

#==================== 1. 基础配置 ====================
input_file <- file.path(base_dir, "总体巴勒斯坦预测表.xlsx")

disease_name_mapping <- list(
  "Death_Male_Trabronchial" = "Tracheal, Bronchial, and Lung Cancer (Male, Death)",
  "Death_Male_Colorectal" = "Colorectal Cancer (Male, Death)",
  "Death_Male_Prostate" = "Prostate Cancer (Male, Death)",
  "Death_Female_Breast" = "Breast Cancer (Female, Death)",
  "Death_Female_Colorectal" = "Colorectal Cancer (Female, Death)",
  "Death_Female_Trabronchial" = "Tracheal, Bronchial, and Lung Cancer (Female, Death)",
  "Incidence_Male_Colorectal" = "Colorectal Cancer (Male, Incidence)",
  "Incidence_Male_Trabronchial" = "Tracheal, Bronchial, and Lung Cancer (Male, Incidence)",
  "Incidence_Male_Other" = "Other Non-malignant Tumors (Male, Incidence)",
  "Incidence_Female_Colorectal" = "Colorectal Cancer (Female, Incidence)",
  "Incidence_Female_Breast" = "Breast Cancer (Female, Incidence)",
  "Incidence_Female_Other" = "Other Non-malignant Tumors (Female, Incidence)"
)

#==================== 2. 数据准备（显式构建12类别） ====================
prepare_epidemic_data <- function(file_path, year_start, year_end, window_name) {
  cat(sprintf("\n=== 数据准备: %s (%d-%d) ===\n", window_name, year_start, year_end))
  
  if (!file.exists(file_path)) {
    stop(paste("错误: 文件不存在 -", file_path))
  }
  
  raw_data <- readxl::read_excel(file_path)
  required_columns <- c("measure_name", "sex_name", "cause_name", "year", "val")
  missing_columns <- setdiff(required_columns, colnames(raw_data))
  
  if (length(missing_columns) > 0) {
    stop(paste("数据文件缺少必要的列:", paste(missing_columns, collapse = ",")))
  }
  
  total_records <- nrow(raw_data)
  cat(sprintf("  原始数据记录数: %d\n", total_records))
  
  categories_to_keep <- data.frame(
    measure_name = c(rep("死亡", 6), rep("发病率", 6)),
    sex_name = c("男", "男", "男", "女", "女", "女", "男", "男", "男", "女", "女", "女"),
    cause_name = c(
      "气管、支气管和肺癌", "结肠和直肠癌", "前列腺癌", "乳腺癌", "结肠和直肠癌", 
      "气管、支气管和肺癌", "结肠和直肠癌", "气管、支气管和肺癌", "其他非恶性肿瘤", 
      "结肠和直肠癌", "乳腺癌", "其他非恶性肿瘤"
    )
  )
  
  data <- raw_data %>%
    filter(year >= year_start & year <= year_end) %>%
    inner_join(categories_to_keep, by = c("measure_name", "sex_name", "cause_name")) %>%
    mutate(
      val = suppressWarnings(as.numeric(val))
    ) %>%
    group_by(measure_name, sex_name, cause_name) %>%
    mutate(
      val = zoo::na.approx(val, rule = 2),
      val = ifelse(val < 0, 0, val),
    ) %>%
    ungroup() %>%
    filter(!is.na(val))
  
  # 报告零值情况
  zero_values <- raw_data %>% filter(val == 0)
  if (nrow(zero_values) > 0) {
    cat(sprintf("  注意: 数据包含%d个零值（保留用于分析）\n", nrow(zero_values)))
  }
  
  cat(sprintf("  清洗后记录数: %d\n", nrow(data)))
  
  # 显式构建12个类别（确保与mapping完全对应）
  detailed_categories <- list()
  detailed_categories[["Death_Male_Trabronchial"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "男", cause_name == "气管、支气管和肺癌")
  detailed_categories[["Death_Male_Colorectal"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "男", cause_name == "结肠和直肠癌")
  detailed_categories[["Death_Male_Prostate"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "男", cause_name == "前列腺癌")
  detailed_categories[["Death_Female_Breast"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "女", cause_name == "乳腺癌")
  detailed_categories[["Death_Female_Colorectal"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "女", cause_name == "结肠和直肠癌")
  detailed_categories[["Death_Female_Trabronchial"]] <- data %>% 
    filter(measure_name == "死亡", sex_name == "女", cause_name == "气管、支气管和肺癌")
  detailed_categories[["Incidence_Male_Colorectal"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "男", cause_name == "结肠和直肠癌")
  detailed_categories[["Incidence_Male_Trabronchial"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "男", cause_name == "气管、支气管和肺癌")
  detailed_categories[["Incidence_Male_Other"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "男", cause_name == "其他非恶性肿瘤")
  detailed_categories[["Incidence_Female_Colorectal"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "女", cause_name == "结肠和直肠癌")
  detailed_categories[["Incidence_Female_Breast"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "女", cause_name == "乳腺癌")
  detailed_categories[["Incidence_Female_Other"]] <- data %>% 
    filter(measure_name == "发病率", sex_name == "女", cause_name == "其他非恶性肿瘤")
  
  # 检查空类别
  empty_categories <- names(detailed_categories)[sapply(detailed_categories, function(x) nrow(x) == 0)]
  if (length(empty_categories) > 0) {
    cat(sprintf("  警告: 空类别（数据问题）: %s\n", paste(empty_categories, collapse = ", ")))
  }
  
  cat(sprintf("✓ %s完成: %d个类别, %d条有效记录\n\n", 
              window_name, length(detailed_categories), sum(sapply(detailed_categories, function(x) nrow(x)))))
  
  return(list(categories = detailed_categories, name_mapping = disease_name_mapping))
}

#==================== 3. 自动ARIMA预测 ====================
auto_arima_forecast <- function(series, series_years, forecast_years = 7, model_name = "Auto_ARIMA") {
  cat("  → 自动ARIMA:", model_name, "\n")
  
  tryCatch({
    validate_inputs(series, series_years)
    window_start <- as.integer(min(series_years))
    if (max(series_years) != 2023) stop("数据终点必须是2023年")
    
    full_years <- window_start:2023
    aligned_data <- data.frame(year = full_years) %>%
      left_join(data.frame(year = series_years, val = series), by = "year") %>%
      mutate(val = zoo::na.approx(val, rule = 2))
    
    ts_data <- ts(aligned_data$val, start = window_start, frequency = 1)
    cat(sprintf("    数据长度: %d | 窗口: %d-%d\n", length(ts_data), window_start, 2023))
    
    kpss_test <- tseries::kpss.test(ts_data)
    adf_test <- tseries::adf.test(ts_data)
    cat(sprintf("    KPSS检验: p=%.3f (拒绝原假设=非平稳) | ADF检验: p=%.3f (拒绝原假设=平稳)\n", 
                kpss_test$p.value, adf_test$p.value))
    
    fit <- auto.arima(ts_data, stepwise = FALSE, approximation = FALSE, trace = FALSE)
    order <- arimaorder(fit)
    cat(sprintf("    阶数: ARIMA(%d,%d,%d) AIC=%.2f\n", order[1], order[2], order[3], fit$aic))
    
    final_forecast <- forecast(fit, h = forecast_years, level = c(90, 95))
    
    return(list(
      forecast_values = as.numeric(final_forecast$mean),
      forecast_lower_95 = as.numeric(final_forecast$lower[, 2]),
      forecast_upper_95 = as.numeric(final_forecast$upper[, 2]),
      forecast_lower_90 = as.numeric(final_forecast$lower[, 1]),
      forecast_upper_90 = as.numeric(final_forecast$upper[, 1]),
      forecast_years = (2023 + 1):(2023 + forecast_years),
      model = fit,
      params = list(model_type = "Auto ARIMA", order = order, aic = fit$aic),
      fitted_values = fitted(fit),
      actual_values = as.numeric(ts_data),
      residuals = residuals(fit),
      historical_years = time(ts_data),
      historical_values = as.numeric(ts_data),
      kpss_p = kpss_test$p.value,
      adf_p = adf_test$p.value
    ))
    
  }, error = function(e) {
    warning(paste("自动ARIMA失败:", e$message))
    return(NULL)
  })
}

#==================== 4. 手动网格搜索ARIMA ====================
manual_arima_forecast <- function(series, series_years, forecast_years = 7, model_name = "Manual_ARIMA") {
  cat("  → 手动ARIMA:", model_name, "\n")
  
  tryCatch({
    validate_inputs(series, series_years)
    ts_data <- ts(series, start = min(series_years), frequency = 1)
    cat(sprintf("    数据长度: %d | 年份: %d-%d\n", length(series), min(series_years), max(series_years)))
    
    best_aicc <- Inf
    best_model <- NULL
    
    for (p in 0:2) {
      for (d in 0:2) {
        for (q in 0:2) {
          tryCatch({
            fit <- Arima(ts_data, order = c(p, d, q))
            if (fit$aicc < best_aicc) {
              best_aicc <- fit$aicc
              best_model <- fit
            }
          }, error = function(e) {})
        }
      }
    }
    
    if (is.null(best_model)) stop("未找到合适的手动ARIMA模型")
    
    order <- arimaorder(best_model)
    cat(sprintf("    最优阶数: (%d,%d,%d) AICc=%.2f\n", order[1], order[2], order[3], best_model$aicc))
    
    final_forecast <- forecast(best_model, h = forecast_years, level = c(90, 95))
    
    return(list(
      forecast_values = as.numeric(final_forecast$mean),
      forecast_lower_95 = as.numeric(final_forecast$lower[, 2]),
      forecast_upper_95 = as.numeric(final_forecast$upper[, 2]),
      forecast_lower_90 = as.numeric(final_forecast$lower[, 1]),
      forecast_upper_90 = as.numeric(final_forecast$upper[, 1]),
      forecast_years = (max(series_years) + 1):(max(series_years) + forecast_years),
      model = best_model,
      params = list(model_type = "Manual ARIMA", order = order, aic = best_model$aic),
      fitted_values = fitted(best_model),
      actual_values = series,
      residuals = residuals(best_model),
      historical_years = series_years,
      historical_values = series
    ))
    
  }, error = function(e) {
    warning(paste("手动ARIMA失败:", e$message))
    return(NULL)
  })
}

#==================== 5. ETS指数平滑 ====================
ets_forecast <- function(series, series_years, forecast_years = 7, model_name = "ETS") {
  cat("  → ETS:", model_name, "\n")
  
  tryCatch({
    validate_inputs(series, series_years)
    ts_data <- ts(series, start = min(series_years), frequency = 1)
    
    ets_fit <- ets(ts_data, 
                   lambda = "auto",
                   additive.only = TRUE,
                   damped = NULL,
                   restrict = TRUE,
                   biasadj = TRUE)
    
    cat(sprintf("    模型: %s | λ=%.2f | AIC=%.2f\n", ets_fit$method, ets_fit$lambda, ets_fit$aic))
    
    ets_forecast_obj <- forecast(ets_fit, h = forecast_years, level = c(90, 95),
                                 lambda = ets_fit$lambda, biasadj = TRUE)
    
    return(list(
      forecast_values = as.numeric(ets_forecast_obj$mean),
      forecast_lower_95 = as.numeric(ets_forecast_obj$lower[, 2]),
      forecast_upper_95 = as.numeric(ets_forecast_obj$upper[, 2]),
      forecast_lower_90 = as.numeric(ets_forecast_obj$lower[, 1]),
      forecast_upper_90 = as.numeric(ets_forecast_obj$upper[, 1]),
      forecast_years = (max(series_years) + 1):(max(series_years) + forecast_years),
      model = ets_fit,
      params = list(model_type = "ETS", ets_method = ets_fit$method, aic = ets_fit$aic, lambda = ets_fit$lambda),
      fitted_values = fitted(ets_fit),
      actual_values = series,
      residuals = residuals(ets_fit),
      historical_years = series_years,
      historical_values = series
    ))
    
  }, error = function(e) {
    warning(paste("ETS失败:", e$message))
    return(NULL)
  })
}

#==================== 6. TBATS预测（无季节性） ====================
tbats_forecast <- function(series, series_years, forecast_years = 7, model_name = "TBATS") {
  cat("  → TBATS:", model_name, "\n")
  
  tryCatch({
    validate_inputs(series, series_years)
    ts_data <- ts(series, start = min(series_years), frequency = 1)
    
    # TBATS模型：无季节性周期（seasonal.periods = NULL）
    tbats_fit <- tbats(ts_data,
                       use.box.cox = TRUE,      # 自动Box-Cox变换
                       use.trend = TRUE,        # 包含趋势项
                       use.damped.trend = TRUE, # 包含阻尼趋势
                       seasonal.periods = NULL) # 无季节性周期
    
    cat(sprintf("    模型: %s | λ=%.2f | AIC=%.2f\n", 
                tbats_fit$method, tbats_fit$lambda, tbats_fit$AIC))
    
    tbats_forecast_obj <- forecast(tbats_fit, h = forecast_years, level = c(90, 95))
    
    return(list(
      forecast_values = as.numeric(tbats_forecast_obj$mean),
      forecast_lower_95 = as.numeric(tbats_forecast_obj$lower[, 2]),
      forecast_upper_95 = as.numeric(tbats_forecast_obj$upper[, 2]),
      forecast_lower_90 = as.numeric(tbats_forecast_obj$lower[, 1]),
      forecast_upper_90 = as.numeric(tbats_forecast_obj$upper[, 1]),
      forecast_years = (max(series_years) + 1):(max(series_years) + forecast_years),
      model = tbats_fit,
      params = list(model_type = "TBATS", tbats_method = tbats_fit$method, aic = tbats_fit$AIC, lambda = tbats_fit$lambda),
      fitted_values = fitted(tbats_fit),
      actual_values = series,
      residuals = residuals(tbats_fit),
      historical_years = series_years,
      historical_values = series
    ))
    
  }, error = function(e) {
    warning(paste("TBATS失败:", e$message))
    return(NULL)
  })
}

#==================== 7. 输入验证辅助函数 ====================
validate_inputs <- function(series, series_years) {
  if (length(series) != length(series_years)) {
    stop("序列长度与年份长度不匹配")
  }
  if (any(is.na(series))) {
    warning("输入序列包含NA值")
  }
  if (length(unique(series_years)) != length(series_years)) {
    stop("年份存在重复值")
  }
  return(TRUE)
}

#==================== 8. 双窗口预测执行 ====================
perform_dual_window_predictions <- function(data_categories, forecast_years = 7) {
  cat("\n=== 双窗口预测执行 ===\n")
  cat("  长窗口: 1990-2023 (四次预测)\n")
  cat("  短窗口: 2009-2023 (四次预测)\n\n")
  
  pred_functions <- list(
    auto_arima = auto_arima_forecast,
    manual_arima = manual_arima_forecast,
    ets = ets_forecast,
    tbats = tbats_forecast
  )
  
  all_results_long <- list()
  all_results_short <- list()
  skipped_units <- list()
  
  for (category_name in names(data_categories)) {
    cat(paste0(strrep("=", 60), collapse = ""), "\n")
    cat(sprintf("处理: %s\n", category_name))
    
    category_data <- data_categories[[category_name]]
    if (nrow(category_data) == 0) {
      cat("  警告: 无数据，跳过\n")
      skipped_units[[category_name]] <- "无数据"
      next
    }
    
    category_data <- category_data %>% arrange(year)
    series <- category_data$val
    series_years <- category_data$year
    
    # 长窗口: 1990-2023
    cat(" → 长窗口 (1990-2023):\n")
    idx_long <- which(series_years >= 1990 & series_years <= 2023)
    results_long <- list()
    
    if (length(idx_long) >= 10) {
      for (model_id in names(pred_functions)) {
        model_name <- paste0(category_name, "_", model_id, "_1990_2023")
        results_long[[model_id]] <- pred_functions[[model_id]](
          series = series[idx_long],
          series_years = series_years[idx_long],
          forecast_years = forecast_years,
          model_name = model_name
        )
      }
    } else {
      cat("  警告: 1990-2023数据不足\n")
    }
    
    if (sum(!sapply(results_long, is.null)) >= 1) {
      all_results_long[[category_name]] <- list(
        category_data = category_data %>% filter(year >= 1990),
        predictions = results_long,
        metadata = list(successful_count = sum(!sapply(results_long, is.null)), window = "1990-2023")
      )
    }
    
    # 短窗口: 2009-2023
    cat(" → 短窗口 (2009-2023):\n")
    idx_short <- which(series_years >= 2009 & series_years <= 2023)
    results_short <- list()
    
    if (length(idx_short) >= 10) {
      for (model_id in names(pred_functions)) {
        model_name <- paste0(category_name, "_", model_id, "_2009_2023")
        results_short[[model_id]] <- pred_functions[[model_id]](
          series = series[idx_short],
          series_years = series_years[idx_short],
          forecast_years = forecast_years,
          model_name = model_name
        )
      }
    } else {
      cat("  警告: 2009-2023数据不足\n")
    }
    
    if (sum(!sapply(results_short, is.null)) >= 1) {
      all_results_short[[category_name]] <- list(
        category_data = category_data %>% filter(year >= 2009),
        predictions = results_short,
        metadata = list(successful_count = sum(!sapply(results_short, is.null)), window = "2009-2023")
      )
    }
    
    cat(sprintf(" ✓ 成功: 长窗口 %d/4 | 短窗口 %d/4\n\n",
                sum(!sapply(results_long, is.null)), sum(!sapply(results_short, is.null))))
  }
  
  cat(sprintf("\n=== 双窗口预测完成 ===\n"))
  cat(sprintf("  长窗口成功: %d/%d 个单位\n", length(all_results_long), length(data_categories)))
  cat(sprintf("  短窗口成功: %d/%d 个单位\n", length(all_results_short), length(data_categories)))
  
  return(list(
    long_window = all_results_long,
    short_window = all_results_short,
    skipped_units = skipped_units
  ))
}

#==================== 9. 金标准模型评估（含MAPE） =================###
comprehensive_evaluation <- function(pred, series) {
  if (!is.null(pred$model)) {
    T <- length(pred$actual_values)
    
    # 信息准则计算
    if (pred$params$model_type == "ETS") {
      k <- length(pred$model$par)
      log_lik <- pred$model$loglik
    } else if (grepl("ARIMA", pred$params$model_type)) {
      order_vec <- arimaorder(pred$model)
      k <- sum(order_vec[c(1, 3)]) + 1  # AR+MA+方差
      # 仅非差分模型(d=0)且含截距时+1 (Burnham & Anderson 2002)
      if (order_vec[2] == 0 && "intercept" %in% names(pred$model$coef)) {
        k <- k + 1
      }
      log_lik <- pred$model$loglik
    } else if (pred$params$model_type == "TBATS") {
      k <- length(pred$model$parameters)
      log_lik <- -pred$model$likelihood  # TBATS返回负对数似然
    }
    
    if (!is.null(log_lik) && !is.na(log_lik)) {
      aic <- -2 * log_lik + 2 * k
      aicc <- aic + (k * (k + 1)) / max(1, T - k - 1)
      bic <- aic + k * (log(T) - 2)
    } else {
      aic <- aicc <- bic <- NA
    }
  } else {
    aic <- aicc <- bic <- NA
  }
  
  # 残差白噪声检验
  lb_test <- Box.test(pred$residuals, type = "Ljung-Box", lag = min(5, floor(length(series)/5)))
  
  # MAPE计算（排除实际值≤0的点以避免无穷大）
  valid_idx <- pred$actual_values > 0
  if (sum(valid_idx) == 0) {
    # 如果所有实际值都≤0，返回NA
    mape <- NA
  } else {
    mape <- mean(abs((pred$actual_values[valid_idx] - pred$fitted_values[valid_idx]) / pred$actual_values[valid_idx])) * 100
  }
  
  return(list(
    AIC = aic,
    AICc = aicc,
    BIC = bic,
    Residual_LB_P = lb_test$p.value,
    MAPE = mape
  ))
}

#==================== 10. 完整模型评估矩阵生成（重构） ====================
compare_and_generate_matrix <- function(forecast_results, data_categories) {
  cat("\n=== 模型比较矩阵生成（全模型评估） ===\n")
  cat("  评估模型: 8模型/类别 = 4算法×2窗口\n")
  cat("  评估指标: AIC, AICc, BIC, MAPE, Ljung-Box p值\n\n")
  
  model_specs <- data.frame(
    model_id = c("auto_arima", "manual_arima", "ets", "tbats"),
    model_type_long = c("Auto ARIMA (1990-2023)", "Manual ARIMA (1990-2023)",
                        "ETS (1990-2023)", "TBATS (1990-2023)"),
    model_type_short = c("Auto ARIMA (2009-2023)", "Manual ARIMA (2009-2023)",
                         "ETS (2009-2023)", "TBATS (2009-2023)")
  )
  
  all_results_long <- forecast_results$long_window
  all_results_short <- forecast_results$short_window
  
  detailed_comparison <- data.frame()
  all_predictions <- list()
  row_index_counter <- 1
  
  # 遍历所有类别和模型
  for (category_name in names(data_categories)) {
    cat(sprintf("\n[%s] 评估中:\n", category_name))
    
    series <- data_categories[[category_name]]$val
    available_preds <- list()
    pred_info <- list()
    
    # 收集所有窗口的预测结果
    for (model_id in model_specs$model_id) {
      # 长窗口
      if (category_name %in% names(all_results_long)) {
        long_pred <- all_results_long[[category_name]]$predictions[[model_id]]
        if (!is.null(long_pred)) {
          available_preds[[length(available_preds) + 1]] <- long_pred
          pred_info[[length(pred_info) + 1]] <- list(
            model_type = model_specs$model_type_long[model_specs$model_id == model_id],
            window = "1990-2023",
            model_id = paste0(model_id, "_1990"),
            row_index = row_index_counter
          )
          row_index_counter <- row_index_counter + 1
        }
      }
      # 短窗口
      if (category_name %in% names(all_results_short)) {
        short_pred <- all_results_short[[category_name]]$predictions[[model_id]]
        if (!is.null(short_pred)) {
          available_preds[[length(available_preds) + 1]] <- short_pred
          pred_info[[length(pred_info) + 1]] <- list(
            model_type = model_specs$model_type_short[model_specs$model_id == model_id],
            window = "2009-2023",
            model_id = paste0(model_id, "_2009"),
            row_index = row_index_counter
          )
          row_index_counter <- row_index_counter + 1
        }
      }
    }
    
    if (length(available_preds) == 0) {
      cat(sprintf("  跳过: 无有效预测\n"))
      next
    }
    
    # 计算所有模型的评价指标
    all_scores <- lapply(available_preds, function(p) comprehensive_evaluation(p, series))
    
    # 构建详细比较表
    for (i in 1:length(available_preds)) {
      row_data <- data.frame(
        Row_Index = pred_info[[i]]$row_index,
        Category = category_name,
        Model_Type = pred_info[[i]]$model_type,
        Time_Window = pred_info[[i]]$window,
        Model_ID = pred_info[[i]]$model_id,
        AIC = round(all_scores[[i]]$AIC, 2),
        AICc = round(all_scores[[i]]$AICc, 2),
        BIC = round(all_scores[[i]]$BIC, 2),
        MAPE = round(all_scores[[i]]$MAPE, 2),
        Residual_LB_P = round(all_scores[[i]]$Residual_LB_P, 3),
        stringsAsFactors = FALSE
      )
      detailed_comparison <- rbind(detailed_comparison, row_data)
      all_predictions[[pred_info[[i]]$row_index]] <- available_preds[[i]]
      
      cat(sprintf("  记录模型: %s %s (AICc=%.2f, MAPE=%.2f%%)\n",
                  pred_info[[i]]$model_type, pred_info[[i]]$window,
                  all_scores[[i]]$AICc, all_scores[[i]]$MAPE))
    }
  }
  
  cat(sprintf("\n✓ 评估完成: %d个模型记录\n", nrow(detailed_comparison)))
  
  return(list(
    detailed_comparison = detailed_comparison,
    all_predictions = all_predictions,
    all_scores = all_scores
  ))
}

#==================== 10b. 类别内最终模型选择（ΔAICc<2 + MAPE>0） ====================
select_final_models_per_category <- function(model_comparison) {
  cat("\n=== 最终模型选择（8模型池，ΔAICc<2 + MAPE>0） ===\n")
  cat("  筛选规则: 每个疾病类别独立在8模型内应用三阶段选择\n")
  cat("  阶段1: 白噪声检验 p>0.05\n")
  cat("  阶段2: ΔAICc<2 且 MAPE>0 (排除零值完美拟合)\n")
  cat("  阶段3: MAPE最小\n\n")
  
  detailed_comparison <- model_comparison$detailed_comparison
  all_predictions <- model_comparison$all_predictions
  
  best_predictions <- list()
  selection_summary <- data.frame()
  
  # 按类别分组处理
  categories <- unique(detailed_comparison$Category)
  
  for (cat in categories) {
    cat(sprintf("\n[%s] 选择中:\n", cat))
    
    # 提取该类别的所有模型（最多8个）
    cat_models <- detailed_comparison %>% 
      filter(Category == cat) %>%
      arrange(AICc)  # 在类别内排序
    
    if (nrow(cat_models) == 0) {
      cat(sprintf("  警告: 无可用模型\n"))
      next
    }
    
    # 阶段1: 白噪声检验筛选
    white_noise_models <- cat_models %>% filter(Residual_LB_P > 0.05)
    
    if (nrow(white_noise_models) == 0) {
      cat(sprintf("  警告: 无模型通过白噪声检验，选择LB p值最大者\n"))
      white_noise_models <- cat_models %>% slice_max(order_by = Residual_LB_P, n = 1)
    } else {
      cat(sprintf("  白噪声通过: %d/%d 模型\n", nrow(white_noise_models), nrow(cat_models)))
    }
    
    # 阶段2: ΔAICc计算（类别内基准）
    best_aicc <- min(white_noise_models$AICc, na.rm = TRUE)
    white_noise_models <- white_noise_models %>%
      mutate(Delta_AICc = AICc - best_aicc)
    
    # ✅ 强化筛选：ΔAICc<2 且 MAPE>0
    strong_evidence_models <- white_noise_models %>% 
      filter(Delta_AICc < 2 & MAPE > 0)  # 排除MAPE=0的模型
    
    if (nrow(strong_evidence_models) == 0) {
      cat(sprintf("  警告: 无ΔAICc<2且MAPE>0模型，尝试仅ΔAICc<2\n"))
      strong_evidence_models <- white_noise_models %>% filter(Delta_AICc < 2)
      
      # 如果仍然没有，回退到ΔAICc最小
      if (nrow(strong_evidence_models) == 0) {
        cat(sprintf("  警告: 无ΔAICc<2模型，选择ΔAICc最小者\n"))
        selected_model <- white_noise_models %>% slice_min(order_by = Delta_AICc, n = 1)
      } else {
        # 在ΔAICc<2中选择MAPE最小且>0，如果都=0则选最小的非零
        valid_models <- strong_evidence_models %>% filter(MAPE > 0)
        if (nrow(valid_models) > 0) {
          selected_model <- valid_models %>% slice_min(order_by = MAPE, n = 1)
        } else {
          # 所有MAPE都=0，选择ΔAICc最小
          cat(sprintf("  警告: ΔAICc<2模型MAPE均=0，选择ΔAICc最小者\n"))
          selected_model <- strong_evidence_models %>% slice_min(order_by = Delta_AICc, n = 1)
        }
      }
    } else {
      # 阶段3: 在ΔAICc<2且MAPE>0子集中选择MAPE最小者
      selected_model <- strong_evidence_models %>% slice_min(order_by = MAPE, n = 1)
      cat(sprintf("  强证据模型: %d个 (ΔAICc<2且MAPE>0)\n", nrow(strong_evidence_models)))
    }
    
    # 提取完整预测对象
    selected_row_index <- selected_model$Row_Index[1]
    best_predictions[[cat]] <- all_predictions[[selected_row_index]]
    
    # 记录选择摘要
    summary_row <- data.frame(
      Category = cat,
      Selected_Model = selected_model$Model_Type[1],
      Time_Window = selected_model$Time_Window[1],
      Delta_AICc = selected_model$Delta_AICc[1],
      MAPE = selected_model$MAPE[1],
      LB_P_Value = selected_model$Residual_LB_P[1],
      Total_Models_Considered = nrow(cat_models),
      Models_Passed_White_Noise = nrow(white_noise_models),
      Strong_Evidence_Models = nrow(strong_evidence_models %>% filter(Delta_AICc < 2 & MAPE > 0)),
      stringsAsFactors = FALSE
    )
    selection_summary <- rbind(selection_summary, summary_row)
    
    cat(sprintf(" ✓ 最优: %s %s\n", selected_model$Model_Type[1], selected_model$Time_Window[1]))
    cat(sprintf("   ΔAICc=%.2f, MAPE=%.2f%%, LB p=%.3f\n",
                selected_model$Delta_AICc[1], selected_model$MAPE[1], selected_model$Residual_LB_P[1]))
  }
  
  cat(sprintf("\n✓ 选择完成: %d/%d 个单位成功\n", length(best_predictions), length(categories)))
  
  return(list(
    best_predictions = best_predictions,
    selection_summary = selection_summary
  ))
}

#==================== 11. 最终预测选择与导出 ====================
perform_final_forecast <- function(data_categories, forecast_results, final_selection) {
  cat("\n=== 最终预测导出（基于8模型选择结果） ===\n")
  
  export_dir <- file.path(output_dir, "导出结果")
  dir.create(export_dir, recursive = TRUE, showWarnings = FALSE)
  
  final_forecasts <- list()
  all_forecast_results <- data.frame()
  best_predictions <- final_selection$best_predictions
  
  for (category_name in names(data_categories)) {
    if (!category_name %in% names(best_predictions)) {
      cat(sprintf("  跳过: %s（无最终预测）\n", category_name))
      next
    }
    
    forecast_data <- best_predictions[[category_name]]
    model_info <- final_selection$selection_summary[final_selection$selection_summary$Category == category_name, ]
    
    cat(sprintf("导出: %s (%s)\n", category_name, model_info$Selected_Model))
    
    english_name <- disease_name_mapping[[category_name]]
    forecast_years_seq <- forecast_data$forecast_years
    
    # 预测值合理性检查
    forecast_values <- forecast_data$forecast_values
    if (any(forecast_values < 0)) {
      warning(sprintf("  警告: %s 预测出现负值，已截断为0", category_name))
      forecast_values <- pmax(forecast_values, 0)
    }
    
    # 年度变化率检查（GBD 2021标准：年变化>15%需复核）
    if (length(forecast_values) >= 2) {
      annual_change <- abs(diff(forecast_values) / forecast_values[-length(forecast_values)])
      if (any(annual_change > 0.15, na.rm = TRUE)) {
        warning(sprintf("  警告: %s 年变化率超过15%%，需流行病学专家复核", category_name))
      }
    }
    
    # 导出到数据框
    for (i in 1:length(forecast_years_seq)) {
      forecast_row <- data.frame(
        Category = category_name,
        Disease_Full_Name = english_name,
        Model_Type = model_info$Selected_Model,
        Time_Window = model_info$Time_Window,
        Year = forecast_years_seq[i],
        Forecast_Value = round(forecast_values[i], 2),
        Forecast_Lower_95CI = round(forecast_data$forecast_lower_95[i], 2),
        Forecast_Upper_95CI = round(forecast_data$forecast_upper_95[i], 2),
        Forecast_Lower_90CI = round(forecast_data$forecast_lower_90[i], 2),
        Forecast_Upper_90CI = round(forecast_data$forecast_upper_90[i], 2),
        Delta_AICc = round(model_info$Delta_AICc, 2),
        MAPE = round(model_info$MAPE, 2),
        Selection_Method = "三阶段：白噪声→ΔAICc<2+MAPE>0→MAPE最小（8模型池）",
        stringsAsFactors = FALSE
      )
      all_forecast_results <- rbind(all_forecast_results, forecast_row)
    }
    
    # 保存完整对象
    final_forecasts[[category_name]] <- list(
      Category = category_name,
      Disease_Full_Name = english_name,
      Model_Type = model_info$Selected_Model,
      Time_Window = model_info$Time_Window,
      Forecast_Years = forecast_years_seq,
      Forecast_Values = forecast_values,
      Forecast_Lower_95CI = forecast_data$forecast_lower_95,
      Forecast_Upper_95CI = forecast_data$forecast_upper_95,
      Forecast_Lower_90CI = forecast_data$forecast_lower_90,
      Forecast_Upper_90CI = forecast_data$forecast_upper_90,
      Historical_Years = forecast_data$historical_years,
      Historical_Values = forecast_data$historical_values,
      Fitted_Values = forecast_data$fitted_values,
      Model = forecast_data$model,
      Delta_Values = model_info$Delta_AICc,
      MAPE = model_info$MAPE
    )
  }
  
  cat(sprintf("完成: %d个单位最终预测导出\n", length(final_forecasts)))
  
  return(list(
    final_forecasts = final_forecasts,
    all_forecast_results = all_forecast_results
  ))
}

#==================== 12. 最终模型可视化（纯英文） ====================
create_final_visualizations <- function(data_categories, final_forecasts, output_dir, name_mapping) {
  cat("\n=== 创建最优模型可视化 ===\n")
  
  plots_dir <- file.path(output_dir, "Final_Models")
  dir.create(plots_dir, recursive = TRUE, showWarnings = FALSE)
  
  for (category_name in names(data_categories)) {
    if (!category_name %in% names(final_forecasts$final_forecasts)) {
      cat(sprintf("  跳过 %s（无最终预测结果）\n", category_name))
      next
    }
    
    english_name <- name_mapping[[category_name]]
    forecast_data <- final_forecasts$final_forecasts[[category_name]]
    
    file_name_base <- english_name %>%
      str_replace_all("[^a-zA-Z0-9, ]", "") %>%
      str_replace_all(", ", "_") %>%
      str_replace_all(" ", "_") %>%
      str_to_title() %>%
      gsub("[_]+", "_", .)
    
    combined_df <- data.frame(
      Year = c(forecast_data$Historical_Years, forecast_data$Forecast_Years),
      Value = c(forecast_data$Historical_Values, forecast_data$Forecast_Values),
      Type = c(rep("Historical", length(forecast_data$Historical_Years)),
               rep("Forecast", length(forecast_data$Forecast_Years)))
    )
    
    ci_df <- data.frame(
      Year = forecast_data$Forecast_Years,
      Lower_95 = forecast_data$Forecast_Lower_95CI,
      Upper_95 = forecast_data$Forecast_Upper_95CI,
      Lower_90 = forecast_data$Forecast_Lower_90CI,
      Upper_90 = forecast_data$Forecast_Upper_90CI
    )
    
    subtitle_text <- paste("Selected Model:", forecast_data$Model_Type, 
                           sprintf("| ΔAICc=%.2f, MAPE=%.2f%%", forecast_data$Delta_Values, forecast_data$MAPE))
    
    p <- ggplot() +
      geom_line(data = combined_df, aes(x = Year, y = Value, color = Type), linewidth = 1.2) +
      geom_point(data = filter(combined_df, Type == "Historical"), aes(x = Year, y = Value), size = 2.5, alpha = 0.8) +
      geom_ribbon(data = ci_df, aes(x = Year, ymin = Lower_95, ymax = Upper_95), fill = "#b20437", alpha = 0.3) +
      geom_ribbon(data = ci_df, aes(x = Year, ymin = Lower_90, ymax = Upper_90), fill = "#f4dfdb", alpha = 0.5) +
      geom_vline(xintercept = max(forecast_data$Historical_Years) + 0.5, linetype = "dotted", linewidth = 0.5) +
      labs(
        title = english_name,
        x = "Year",
        y = "Age-Standardized Rate (per 100,000)",
        subtitle = subtitle_text,
        caption = "Shaded: 90% and 95% Prediction Intervals | Points: Observed Data"
      ) +
      scale_color_manual(values = c(Historical = "#30a454", Forecast = "#b20437"), name = "Data Type") +
      theme_bw() +
      theme(
        text = element_text(family = safe_font),
        plot.background = element_rect(fill = "white", color = NA),
        panel.background = element_rect(fill = "white"),
        axis.title = element_text(face = "bold", size = 12),
        axis.text = element_text(size = 10),
        legend.position = "top",
        panel.grid = element_blank(),
        plot.margin = margin(10, 10, 10, 10)
      )
    
    suppressWarnings({
      ggsave(file.path(plots_dir, paste0(file_name_base, "_Final_Forecast.png")), p,
             width = 14, height = 8, dpi = 300)
    })
    
    cat(sprintf("  ✓ %s (ΔAICc=%.2f, MAPE=%.2f%%)\n", file_name_base, forecast_data$Delta_Values, 
                forecast_data$MAPE))
  }
  
  cat("完成: 最终模型图表已创建\n")
}

#==================== 13. MAPE性能热图（全模型显示） ====================
# ✅ 修复版：移除白噪声过滤，显示所有模型
create_mape_heatmap <- function(model_comparison, output_dir) {
  cat("\n=== 创建MAPE性能热图（显示所有模型，与CSV完全一致） ===\n")
  
  export_dir <- file.path(output_dir, "导出结果")
  detailed_comparison <- model_comparison$detailed_comparison
  
  # ✅ 修复点：不再过滤 Residual_LB_P，显示全部8模型
  heatmap_data <- detailed_comparison %>%
    mutate(Model_Label = paste0(Model_Type, "\n(", Time_Window, ")")) %>%
    select(Category, Model_Label, MAPE)  # 移除白噪声过滤
  
  model_labels <- unique(heatmap_data$Model_Label)
  disease_labels <- unique(heatmap_data$Category)
  
  heatmap_pivot <- expand.grid(Disease = disease_labels, Model = model_labels, stringsAsFactors = FALSE) %>%
    left_join(
      heatmap_data %>% rename(Disease = Category, Model = Model_Label),
      by = c("Disease", "Model")
    )
  
  heatmap_pivot$Disease <- factor(heatmap_pivot$Disease, levels = disease_labels)
  heatmap_pivot$Model <- factor(heatmap_pivot$Model, levels = model_labels)
  
  max_mape <- max(heatmap_pivot$MAPE, na.rm = TRUE)
  min_mape <- min(heatmap_pivot$MAPE, na.rm = TRUE)
  
  p <- ggplot(heatmap_pivot, aes(x = Model, y = Disease, fill = MAPE)) +
    geom_tile(color = "gray30", linewidth = 0.5) +
    geom_text(
      data = heatmap_pivot %>% filter(!is.na(MAPE)),
      aes(label = sprintf("%.2f%%", MAPE)),
      size = 4.5,
      color = "black",
      fontface = "bold",
      family = safe_font
    ) +
    scale_fill_gradientn(
      colours = c("#30a454", "#f4dfdb", "#b20437"),
      na.value = "white",  # NA显示为白色
      name = "MAPE (%)\n(Lower is Better)",
      limits = c(min_mape, min(20, max_mape))
    ) +
    labs(
      title = "Model Performance: MAPE Distribution (All Models)",
      subtitle = "Filter: None | All 8 Models Displayed (Consistent with CSV)",  # 明确标注
      x = "Model Type (Time Window)",
      y = "Disease Category",
      caption = "MAPE = Mean Absolute Percentage Error | Hyndman & Koehler (2006)"
    ) +
    theme_bw() +
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1, size = 12, face = "bold", family = safe_font),
      axis.text.y = element_text(size = 12, face = "bold", family = safe_font),
      axis.title = element_text(size = 14, face = "bold", family = safe_font),
      plot.title = element_text(size = 16, face = "bold", hjust = 0.5, family = safe_font),
      plot.subtitle = element_text(size = 13, hjust = 0.5, family = safe_font),
      plot.caption = element_text(size = 11, hjust = 0, family = safe_font),
      legend.position = "right",
      panel.grid = element_blank()
    )
  
  ggsave(
    file.path(export_dir, "MAPE_Performance_Heatmap_All_Models.png"),  # 重命名区分
    p,
    width = 18,
    height = 12,
    dpi = 300
  )
  
  cat(" ✓ MAPE热图已保存（全模型显示，与CSV数据一致）\n")
}

#==================== 14. 结果导出（学术期刊格式） ====================
export_all_results <- function(data_categories, forecast_results, final_selection, final_forecasts, model_comparison, output_dir) {
  cat("\n=== 结果导出（学术格式） ===\n")
  
  # ✅ 安全检查：验证model_comparison是否存在
  if (!exists("model_comparison") || is.null(model_comparison)) {
    stop("export_all_results: model_comparison对象为NULL或未传入")
  }
  
  export_dir <- file.path(output_dir, "导出结果")
  dir.create(export_dir, recursive = TRUE, showWarnings = FALSE)
  
  # 导出实际观测值
  actual_values_df <- data.frame()
  for (category_name in names(data_categories)) {
    category_data <- data_categories[[category_name]]
    if (nrow(category_data) > 0) {
      temp_df <- category_data %>%
        mutate(Category = category_name, Data_Type = "Actual") %>%
        select(Category, Data_Type, year, val) %>%
        rename(Year = year, Value = val)
      actual_values_df <- rbind(actual_values_df, temp_df)
    }
  }
  write.csv(actual_values_df, file.path(export_dir, "01_Actual_Values_1990-2023.csv"), row.names = FALSE)
  
  # 导出最终预测值
  write.csv(final_forecasts$all_forecast_results, 
            file.path(export_dir, "02_Final_Forecast_2024-2030.csv"), row.names = FALSE)
  
  # 导出选择摘要
  write.csv(final_selection$selection_summary, 
            file.path(export_dir, "03_Selection_Summary.csv"), row.names = FALSE)
  
  # 导出完整比较表（含MAPE）- 关键修复点
  write.csv(model_comparison$detailed_comparison, 
            file.path(export_dir, "04_Full_Comparison_with_MAPE.csv"), row.names = FALSE)
  
  # Data Availability Statement（更新方法学描述）
  das_text <- paste(
    "Data Availability Statement\n",
    "This study utilized publicly available data from the Global Burden of Disease Study 2021.",
    "The raw data file contains age-standardized cancer incidence and mortality rates for",
    "Palestine (1990-2023). Four forecasting methods were evaluated: Auto ARIMA, Manual ARIMA",
    "(p,d,q≤2), ETS (additive), and TBATS (non-seasonal, De Livera 2011). Model selection",
    "followed three stages: (1) Ljung-Box test, (2) Burnham & Anderson (2002) ΔAICc<2 + MAPE>0,",
    "(3) MAPE minimization (Hyndman & Koehler 2006). Analysis conducted in R", 
    R.version.string, "with packages: forecast (v", packageVersion("forecast"), "),", 
    "tseries (v", packageVersion("tseries"), ").",
    sep = "\n"
  )
  writeLines(das_text, file.path(export_dir, "Data_Availability_Statement.txt"))
  
  cat("完成: 结果已导出至", export_dir, "\n")
  cat("  新增: 方法学参考文献完整引用\n")
}

#==================== 15. 数据参与度验证函数 ====================
validate_data_participation <- function(data_categories, target_conditions) {
  cat("\n=== 数据参与度验证 ===\n")
  cat(sprintf("目标子集: %s\n", paste(names(target_conditions), target_conditions, sep="=", collapse=", ")))
  
  participation_report <- data.frame(
    Category = character(),
    Total_Records = integer(),
    Target_Records = integer(),
    Used_In_Fitting = logical(),
    Years_Present = character(),
    stringsAsFactors = FALSE
  )
  
  for (cat_name in names(data_categories)) {
    cat_data <- data_categories[[cat_name]]
    
    # 匹配目标条件
    matches_target <- Reduce(`&`, Map(`==`, cat_data[, names(target_conditions)], target_conditions))
    target_subset <- cat_data[matches_target, ]
    
    # 检查是否用于拟合（有实际值且无NA）
    used_in_fitting <- nrow(target_subset) > 0 && all(!is.na(target_subset$val))
    
    # 获取存在的年份
    years_present <- if (nrow(target_subset) > 0) {
      paste(range(target_subset$year), collapse="-")
    } else {
      "None"
    }
    
    report_row <- data.frame(
      Category = cat_name,
      Total_Records = nrow(cat_data),
      Target_Records = nrow(target_subset),
      Used_In_Fitting = used_in_fitting,
      Years_Present = years_present,
      stringsAsFactors = FALSE
    )
    
    participation_report <- rbind(participation_report, report_row)
    
    if (nrow(target_subset) > 0) {
      cat(sprintf(" ✓ [%s] 目标子集: %d条记录 (%s) → %s\n", 
                  cat_name, nrow(target_subset), years_present,
                  ifelse(used_in_fitting, "已参与拟合", "未使用")))
    }
  }
  
  # 汇总
  total_target <- sum(participation_report$Target_Records)
  total_used <- sum(participation_report$Used_In_Fitting)
  
  cat(sprintf("\n=== 验证结果 ===\n"))
  cat(sprintf("目标子集总记录数: %d\n", total_target))
  if (total_target > 0) {
    cat(sprintf("参与模型拟合数: %d (%.1f%%)\n", total_used, 100*total_used/total_target))
    
    if (total_used == 0) {
      cat(" ❌ 警告: 目标子集未参与任何模型拟合！\n")
    } else {
      cat(" ✅ 确认: 目标子集已参与模型训练\n")
    }
  } else {
    cat(" ❌ 警告: 未找到匹配的目标子集数据！\n")
  }
  
  return(participation_report)
}

#==================== 16. 主运行函数====================
run_hybrid_forecast_system <- function() {
  cat("\n", paste0(strrep("=", 70), collapse = ""), "\n")
  cat("ARIMA-ETS-TBATS Forecast System (8-Model Selection, v7.2)\n")
  cat(paste0(strrep("=", 70), collapse = ""), "\n")
  cat("方法论框架 (Burnham & Anderson 2002):\n")
  cat("  1. 数据准备: 12类别显式构建，零值保留（避免MAPE失真）\n")
  cat("  2. 双窗口预测: 8独立模型/类别（4算法×2窗口）\n")
  cat("  3. 探索性分析: MAPE热图（全模型显示，无选择标准）\n")  # ✅ 更新描述
  cat("  4. 最终选择: ΔAICc<2 + MAPE>0 + MAPE最小（严格双重约束）\n")
  cat("  5. 数据验证: 追踪指定子集参与度\n\n")
  
  start_time <- Sys.time()
  
  # 阶段1: 数据准备
  cat("=== 阶段1: 数据准备 ===\n")
  data_prep <- prepare_epidemic_data(input_file, 1990, 2023, "主窗口")
  
  # 阶段1a: 数据参与度验证
  cat("\n=== 阶段1a: 数据参与度验证 ===\n")
  target_conditions <- list(
    measure_name = "发病率",
    sex_name = "女",
    cause_name = "其他非恶性肿瘤"
  )
  participation_report <- validate_data_participation(data_prep$categories, target_conditions)
  
  # 阶段2: 双窗口预测执行
  cat("\n=== 阶段2: 双窗口预测执行 ===\n")
  forecast_results <- perform_dual_window_predictions(data_prep$categories, forecast_years = 7)
  
  # 阶段3: 完整模型评估矩阵
  cat("\n=== 阶段3: 完整模型评估矩阵 ===\n")
  model_comparison <- compare_and_generate_matrix(forecast_results, data_prep$categories)
  
  # 阶段3a: MAPE热图（探索性分析）- 与CSV完全一致
  cat("\n=== 阶段3a: MAPE热图（全模型，无过滤） ===\n")
  create_mape_heatmap(model_comparison, output_dir)
  
  # 阶段3b: 最终模型选择（8模型池，强化标准）
  cat("\n=== 阶段3b: 最终模型选择（ΔAICc<2 + MAPE>0 + MAPE最小） ===\n")
  final_selection <- select_final_models_per_category(model_comparison)
  
  # 阶段4: 最终预测聚合
  cat("\n=== 阶段4: 最终预测聚合 ===\n")
  final_forecasts <- perform_final_forecast(data_prep$categories, forecast_results, final_selection)
  
  # 阶段5: 可视化生成
  cat("\n=== 阶段5: 可视化生成 ===\n")
  create_final_visualizations(data_prep$categories, final_forecasts, output_dir, disease_name_mapping)
  
  # 阶段6: 结果导出
  cat("\n=== 阶段6: 结果导出 ===\n")
  export_all_results(data_prep$categories, forecast_results, final_selection,
                     final_forecasts, model_comparison, output_dir)
  
  # 导出参与度报告
  write.csv(participation_report, 
            file.path(output_dir, "导出结果", "Data_Participation_Report.csv"), 
            row.names = FALSE)
  
  # 完成报告
  cat("\n", paste0(strrep("=", 70), collapse = ""), "\n")
  cat("✅ 系统运行完成! 所有模块执行成功\n")
  cat(sprintf("✅ 成功预测单位: %d/%d\n", length(final_forecasts$final_forecasts), length(data_prep$categories)))
  cat(sprintf("✅ 数据验证完成: %s\n", ifelse(sum(participation_report$Used_In_Fitting) > 0, "目标子集已参与", "目标子集未参与")))
  
  cat("\n📊 输出文件目录:\n")
  cat(sprintf("   %s/\n", output_dir))
  cat("   ├── forecast_log.txt\n")
  cat("   ├── Final_Models/\n")
  cat("   └── 导出结果/\n")
  cat("       ├── 01_Actual_Values_1990-2023.csv\n")
  cat("       ├── 02_Final_Forecast_2024-2030.csv\n")
  cat("       ├── 03_Selection_Summary.csv\n")
  cat("       ├── 04_Full_Comparison_with_MAPE.csv\n")
  cat("       ├── MAPE_Performance_Heatmap_All_Models.png\n")  # ✅ 重命名
  cat("       ├── Data_Availability_Statement.txt\n")
  cat("       └── Data_Participation_Report.csv\n")
  
  # 方法学局限性声明
  cat("\n=== 方法学局限性声明 ===\n")
  cat("1. 趋势外推法假设未来模式与历史一致\n")
  cat("2. TBATS模型假设无季节性\n")
  cat("3. 零值保留：可能增加MAPE计算变异度\n")
  cat("4. 严格选择标准可能减少可用模型数\n")
  cat("5. 适用于短期规划（5-10年），长期预测需政策情景分析\n")
  
  # 计算环境报告
  cat("\n=== 计算环境报告 ===\n")
  cat(sprintf("R版本: %s\n", R.version.string))
  cat(sprintf("操作系统: %s\n", R.version$platform))
  cat(sprintf("forecast包版本: %s\n", packageVersion("forecast")))
  cat(sprintf("tseries包版本: %s\n", packageVersion("tseries")))
  cat(sprintf("运行时间: %.1f秒\n", as.numeric(Sys.time() - start_time)))
  cat(sprintf("峰值内存: %.2f MB\n", as.numeric(gc()[2, 2]) / 1024^2))
  
  cat(paste0(strrep("=", 70), collapse = ""), "\n")
  
  sink(NULL)
  
  return(list(
    participation_report = participation_report,
    forecast_results = forecast_results,
    model_comparison = model_comparison,
    final_selection = final_selection,
    final_forecasts = final_forecasts,
    output_dir = output_dir,
    log_file = log_file
  ))
}

#==================== 17. 执行主程序 ====================
cat("正在初始化预测系统...\n")
cat(sprintf("输出目录: %s\n", output_dir))

# 执行完整流程
results <- run_hybrid_forecast_system()

cat("\n系统退出\n")