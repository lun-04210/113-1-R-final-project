# 引入套件
library(readr)
library(dplyr)
library(ggplot2)
library(openxlsx)

# 讀取資料
foreign_data <- read_csv("foreign currencies.csv")

# 數據處理
foreign_data <- foreign_data %>%
  mutate(
    年度 = as.numeric(年度),
    幣別 = as.character(幣別),
    出入境 = as.character(出入境),
    數量 = as.numeric(數量)
  ) %>%
  # 將「人境」視為「入境」
  mutate(出入境 = ifelse(出入境 == "人境", "入境", 出入境))

# 合併同一年、同幣別的入境和出境數據
foreign_data <- foreign_data %>%
  group_by(幣別, 年度, 出入境) %>%
  summarise(數量 = sum(數量, na.rm = TRUE), .groups = "drop")

# 按幣別和出入境計算總數量
currency_data <- foreign_data %>%
  group_by(幣別, 出入境) %>%
  summarise(數量 = sum(數量, na.rm = TRUE), .groups = "drop")

# 按幣別和年度計算出入境數量
trend_data <- foreign_data

# 轉換為寬格式，便於比較出境與入境
trend_data_wide <- trend_data %>%
  pivot_wider(names_from = 出入境, values_from = 數量, values_fill = 0)

# 計算幣別總額排名
ranking_data <- currency_data %>%
  group_by(幣別) %>%
  summarise(總金額 = sum(數量, na.rm = TRUE)) %>%
  arrange(desc(總金額))

# 創建 Excel 工作簿
wb <- createWorkbook()

# 添加排名表
addWorksheet(wb, "幣別金額排名")
writeData(wb, sheet = "幣別金額排名", ranking_data)

# 獲取所有幣別
currencies <- unique(currency_data$幣別)

# 為每種幣別生成表格與圖表
for (currency in currencies) {
  # 篩選該幣別的數據
  bar_data <- currency_data %>%
    filter(幣別 == currency)
  
  line_data <- trend_data_wide %>%
    filter(幣別 == currency)
  
  # 添加工作表
  addWorksheet(wb, sheetName = currency)
  
  # 寫入數據表到工作表
  writeData(wb, sheet = currency, bar_data, startRow = 1, startCol = 1)
  writeData(wb, sheet = currency, line_data, startRow = nrow(bar_data) + 4, startCol = 1)
  
  # 計算圖片插入行位置，避免遮擋表格數據
  bar_image_start_row <- nrow(bar_data) + nrow(line_data) + 6
  line_image_start_row <- bar_image_start_row + 20  # 為圖片留出空間
  
  # 繪製長條圖（白底設置）
  bar_plot <- ggplot(bar_data, aes(x = 出入境, y = 數量, fill = 出入境)) +
    geom_bar(stat = "identity", show.legend = FALSE) +
    labs(
      title = paste(currency, "出入境數量"),
      x = "出入境",
      y = "數量"
    ) +
    theme_minimal() +
    theme(
      panel.background = element_rect(fill = "white", color = NA),
      plot.background = element_rect(fill = "white", color = NA),
      panel.grid.major = element_line(color = "gray90"),
      panel.grid.minor = element_blank()
    )
  
  # 保存長條圖為臨時文件
  bar_path <- tempfile(fileext = ".png")
  ggsave(filename = bar_path, plot = bar_plot, width = 6, height = 4)
  
  # 插入長條圖到 Excel
  insertImage(wb, sheet = currency, file = bar_path, startRow = bar_image_start_row, startCol = 1, width = 6, height = 4)
  
  # 繪製趨勢折線圖（白底設置）
  line_plot <- ggplot(line_data, aes(x = 年度)) +
    geom_line(aes(y = 入境, color = "入境"), size = 1) +
    geom_line(aes(y = 出境, color = "出境"), size = 1) +
    geom_point(aes(y = 入境, color = "入境"), size = 2) +
    geom_point(aes(y = 出境, color = "出境"), size = 2) +
    scale_color_manual(values = c("入境" = "blue", "出境" = "red")) +
    labs(
      title = paste(currency, "年度出入境趨勢"),
      x = "年度",
      y = "數量",
      color = "類型"
    ) +
    theme_minimal() +
    theme(
      panel.background = element_rect(fill = "white", color = NA),
      plot.background = element_rect(fill = "white", color = NA),
      panel.grid.major = element_line(color = "gray90"),
      panel.grid.minor = element_blank()
    ) +
    scale_x_continuous(breaks = seq(min(line_data$年度, na.rm = TRUE), max(line_data$年度, na.rm = TRUE), by = 1)) +
    scale_y_continuous(labels = scales::comma)
  
  # 保存折線圖為臨時文件
  line_path <- tempfile(fileext = ".png")
  ggsave(filename = line_path, plot = line_plot, width = 6, height = 4)
  
  # 插入折線圖到 Excel
  insertImage(wb, sheet = currency, file = line_path, startRow = line_image_start_row, startCol = 1, width = 6, height = 4)
}

# 保存 Excel 文件
saveWorkbook(wb, "旅客入出境幣別數據與趨勢分析_帶排名.xlsx", overwrite = TRUE)



