##Table S10
setwd("E:/工作电脑D盘/科研文档/UCD/Thesis/Qiaowei_PhD_Thesis_Supporting/Chapter 3/Table S3.10")

library(openxlsx)

files <- list.files(pattern = "\\.txt$")
wb <- createWorkbook()

for (f in files) {
  # 读取
  df <- read.table(f, header = TRUE, sep = "\t", stringsAsFactors = FALSE)
  
  # circ_name 只保留第一个下划线之后
  df$circ_name <- sub("^[^_]*_", "", df$circ_name)
  
  # 保留前 8 列
  if (ncol(df) > 8) df <- df[, 1:8]
  
  # 对所有字符串做替换，把 sus_circME1_14 换成 circME1
  # df[] <- lapply(df, function(x) {
  #   if (is.character(x)) {
  #     gsub("circME1_14", "circME1", x)
  #   } else {
  #     x
  #   }
  # })
  # 替换规则
  replace_map <- c(
    "circME1_14"   = "circME1",
    "circOXR1_16"  = "circOXR1",
    "circEHBP1_10" = "circEHBP1"
  )
  
  df[] <- lapply(df, function(x) {
    if (is.character(x)) {
      for (pattern in names(replace_map)) {
        x <- gsub(pattern, replace_map[pattern], x)
      }
    }
    x
  })
  
  
  # 生成简短 sheet 名：去掉 DEC_Adipose_
  sheet_name <- sub("^DEC_Adipose_", "", f)
  sheet_name <- sub("\\.txt$", "", sheet_name)
  sheet_name <- substr(sheet_name, 1, 31)  # 保证 <=31 个字符
  
  # 添加 sheet 并写入数据
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df)
}

# 保存 Excel
saveWorkbook(wb, "All_circRNAs_processed.xlsx", overwrite = TRUE)
