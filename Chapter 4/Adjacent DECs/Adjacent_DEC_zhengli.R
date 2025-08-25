##将相邻阶段差异表达的circRNA整合到一个excel的不同sheet中
setwd("E:/工作电脑D盘/科研文档/UCD/Thesis/Qiaowei_PhD_Thesis_Supporting/Chapter 4/Adjacent DECs")
# 加载必要的包
library(openxlsx)

# 读取当前文件夹下的所有txt文件
files <- list.files(pattern = "\\.txt$")

# 创建一个Excel工作簿
wb <- createWorkbook()

for (file in files) {
  # 读取数据
  df <- read.table(file, header = TRUE, sep = "\t", stringsAsFactors = FALSE)
  
  # 删除第2、3列和最后三列
  drop_cols <- c(2, 3, (ncol(df)-2):ncol(df))  # 指定要删除的列
  df <- df[, -drop_cols, drop = FALSE]
  
  # 替换 circUBE3A_4 为 circUBE3A
  df[] <- lapply(df, function(x) {
    if (is.character(x)) {
      gsub("circUBE3A_4", "circUBE3A", x)
    } else {
      x
    }
  })
  
  # 去掉文件名前缀 20250407 和 "_decirc"
  sheet_name <- tools::file_path_sans_ext(file)
  sheet_name <- gsub("^20250407", "", sheet_name)   # 去掉开头的20250407
  sheet_name <- gsub("_decirc", "", sheet_name)     # 去掉_decirc
  
  # 避免重复sheet名：如果重复则加上编号
  if (sheet_name %in% names(wb$worksheets)) {
    i <- 1
    new_name <- paste0(sheet_name, "_", i)
    while (new_name %in% names(wb$worksheets)) {
      i <- i + 1
      new_name <- paste0(sheet_name, "_", i)
    }
    sheet_name <- new_name
  }
  
  # 添加sheet并写入数据
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df)
}

# 保存Excel文件
saveWorkbook(wb, "All_results.xlsx", overwrite = TRUE)
