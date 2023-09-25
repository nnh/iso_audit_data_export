# 必要なライブラリを読み込む
library(tidyverse)

# フォルダ内のファイル名を取得する関数
GetFilenames <- function(folder_path) {
  list.files(folder_path, full.names = FALSE)
}

# フォルダAとフォルダBの比較を行う関数
CompareFolders <- function(folder_path_A, folder_path_B) {
  # フォルダAとフォルダBのパスを表示
  cat("フォルダAのパス:", folder_path_A, "\n")
  cat("フォルダBのパス:", folder_path_B, "\n")

  # フォルダAとフォルダBからファイル名を取得
  file_list_A <- GetFilenames(folder_path_A)
  file_list_B <- GetFilenames(folder_path_B)

  target_pattern <- ".txt$"
  matching_files_A <- file_list_A[str_detect(file_list_A, target_pattern)]
  matching_files_B <- file_list_B[str_detect(file_list_B, target_pattern)]
  matching_files_B <- gsub("デ", "デ", matching_files_B)
  matching_files_B <- gsub("ジ", "ジ", matching_files_B)
  matching_files_B <- gsub("ビ", "ビ", matching_files_B)
  matching_files_B <- gsub("グ", "グ", matching_files_B)
  matching_files_B <- gsub("バ", "バ", matching_files_B)
  matching_files_B <- gsub("ISF08 マネジメントレビュー", "ISF08 マネジメントレビュー議事録", matching_files_B)
  matching_files_B <- gsub("ISF13 守秘義務誓約書 ", "ISF13 守秘義務誓約書　 ", matching_files_B)
  matching_files_B <- gsub("ISF14 秘密保持契約書", "ISF14 秘密保持覚書", matching_files_B)
  matching_files_B <- gsub("ISF29 入退室管理台帳 第三者用", "ISF29 入退室管理台帳 第三者用（セキュリティーカード対応に移行）", matching_files_B)
  matching_files_B <- gsub("ISF17 協力先評価表（情報システム研究室）", "ISF17 協力先評価表（情報システム研究室） ", matching_files_B)

  # フォルダAにのみ存在するファイル名を抽出
  unique_to_A <- matching_files_A[!matching_files_A %in% matching_files_B]
  # QF7-QF12, QF15-AF20は2022年度はBoxになくてもOK
  if (parent[2] == "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/04 記録/2022年度/固定/"){
    unique_to_A <- unique_to_A[!str_detect(unique_to_A, "^QF(0|10|11|12)")]
    unique_to_A <- unique_to_A[!str_detect(unique_to_A, "^QF(15|16|17|18|19|20)")]
  }
  # フォルダBにのみ存在するファイル名を抽出
  unique_to_B <- matching_files_B[!matching_files_B %in% matching_files_A]
  if (parent[2] == "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/04 記録/2022年度/固定/"){
    unique_to_B <- unique_to_B[!str_detect(unique_to_B, "QF07-QF12 データ管理室のみ.txt")]
    unique_to_B <- unique_to_B[!str_detect(unique_to_B, "QF15-QF20 データ管理室のみ.txt")]
  }

  # 結果をデータフレームに格納
  result <- list(
    UniqueToA = unique_to_A,
    UniqueToB = unique_to_B
  )

  # 結果を表示
  cat("\nフォルダAにのみ存在するファイル名:\n")
  print(result$UniqueToA)

  cat("\nフォルダBにのみ存在するファイル名:\n")
  print(result$UniqueToB)

  return(result)
}

target <- c("ISMS（情報システム研究室）", "QMS（情報システム研究室）")
parent <- c("~/Downloads/固定/",
            "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/04 記録/2022年度/固定/")
target_folders <- list()
# ベクトルAとベクトルBの組み合わせを作成し、結合してリストに格納
for (b in target) {
  combined_values <- str_c(parent, b)
  target_folders <- append(target_folders, list(combined_values))
}
comparison_result <- list()
i <- 1
for (folder in target_folders) {
  comparison_result[[i]] <- CompareFolders(folder[1], folder[2])
  i <- i + 1
}
