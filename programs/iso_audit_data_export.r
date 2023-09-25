#' Process to output fixed files required for ISO audits.
#' This script generates the required fixed files for ISO audits.
#' @file iso_audit_data_export.r
#' @author Mariko Ohtsuka
#' @date 2023.09.25

# Clear temporary variables
rm(list=ls())
# ------ libraries ------
library(tidyverse)
library(readxl)
library(lubridate)
# ------ constants ------
kInputPath <- "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/02 文書（ドラフト）/D000 QMS・ISMS文書一覧 230922.xlsx"
fiscal_year <- CalculateTodayFiscalYear()
kOutputParentPath <- str_c("~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/04 記録/", fiscal_year, "年度/")
kQmfHeader <- "QF"
kIsmsHeader <- "ISF"
kIsms <- "ISMS"
kQms <- "QMS"
kRefIsms <- str_c(kIsms, "参照")
kRefQms <- str_c(kQms, "参照")
kRefCommon <- "共通"
kTargetSign <- "○"
kIsrColName <- "...10"
kDcColName <- "保管部門"
kCategory <- "区分"
kFormat <- "形式"
kItemName <- "記録名"
kIsrName <- "情報システム研究室"
kDcName <- "データ管理室"
kPaper <- "紙"
# Lock constants
ls(pattern="^k") %>% map( ~ lockBinding(., .GlobalEnv))
# ------ functions ------
FilterTargetDs <- function(ds){
  res1 <- ds %>% filter(is.na(.[[kIsrColName]]) | str_detect(.[[kIsrColName]], kRefIsms) | str_detect(.[[kIsrColName]], kRefQms)) %>%
    filter(str_detect(.[[kDcColName]], kTargetSign) | !is.na(.[[kIsrColName]]))
  anti_res1 <- ds %>% anti_join(res1, kCategory)
  res2 <- anti_res1 %>% filter(str_detect(.[[kFormat]], kPaper))
  res <- bind_rows(res1, res2) %>% distinct() %>%
    select(c(all_of(kCategory), all_of(kItemName), all_of(kFormat), all_of(kDcColName), all_of(kIsrColName)))
  pattern <- str_c("(", kQmfHeader, "|", kIsmsHeader, ")\\d+|", "^", kPaper, "($|\r\n)")
  res[[kFormat]] <- res[[kFormat]] %>% str_extract(pattern)
  return(res)
}
FilterByFormat <- function(input_ds, targetColName){
  kCategoryX <- str_c(kCategory, ".x")
  kCategoryY <- str_c(kCategory, ".y")
  ds <- input_ds %>% filter(str_detect(.[[targetColName]], kRefIsms) | str_detect(.[[targetColName]], kRefQms))
  na_values <- ds %>% filter(is.na(.[[kFormat]])) %>% left_join(all_data, by=kItemName) %>%
    filter(.[[kCategoryX]] != .[[kCategoryY]])
  na_values[[kCategory]] <- na_values[[kCategoryX]]
  na_values[[kFormat]] <- na_values[[kCategoryY]]
  na_values <- na_values %>% select(c(all_of(kCategory), all_of(kItemName), all_of(kFormat)))
  non_na_values <- ds %>% filter(!is.na(.[[kFormat]])) %>% select(c(all_of(kCategory), all_of(kItemName), all_of(kFormat)))
  res <- bind_rows(na_values, non_na_values)
  return (res)
}
GenerateFilenames <- function(ds, dept){
  ds$format <- ifelse(str_detect(ds[[kFormat]], kPaper), "紙保管", str_c(ds[[kFormat]], "参照"))
  res <- ds %>%  rowwise() %>% mutate(filename=str_c(get(kCategory), " ", get(kItemName), " ", dept, format, ".txt"))
  return (res)
}
GenerateFilteredOtherData <- function(input_ds, dept, filter_condition) {
  output_ds <- input_ds %>% filter(str_detect(.[[kCategory]], filter_condition))
  output_ds[[kFormat]] <- ifelse(is.na(output_ds[[kFormat]]), output_ds[[kCategory]], output_ds[[kFormat]])
  output_ds <- GenerateFilenames(output_ds, dept)
  output_ds <- output_ds %>% select(-all_of(kIsrColName), -all_of(kDcColName))
  return(output_ds)
}
WriteTextFiles <- function(filename){
  output_path <- GenerateOutputFilepath(filename)
  if (!is.na(output_path)){
    file.create(output_path)
  }
}
GenerateOutputFilepath <- function(filename) {
  if (str_detect(filename, str_c("^", kQmfHeader))) {
    output_path <- str_c(kOutputParentPath, "固定/QMS（", kIsrName, "）/")
  } else if  (str_detect(filename, str_c("^", kIsmsHeader))) {
    output_path <- str_c(kOutputParentPath, "固定/ISMS（", kIsrName, "）/")
  } else {
    print(str_c("Error_filename:", filename))
    output_path <- NA
  }
  output_filename <- str_remove_all(filename, "\r|\n")
  output_file <- ifelse(!is.na(output_path), str_c(output_path, output_filename), NA)
  return(output_file)
}

WriteVectorToFile <- function(vector, filename) {
  output_path <- GenerateOutputFilepath(filename)
  if (!is.na(output_path)){
    cat(file=output_path, vector)
  }
}
SplitDatasetByCategory <- function(input_dataset, category){
  matching_category <- input_dataset %>% filter(.[[kCategory]] == category)
  non_matching_category <- input_dataset %>% filter(.[[kCategory]] != category)
  return(list(matching_category, non_matching_category))
}
CalculateTodayFiscalYear <- function() {
  today <- Sys.Date()
  fiscal_year_start <- as.Date(paste0(year(today), "-04-01"))
  if (today >= fiscal_year_start) {
    year_value <- year(today)
  } else {
    year_value <- year(today) - 1
  }
  return(year_value)
}
# ------ main ------
# Input file path
input_path <- kInputPath

# Read all sheets from the Excel file
all_sheets <- excel_sheets(input_path)
sheet_data <- all_sheets %>% map( ~ read_excel(input_path, sheet=., skip=4))
raw_qms <- sheet_data[[2]]  # 文書管理台帳(2)
raw_isms <- sheet_data[[3]]  # 文書管理台帳(3)
all_data <- bind_rows(raw_isms, raw_qms)

temp <- SplitDatasetByCategory(all_data, "ISF19")
isf19 <- temp[[1]]
temp <- SplitDatasetByCategory(all_data, "ISF22")
isf22 <- temp[[1]]
isf22_title_and_path <- list(c("DC入退室", "¥¥aronas¥Archives¥Log¥DC入退室"),
                             c("PivotalTracker", "¥¥aronas¥Archives¥PivotalTracker"),
                             c("UTM", "¥¥aronas¥Archives¥ISR¥SystemAssistant¥monthlyOperations"),
                             c("VPN", "¥¥aronas¥Archives¥Log¥VPN")
                            )
temp <- SplitDatasetByCategory(all_data, "ISF29")
isf29 <- temp[[1]]
# Filter target data
target_ds <- all_data %>% FilterTargetDs()

# Filter data by format
temp <- SplitDatasetByCategory(target_ds, "QF30")
qf30 <- temp[[1]]
target_ds <- temp[[2]]
isr_ref_ds <- target_ds %>% FilterByFormat(kIsrColName)
dc_ref_ds <- target_ds %>% FilterByFormat(kDcColName)
other_ds <- target_ds %>% anti_join(isr_ref_ds, by=kCategory) %>% anti_join(dc_ref_ds, by=kCategory)

# Generate filenames and export data
isr <- GenerateFilenames(isr_ref_ds, kIsrName)
dc <- GenerateFilenames(dc_ref_ds, kDcName)
dc <- dc %>% anti_join(isr, by=kCategory)
other_qms <- GenerateFilteredOtherData(other_ds, kDcName, kQmfHeader)
other_isms <- GenerateFilteredOtherData(other_ds, kDcName, kIsmsHeader)

# Combine all data
output_ds_all <- bind_rows(isr, dc, other_qms, other_isms) %>% arrange(get(kCategory))

# Output files
dummy <- output_ds_all$filename %>% map( ~ WriteTextFiles(.) )
dummy <- WriteVectorToFile(str_c("Box¥Projects¥ISO¥QMS・ISMS文書¥06 その他¥研修資料¥", fiscal_year, "年度"),
                           str_c(qf30[ , kCategory, drop=T], " ", qf30[ , kItemName, drop=T], ".txt"))
dummy <- isf22_title_and_path %>% map( ~ {
  vector <- .[2]
  filename <- str_c(isf22[ , kCategory, drop=T], " ", isf22[ , kItemName, drop=T], " ", .[1], ".txt")
  dummy <- WriteVectorToFile(vector, filename)
})
dummy <- WriteVectorToFile(str_c("Box¥Datacenter¥ISR¥Ptosh¥Ptosh Validation"),
                           str_c(isf19[ , kCategory, drop=T], " ", isf19[ , kItemName, drop=T], ".txt"))
dummy <- WriteVectorToFile(str_c("¥¥aronas¥Archives¥ISR¥SystemAssistant¥サーバ室作業報告書"),
                           "ISF29 サーバ室作業報告書.txt")
