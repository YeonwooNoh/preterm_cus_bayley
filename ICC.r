library(readxl)
library(dplyr)
library(irr)
library(writexl)

# ==============================================================================
# [중요] 파일 경로 및 시트 이름 설정
# ==============================================================================
file_path <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/data/data_cus_final.xlsx"
sheet_name_J <- "초음파지표_J"
sheet_name_H <- "초음파지표_H"
# ==============================================================================

# 1. 데이터 불러오기
tryCatch({
  data_J <- read_excel(file_path, sheet = sheet_name_J, skip = 1)
  data_H <- read_excel(file_path, sheet = sheet_name_H, skip = 1)
  print("데이터 로드 성공!")
}, error = function(e) {
  stop("파일을 읽을 수 없습니다. 경로와 시트명을 확인해주세요.\n", e$message)
})

# 2. 전처리
data_J <- data_J %>% filter(!is.na(No))
data_H <- data_H %>% filter(!is.na(No))

common_ids <- intersect(data_J$No, data_H$No)
if(length(common_ids) == 0) stop("일치하는 환자 번호(No)가 없습니다.")

data_J <- data_J %>% filter(No %in% common_ids) %>% arrange(No)
data_H <- data_H %>% filter(No %in% common_ids) %>% arrange(No)

cat(sprintf("분석 대상 환자 수: %d명\n", length(common_ids)))

# 3. 변수 선정
common_cols <- intersect(names(data_J), names(data_H))
exclude_cols <- c("No", "Composite Category", "Scoring system", "개별 지표") 
target_cols <- setdiff(common_cols, exclude_cols)

numeric_cols <- target_cols[sapply(data_J[target_cols], is.numeric)]
numeric_cols <- numeric_cols[!grepl("^\\.\\.\\.", numeric_cols)]

# 결과 저장할 데이터프레임 (Outlier_No 컬럼 추가됨)
icc_results <- data.frame(
  Variable = character(),
  ICC_Value = numeric(),
  F_Value = numeric(),
  p_Value = numeric(),
  Lower_CI = numeric(),
  Upper_CI = numeric(),
  Outlier_No = character(), # 문제가 되는 환자 번호를 담을 열
  stringsAsFactors = FALSE
)

print(paste("총", length(numeric_cols), "개 변수에 대해 분석을 시작합니다..."))

# 4. 반복문 실행
for (var in numeric_cols) {
  
  # 데이터와 환자 번호를 함께 추출 (NA 제거 시 No도 같이 제거하기 위해)
  df_temp <- data.frame(
    No = data_J$No,
    J = data_J[[var]],
    H = data_H[[var]],
    stringsAsFactors = FALSE
  )
  
  # 결측치 제거
  df_temp <- na.omit(df_temp)
  
  if(nrow(df_temp) < 2) next 
  
  # ICC 계산용 데이터
  ratings <- df_temp[, c("J", "H")]
  
  tryCatch({
    # ICC(3,2) Consistency
    icc_res <- icc(ratings, model = "twoway", type = "consistency", unit = "average")
    icc_val <- round(icc_res$value, 3)
    
    # -----------------------------------------------------------
    # [추가 기능] ICC < 0.75 일 때 원인이 되는 Outlier 찾기
    # -----------------------------------------------------------
    outlier_str <- "" # 기본값은 빈 문자열
    
    if (!is.na(icc_val) && icc_val < 0.75) {
      # Consistency는 '경향성'이므로, 각각 표준화(Z-score)하여 비교합니다.
      # 즉, J가 평균보다 높을 때 H도 평균보다 높아야 하는데, 반대인 경우를 찾습니다.
      z_J <- scale(df_temp$J)
      z_H <- scale(df_temp$H)
      
      # 표준화된 점수의 차이 (절대값) 계산
      # 이 값이 클수록 경향성에서 벗어난 데이터입니다.
      diffs <- abs(z_J - z_H)
      
      # 차이가 큰 순서대로 정렬하여 상위 5개 인덱스 추출
      # top_indices <- order(diffs, decreasing = TRUE)[1:5]
      top_indices <- order(diffs, decreasing = TRUE)[1:min(10, length(diffs))]
      
      # 해당 인덱스가 유효한 범위인지 확인 후 No 추출
      top_indices <- top_indices[!is.na(top_indices)]
      
      if (length(top_indices) > 0) {
        bad_nos <- df_temp$No[top_indices]
        # 엑셀 셀에 "No: 101, 105, 203" 형태로 저장
        outlier_str <- paste(bad_nos, collapse = ", ")
      }
    }
    # -----------------------------------------------------------
    
    icc_results <- rbind(icc_results, data.frame(
      Variable = var,
      ICC_Value = icc_val,
      F_Value = round(icc_res$Fvalue, 3),
      p_Value = round(icc_res$p.value, 4),
      Lower_CI = round(icc_res$lbound, 3),
      Upper_CI = round(icc_res$ubound, 3),
      Outlier_No = outlier_str # 이상치 환자 번호 저장
    ))
    
  }, error = function(e) {
    message(paste("변수명 [", var, "] 에러: ", e$message))
  })
}

# 5. 결과 저장
print(head(icc_results))

save_folder <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/"
output_filename <- paste0(save_folder, "ICC_Results_With_Outliers.xlsx")

write_xlsx(icc_results, output_filename)

cat("\n=======================================================\n")
cat(" 분석 완료! \n")
cat(" ICC가 0.6 미만인 경우 'Outlier_No' 열에 \n")
cat(" 불일치가 가장 큰 환자 번호(Top 5)가 표시됩니다.\n")
cat(" 파일 저장 위치: ", file.path(save_folder, "ICC_Results_With_Outliers.xlsx"), "\n")
cat("=======================================================\n")