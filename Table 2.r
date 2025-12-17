library(readxl)
library(dplyr)
library(writexl)
library(stringr)

# 1. 데이터 불러오기 (엑셀 파일의 각 시트 지정)
# 파일명이 'data_final.xlsx'라고 가정합니다. 경로가 다르다면 수정해주세요.
file_path <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/data/data_final.xlsx"

df_demo <- read_excel(file_path, sheet = "Demographic data", skip = 1)
df_us <- read_excel(file_path, sheet = "초음파지표", skip = 1)

# 2. 데이터 전처리 및 병합
# 초음파 지표에서 필요한 컬럼(No, Composite Category) 선택
# V열은 22번째 컬럼이지만, 이름으로 선택하는 것이 안전합니다.
us_subset <- df_us %>%
  select(No, Group = `Composite Category`) %>%
  mutate(Group = str_trim(tolower(Group))) # 소문자 변환 및 공백 제거

# Demographic 데이터와 병합 (No 기준)
merged_df <- df_demo %>%
  inner_join(us_subset, by = "No")

# 3. 분석 대상 컬럼 정의
# B~F열: Bayley 점수들
bayley_cols <- c("Bayley - 인지", "Bayley - 언어", "Bayley - 운동", 
                 "Bayley - 사회 정서", "Bayley - 적응행동")
# I열: M-CHAT-R
mchat_col <- "M-CHAT-R (점수)"
# G, H열: K-M-B CDI
cdi_cols <- c("K-M-B CDI (표현낱말수)", "K-M-B CDI (문법사용)")

# 4. 통계 계산 함수 정의

# 평균 ± 표준편차 계산 함수 (소수점 둘째자리까지)
calc_mean_sd <- function(data, col_name) {
  vals <- data[[col_name]]
  vals <- as.numeric(vals) # 숫자로 변환
  vals <- vals[!is.na(vals)] # 결측치 제거
  
  if(length(vals) == 0) return("-")
  
  m <- mean(vals)
  s <- sd(vals)
  return(sprintf("%.2f ± %.2f", m, s))
}

# 빈도수(0, 1, 2) 계산 함수
calc_counts <- function(data, col_name) {
  vals <- data[[col_name]]
  vals <- as.numeric(vals)
  vals <- vals[!is.na(vals)]
  
  count_0 <- sum(vals == 0)
  count_1 <- sum(vals == 1)
  count_2 <- sum(vals == 2)
  
  return(c(count_0, count_1, count_2))
}

# 5. 그룹별 결과 계산 ('no abnormalities', 'mild abnormalities')
target_groups <- c("no abnormalities", "mild abnormalities")
result_list <- list()

for (grp in target_groups) {
  # 해당 그룹 데이터 필터링
  sub_data <- merged_df %>% filter(Group == grp)
  
  col_data <- c()
  row_names <- c()
  
  # Bayley Score (Mean ± SD)
  for (var in bayley_cols) {
    col_data <- c(col_data, calc_mean_sd(sub_data, var))
    row_names <- c(row_names, var)
  }
  
  # M-CHAT-R (Mean ± SD)
  col_data <- c(col_data, calc_mean_sd(sub_data, mchat_col))
  row_names <- c(row_names, mchat_col)
  
  # K-M-B CDI (Counts for 0, 1, 2)
  for (var in cdi_cols) {
    counts <- calc_counts(sub_data, var)
    col_data <- c(col_data, counts[1], counts[2], counts[3])
    row_names <- c(row_names, 
                   paste0(var, " (0점)"), 
                   paste0(var, " (1점)"), 
                   paste0(var, " (2점)"))
  }
  
  # 데이터프레임 생성
  temp_df <- data.frame(Variable = row_names, Value = col_data)
  colnames(temp_df)[2] <- grp # 컬럼명을 그룹명으로 변경
  
  result_list[[grp]] <- temp_df
}

# 6. 결과 병합
final_table <- result_list[[1]]
for (i in 2:length(result_list)) {
  final_table <- left_join(final_table, result_list[[i]], by = "Variable")
}

# 결과 확인
print(final_table)

# 7. 엑셀 파일로 저장
write_xlsx(final_table, path = "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/table_1_results.xlsx")
