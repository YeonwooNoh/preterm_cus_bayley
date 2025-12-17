library(readxl)
library(dplyr)
library(writexl)
library(stringr)
library(effectsize)

# 2. 데이터 불러오기 및 전처리
file_path <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/data/data_final.xlsx"

df_demo <- read_excel(file_path, sheet = "Demographic data", skip = 1)
df_us <- read_excel(file_path, sheet = "초음파지표", skip = 1)

us_subset <- df_us %>%
  select(No, Group = `Composite Category`) %>%
  mutate(Group = str_trim(tolower(Group)))

merged_df <- df_demo %>%
  inner_join(us_subset, by = "No") %>%
  filter(Group %in% c("no abnormalities", "mild abnormalities"))

# 3. 분석 변수 리스트 정의
analysis_vars <- c(
  "Bayley - 인지", 
  "Bayley - 언어", 
  "Bayley - 운동", 
  "Bayley - 사회 정서", 
  "Bayley - 적응행동",
  "M-CHAT-R (점수)",
  "K-M-B CDI (표현낱말수)", 
  "K-M-B CDI (문법사용)"
)

# 4. 통계 분석 루프
stats_results <- data.frame()

for (var in analysis_vars) {
  
  # 데이터 추출 및 전처리
  temp_data <- merged_df %>%
    select(Group, Value = all_of(var)) %>%
    mutate(Value = as.numeric(Value)) %>%
    filter(!is.na(Value))
  
  # 데이터가 충분한지 확인
  if(n_distinct(temp_data$Group) == 2) {
    
    # 1) Mann-Whitney U Test
    mw_test <- wilcox.test(Value ~ Group, data = temp_data, exact = FALSE)
    
    # 2) Rank Biserial Correlation (95% CI 포함)
    # ci = 0.95는 기본값이지만 명시적으로 작성함
    rb_effect <- rank_biserial(Value ~ Group, data = temp_data, ci = 0.95)
    
    # 값 추출
    p_value <- mw_test$p.value
    r_rb <- rb_effect$r_rank_biserial
    ci_low <- rb_effect$CI_low
    ci_high <- rb_effect$CI_high
    
    # 유의성 표시
    sig_mark <- ifelse(p_value < 0.05, "*", "")
    
    # CI 포맷팅 (예: [0.12, 0.45])
    ci_str <- sprintf("[%.3f, %.3f]", ci_low, ci_high)
    
    stats_results <- rbind(stats_results, data.frame(
      Variable = var,
      Test_Method = "Mann-Whitney U",
      P_Value = p_value,
      P_Label = sprintf("%.3f %s", p_value, sig_mark),
      Rank_Biserial_Correlation = r_rb,
      Rank_Biserial_95_CI = ci_str  # 신뢰구간 컬럼 추가
    ))
    
  } else {
    # 분석 불가능한 경우
    stats_results <- rbind(stats_results, data.frame(
      Variable = var,
      Test_Method = "Data Insufficient",
      P_Value = NA,
      P_Label = "-",
      Rank_Biserial_Correlation = NA,
      Rank_Biserial_95_CI = NA
    ))
  }
}

# 5. 결과 정리 및 엑셀 저장
print(stats_results)

output_path <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/Table1_Statistical_Analysis_Results.xlsx"
write_xlsx(stats_results, path = output_path)

cat("\n[완료] 95% CI가 포함된 통계 분석 결과가 다음 경로에 저장되었습니다:\n", output_path)