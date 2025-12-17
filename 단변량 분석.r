## 필요한 패키지
library(readxl)
library(dplyr)
library(purrr)
library(writexl)


## 1. 데이터 불러오기
df <- read_excel(
  path  = "C:\\Users\\nyw02\\OneDrive\\바탕 화면\\연구과정\\R 코드\\data\\data_final.xlsx",          # <- 파일 경로/이름 맞게 수정
  sheet = "Demographic data",
  skip  = 1                     # 2행이 변수명이라 1행을 건너뜀
)

## 2. 예측변수(독립변수) 열 인덱스 정의 (엑셀 기준)
# A=1, B=2, C=3, D=4, E=5, F=6, ..., J=10, ..., Z=26, AA=27, AB=28, AC=29, AD=30, AE=31
continuous_idx <- c(12, 13, 19, 20, 27)                 # L, M, S, T, AA
binary_idx     <- c(14, 15, 16, 17, 18, 23, 24, 25, 26,
                    28, 30, 31, 32)                         # N, O, P, Q, R, W, X, Y, Z, AB, AD, AE, AF
multicat_idx   <- 21                                       # U
ordinal_idx    <- c(22, 29)                                # V, AC

continuous_vars <- colnames(df)[continuous_idx]
binary_vars     <- colnames(df)[binary_idx]
multicat_vars   <- colnames(df)[multicat_idx]
ordinal_vars    <- colnames(df)[ordinal_idx]

## 3. 단변량 분석 함수 (하나의 종속변수에 대해 실행)
run_univariate <- function(outcome_var_name) {
  y <- df[[outcome_var_name]]
  
  ## 3-1. 연속형 변수: Spearman 상관
  cont_results <- map_dfr(continuous_vars, function(var){
    x  <- df[[var]]
    cc <- complete.cases(y, x)
    
    if (sum(cc) < 3) {
      return(tibble(
        outcome  = outcome_var_name,
        variable = var,
        type     = "continuous",
        n        = sum(cc),
        stat     = NA_real_,
        p_value  = NA_real_
      ))
    }
    
    test <- suppressWarnings(
      cor.test(y[cc], x[cc], method = "spearman", exact = FALSE)
    )
    
    tibble(
      outcome  = outcome_var_name,
      variable = var,
      type     = "continuous",
      n        = sum(cc),
      stat     = unname(test$estimate),  # Spearman rho
      p_value  = test$p.value
    )
  })
  

  ## 3-2. 이진 변수: Mann-Whitney U test + rank biserial correlation
  bin_results <- map_dfr(binary_vars, function(var){
    x  <- df[[var]]
    cc <- complete.cases(y, x)
    
    # 두 그룹이 존재하지 않으면 NA 반환
    if (length(unique(x[cc])) < 2) {
      return(tibble(
        outcome  = outcome_var_name,
        variable = var,
        type     = "binary",
        n        = sum(cc),
        stat     = NA_real_,
        p_value  = NA_real_
      ))
    }
    
    group <- as.factor(x[cc])
    y_use <- y[cc]
    
    test  <- suppressWarnings(wilcox.test(y_use ~ group, exact = FALSE))
    
    # --- rank biserial correlation 계산 ---
    U <- unname(test$statistic)       # U 값
    n1 <- sum(group == levels(group)[1])
    n2 <- sum(group == levels(group)[2])
    
    r_rb <- 1 - (2 * U) / (n1 * n2)   # rank biserial correlation
    
    tibble(
      outcome  = outcome_var_name,
      variable = var,
      type     = "binary",
      n        = n1 + n2,
      stat     = r_rb,                # <-- 효과크기 저장
      p_value  = test$p.value
    )
  })
  
  ## 3-3. 다중범주형 변수: Kruskal-Wallis test
  multi_results <- map_dfr(multicat_vars, function(var){
    x  <- df[[var]]
    cc <- complete.cases(y, x)
    
    if (length(unique(x[cc])) < 2) {
      return(tibble(
        outcome  = outcome_var_name,
        variable = var,
        type     = "multicategory",
        n        = sum(cc),
        stat     = NA_real_,
        p_value  = NA_real_
      ))
    }
    
    group <- as.factor(x[cc])
    test  <- suppressWarnings(
      kruskal.test(y[cc] ~ group)
    )
    
    tibble(
      outcome  = outcome_var_name,
      variable = var,
      type     = "multicategory",
      n        = sum(cc),
      stat     = unname(test$statistic),  # Kruskal-Wallis chi-squared
      p_value  = test$p.value
    )
  })
  
  ## 3-4. 순서형 변수: Spearman 상관
  ord_results <- map_dfr(ordinal_vars, function(var){
    x  <- df[[var]]
    cc <- complete.cases(y, x)
    
    if (sum(cc) < 3) {
      return(tibble(
        outcome  = outcome_var_name,
        variable = var,
        type     = "ordinal",
        n        = sum(cc),
        stat     = NA_real_,
        p_value  = NA_real_
      ))
    }
    
    if (is.numeric(x)) {
      x_num <- x
    } else {
      x_num <- as.numeric(as.factor(x))
    }
    
    test <- suppressWarnings(
      cor.test(y[cc], x_num[cc], method = "spearman", exact = FALSE)
    )
    
    tibble(
      outcome  = outcome_var_name,
      variable = var,
      type     = "ordinal",
      n        = sum(cc),
      stat     = unname(test$estimate),  # Spearman rho
      p_value  = test$p.value
    )
  })
  
  ## 3-5. 하나로 합치기
  bind_rows(cont_results, bin_results, multi_results, ord_results) %>%
    arrange(type, p_value)
}

## 4. B, C, D, E, F, G, H, I열(2~9열)을 대상으로 반복 실행
outcome_indices <- 2:9
outcome_vars    <- colnames(df)[outcome_indices]

# 각 outcome별 결과를 리스트로 저장
results_list <- map(outcome_vars, run_univariate)
names(results_list) <- outcome_vars

# 모든 outcome 결과를 하나의 데이터프레임으로 합치기
all_results <- bind_rows(results_list)

write_xlsx(all_results,
           "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/univariate_results.xlsx")