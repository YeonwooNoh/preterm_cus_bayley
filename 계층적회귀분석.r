# ==============================================================================
# 1. 기본 설정 및 데이터 불러오기
# ==============================================================================
# 필요한 패키지가 없다면 설치 후 로드 (최초 1회만 설치)
library(readxl)
library(dplyr)
library(broom)
library(MASS)
library(lmtest)
library(writexl)



# [중요] 데이터 파일 경로가 올바른지 확인하세요.
file_name <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/data/data_final.xlsx" 

# ------------------------------------------------------------------------------
# 데이터 로드
# 설명: 1행(대분류)은 건너뛰고(skip=1), 2행(변수명)을 헤더로 불러옵니다.
# ------------------------------------------------------------------------------
df_us <- read_excel(file_name, sheet = "초음파지표", skip = 1)
df_demo <- read_excel(file_name, sheet = "Demographic data", skip = 1)

# 데이터 병합 (환자 번호 'No' 기준)
data_merged <- merge(df_demo, df_us, by = "No", all = TRUE)

# ==============================================================================
# 2. [사용자 입력란] 8개 결과별 변수 세팅
# ==============================================================================
# 여기에 각 종속변수(Outcome)에 대해 '단변량에서 유의했던 임상지표(clinical)'와
# '다변량에서 선별된 초음파지표(us)'의 정확한 이름을 엑셀에서 복사해 넣어주세요.
# ※ 변수명은 따옴표("") 안에 넣고, 여러 개일 경우 콤마(,)로 구분합니다.

analysis_config <- list(
  
  # 1. Bayley - 인지
  list(
    name     = "Bayley_Cognitive",
    outcome  = "Bayley - 인지",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "BPD /c steroid", "invasive mechanical ventilation (days)", "Retinopathy of prematurity (stage)"),  # 예시: 임상변수 입력
    us       = c("Size of frontal horns - short axis (mm)")       # 예시: 초음파변수 입력
  ),
  
  # 2. Bayley - 언어
  list(
    name     = "Bayley_Language",
    outcome  = "Bayley - 언어",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "분만방식", "BPD /c steroid", "invasive mechanical ventilation (days)", "Retinopathy of prematurity (수술 여부)"), 
    us       = c("Corpus callosum thickness (mm)")
  ),
  
  # 3. Bayley - 운동
  list(
    name     = "Bayley_Motor",
    outcome  = "Bayley - 운동",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "invasive mechanical ventilation (days)", "Retinopathy of prematurity (수술 여부)"), 
    us       = c("Size of ventricular midbody (mm)", "Inter-hemispheric fissure (mm)")
  ),
  
  # 4. Bayley - 사회 정서
  list(
    name     = "Bayley_Social",
    outcome  = "Bayley - 사회 정서",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "invasive mechanical ventilation (days)", "Retinopathy of prematurity (stage)"), 
    us       = c("Corpus callosum thickness (mm)")
  ),
  
  # 5. Bayley - 적응행동
  list(
    name     = "Bayley_Adaptive",
    outcome  = "Bayley - 적응행동",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "BPD /c steroid", "invasive mechanical ventilation (days)", "Retinopathy of prematurity (stage)", "Retinopathy of prematurity (수술 여부)"), 
    us       = c("Sinu-cortical width (mm)")
  ),
  
  # 6. K-M-B CDI (표현낱말수)
  list(
    name     = "KMB_CDI_Words",
    outcome  = "K-M-B CDI (표현낱말수)",
    type="ordinal",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "BPD /c steroid", "invasive mechanical ventilation (days)", "Congenital infection"), 
    us       = c("Vermis height")
  ),
  
  # 7. K-M-B CDI (문법사용)
  list(
    name     = "KMB_CDI_Grammar",
    outcome  = "K-M-B CDI (문법사용)",
    type="ordinal",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "BPD /c steroid", "invasive mechanical ventilation (days)", "Sepsis"), 
    us       = c("Size of ventricular midbody (mm)", "Corpus callosum thickness (mm)")
  ),
  
  # 8. M-CHAT-R (점수)
  list(
    name     = "M_CHAT_R",
    outcome  = "M-CHAT-R (점수)",
    type="continuous",
    clinical = c("Gestational age (total)", "Birth weight (g)", "Gender", "RDS", "BPD /c steroid", "invasive mechanical ventilation (days)", "noninvasive mechanical ventilation (days)", "Congenital infection", "Parenteral nutrition (days)"), 
    us       = c("Size of ventricular midbody (mm)")
  )
)

# ==============================================================================
# 3. 분석 수행 함수 (수정 불필요)
# ==============================================================================

run_analysis <- function(config, data){

  y_var     <- trimws(config$outcome)
  clin_vars <- trimws(config$clinical)
  us_vars   <- trimws(config$us)
  type      <- config$type  # "continuous", "ordinal", "count"

  req_cols <- unique(c(y_var, clin_vars, us_vars))
  missing_cols <- setdiff(req_cols, names(data))
  if(length(missing_cols) > 0){
    return(data.frame(Outcome=y_var, Error=paste("변수명 오류:", paste(missing_cols, collapse=", "))))
  }

  data_clean <- na.omit(data[, req_cols, drop=FALSE])
  n_count <- nrow(data_clean)

  f1_str <- paste0("`", y_var, "` ~ ", paste(paste0("`", clin_vars, "`"), collapse=" + "))
    rhs2 <- c(paste0("`", clin_vars, "`"), paste0("`", us_vars, "`"))
    rhs2 <- rhs2[rhs2 != ""]
    f2_str <- paste0("`", y_var, "` ~ ", paste(rhs2, collapse=" + "))



  # --- 연속형: lm ---
  if(type == "continuous"){
    m1 <- lm(as.formula(f1_str), data=data_clean)
    m2 <- lm(as.formula(f2_str), data=data_clean)
    p_improve <- anova(m1, m2)$`Pr(>F)`[2]
    s1 <- summary(m1); s2 <- summary(m2)

    return(data.frame(
      Outcome_Variable=y_var, N_Patients=n_count,
      FitMetric="Adj_R2",
      Model1_Fit=round(s1$adj.r.squared, 3),
      Model2_Fit=round(s2$adj.r.squared, 3),
      Fit_Change=round(s2$adj.r.squared - s1$adj.r.squared, 3),
      P_for_Improvement=round(p_improve, 4),
      Significant_Improvement=ifelse(p_improve < 0.05, "YES", "NO")
    ))
  }

  # --- 순서형 3단계: polr (ordinal logistic) ---
  if(type == "ordinal"){
    data_clean[[y_var]] <- factor(data_clean[[y_var]], levels = c(0,1,2), ordered = TRUE)
  # 0<1<2 순서형

    m1 <- MASS::polr(as.formula(f1_str), data=data_clean, Hess=TRUE)
    m2 <- MASS::polr(as.formula(f2_str), data=data_clean, Hess=TRUE)

    # LRT로 개선 여부
    lr <- lmtest::lrtest(m1, m2)
    p_improve <- lr$`Pr(>Chisq)`[2]

    # McFadden pseudo R2
    mcfadden <- function(mod){
      ll  <- as.numeric(logLik(mod))
      ll0 <- as.numeric(logLik(update(mod, . ~ 1)))
      1 - ll/ll0
    }

    r1 <- mcfadden(m1); r2 <- mcfadden(m2)

    return(data.frame(
      Outcome_Variable=y_var, N_Patients=n_count,
      FitMetric="McFadden_R2",
      Model1_Fit=round(r1, 3),
      Model2_Fit=round(r2, 3),
      Fit_Change=round(r2 - r1, 3),
      P_for_Improvement=round(p_improve, 4),
      Significant_Improvement=ifelse(p_improve < 0.05, "YES", "NO")
    ))
  }

  # --- 점수/카운트: Poisson glm (원하면 NB로 확장 가능) ---
  if(type == "count"){
    m1 <- glm(as.formula(f1_str), data=data_clean, family=poisson())
    m2 <- glm(as.formula(f2_str), data=data_clean, family=poisson())
    p_improve <- anova(m1, m2, test="Chisq")$`Pr(>Chi)`[2]

    mcfadden <- function(mod){
      ll  <- as.numeric(logLik(mod))
      ll0 <- as.numeric(logLik(update(mod, . ~ 1)))
      1 - ll/ll0
    }
    r1 <- mcfadden(m1); r2 <- mcfadden(m2)

    return(data.frame(
      Outcome_Variable=y_var, N_Patients=n_count,
      FitMetric="McFadden_R2",
      Model1_Fit=round(r1, 3),
      Model2_Fit=round(r2, 3),
      Fit_Change=round(r2 - r1, 3),
      P_for_Improvement=round(p_improve, 4),
      Significant_Improvement=ifelse(p_improve < 0.05, "YES", "NO")
    ))
  }

  stop("type을 continuous / ordinal / count 중 하나로 넣어주세요.")
}

# ==============================================================================
# 4. 전체 실행 및 결과 저장
# ==============================================================================

# 위에서 설정한 리스트를 돌면서 분석 수행
results_list <- lapply(analysis_config, function(cfg) {
  run_analysis(cfg, data_merged)
})

# 결과를 하나의 표로 합치기
final_table <- do.call(rbind, results_list)

# 화면에 결과 출력
print(final_table)

# ------------------------------------------------------------------------------
# 엑셀(CSV) 파일로 저장
# 설명: 이 파일은 엑셀에서 바로 열 수 있습니다.
# ------------------------------------------------------------------------------
output_filename <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/Model_Comparison_Results.xlsx"
writexl::write_xlsx(final_table, output_filename)

message(paste0("\n[완료] 분석 결과가 아래 경로에 저장되었습니다:\n", output_filename))