library(readxl)
library(dplyr)
library(purrr)
library(broom)
library(MASS)
library(writexl)

#------------------------------------------------------------
# 1) 데이터 불러오기 & merge
#------------------------------------------------------------
path_xlsx <- "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/data/data_final.xlsx"

demo <- read_excel(path_xlsx, sheet = "Demographic data", skip = 1)
us   <- read_excel(path_xlsx, sheet = "초음파지표", skip = 1)

id_var <- colnames(demo)[1]  # ID assumed

us_idx  <- 24:36
us_vars <- colnames(us)[us_idx]

data_all <- demo %>%
  inner_join(
    us %>% dplyr::select(all_of(c(id_var, us_vars))),
    by = id_var
  )

#------------------------------------------------------------
# 2) outcome 변수 분리 (연속형 vs 순서형)
#    B~I = columns 2:9
#------------------------------------------------------------
outcome_vars <- colnames(demo)[2:9]
continuous_outcomes <- colnames(demo)[c(2,3,4,5,6,9)]
ordinal_outcomes    <- colnames(demo)[c(7,8)]

#------------------------------------------------------------
# 3) 보정변수 선택을 위해 index → 변수명 변환
#------------------------------------------------------------
demo_cols <- colnames(demo)

adjust_index_list <- list(
  c(12, 13, 14, 18, 19, 30),                  # Bayley-인지
  c(12, 13, 14, 16, 18, 19, 30),              # Bayley-언어
  c(12, 13, 14, 19, 30),                      # Bayley-운동
  c(12, 13, 14, 19, 29),                      # Bayley-사회정서
  c(12, 13, 14, 18, 19, 30, 31),              # Bayley-적응행동
  c(12, 13, 14, 18, 19, 24),                  # K-M-B CDI(표현낱말수)
  c(12, 13, 14, 18, 19, 25),                  # K-M-B CDI(문법사용)
  c(12, 13, 14, 17, 18, 19, 20, 24, 27)       # M-CHAT-R
)
names(adjust_index_list) <- outcome_vars

adjust_vars_list <- lapply(adjust_index_list, function(idx){
  if (length(idx) == 0) character(0) else demo_cols[idx]
})

#------------------------------------------------------------
# 4) 다변량 분석 함수: outcome + 보정변수 + 초음파 1개
#    lm()이면 beta 95% CI 추가
#    polr()이면 OR/CI는 기존대로
#------------------------------------------------------------
run_multiv_single_us <- function(outcome_var, us_var) {

  adjust_vars <- adjust_vars_list[[outcome_var]]
  vars_needed <- c(outcome_var, adjust_vars, us_var)

  dat <- data_all %>%
    dplyr::select(all_of(vars_needed)) %>%
    dplyr::filter(complete.cases(.))

  if (nrow(dat) < 10) {
    return(tibble(
      outcome   = outcome_var,
      us_var    = us_var,
      n         = nrow(dat),
      term      = us_var,
      estimate  = NA_real_,
      std.error = NA_real_,
      statistic = NA_real_,
      p.value   = NA_real_,
      beta_CI_lower = NA_real_,
      beta_CI_upper = NA_real_,
      OR        = NA_real_,
      CI_lower  = NA_real_,
      CI_upper  = NA_real_
    ))
  }

  safe_outcome <- paste0("`", outcome_var, "`")
  safe_adjust  <- if (length(adjust_vars) > 0) paste0("`", adjust_vars, "`") else character(0)
  safe_us      <- paste0("`", us_var, "`")

  formula_str <- paste(
    safe_outcome, "~",
    paste(c(safe_adjust, safe_us), collapse = " + ")
  )

  #-----------------------------
  # 연속형 outcome → lm()
  #-----------------------------
  if (outcome_var %in% continuous_outcomes) {

    fit <- lm(as.formula(formula_str), data = dat)

    res <- broom::tidy(fit) %>%
      dplyr::mutate(
        outcome = outcome_var,
        us_var  = us_var,
        n       = nrow(dat)
      ) %>%
      dplyr::select(outcome, us_var, n, dplyr::everything())

    res_us <- res %>%
      dplyr::mutate(term_clean = gsub("`", "", term)) %>%
      dplyr::filter(term_clean == us_var) %>%
      dplyr::select(-term_clean)

    if (nrow(res_us) == 0) {
      res_us <- tibble(
        outcome   = outcome_var,
        us_var    = us_var,
        n         = nrow(dat),
        term      = us_var,
        estimate  = NA_real_,
        std.error = NA_real_,
        statistic = NA_real_,
        p.value   = NA_real_
      )
    }

    # ---- beta 95% CI 추가 (t 기반) ----
    df_resid <- df.residual(fit)
    t_crit <- qt(0.975, df = df_resid)

    res_us <- res_us %>%
      dplyr::mutate(
        beta_CI_lower = estimate - t_crit * std.error,
        beta_CI_upper = estimate + t_crit * std.error,
        OR       = NA_real_,
        CI_lower = NA_real_,
        CI_upper = NA_real_
      )

    return(res_us)
  }

  #-----------------------------
  # 순서형 outcome → polr()
  #-----------------------------
  if (outcome_var %in% ordinal_outcomes) {

    dat[[outcome_var]] <- factor(dat[[outcome_var]], ordered = TRUE)

    fit <- polr(as.formula(formula_str), data = dat, Hess = TRUE)
    ct <- coef(summary(fit))

    z_vals <- ct[, "t value"]
    p_vals <- 2 * (1 - pnorm(abs(z_vals)))

    out <- as.data.frame(ct) %>%
      tibble::rownames_to_column("term") %>%
      dplyr::mutate(
        estimate  = .data$Value,
        std.error = .data$`Std. Error`,
        statistic = .data$`t value`,
        p.value   = p_vals
      ) %>%
      dplyr::select(term, estimate, std.error, statistic, p.value)

    z_crit <- qnorm(0.975)
    out <- out %>%
      dplyr::mutate(
        OR       = exp(estimate),
        CI_lower = exp(estimate - z_crit * std.error),
        CI_upper = exp(estimate + z_crit * std.error),
        beta_CI_lower = NA_real_,   # (beta CI는 여기서는 NA로 두기)
        beta_CI_upper = NA_real_
      )

    res <- out %>%
      dplyr::mutate(
        outcome = outcome_var,
        us_var  = us_var,
        n       = nrow(dat)
      ) %>%
      dplyr::select(outcome, us_var, n, term,
                    estimate, std.error, statistic, p.value,
                    beta_CI_lower, beta_CI_upper,
                    OR, CI_lower, CI_upper)

    res_us <- res %>%
      dplyr::mutate(term_clean = gsub("`", "", term)) %>%
      dplyr::filter(term_clean == us_var) %>%
      dplyr::select(-term_clean)

    if (nrow(res_us) == 0) {
      res_us <- tibble(
        outcome   = outcome_var,
        us_var    = us_var,
        n         = nrow(dat),
        term      = us_var,
        estimate  = NA_real_,
        std.error = NA_real_,
        statistic = NA_real_,
        p.value   = NA_real_,
        beta_CI_lower = NA_real_,
        beta_CI_upper = NA_real_,
        OR        = NA_real_,
        CI_lower  = NA_real_,
        CI_upper  = NA_real_
      )
    }

    return(res_us)
  }

  tibble(
    outcome   = outcome_var,
    us_var    = us_var,
    n         = nrow(dat),
    term      = us_var,
    estimate  = NA_real_,
    std.error = NA_real_,
    statistic = NA_real_,
    p.value   = NA_real_,
    beta_CI_lower = NA_real_,
    beta_CI_upper = NA_real_,
    OR        = NA_real_,
    CI_lower  = NA_real_,
    CI_upper  = NA_real_
  )
}

#------------------------------------------------------------
# 5) 모든 outcome × 모든 초음파 지표 실행
#------------------------------------------------------------
results_list_single <- map(
  outcome_vars,
  function(outc) {
    map_dfr(us_vars, ~ run_multiv_single_us(outcome_var = outc, us_var = .x))
  }
)
names(results_list_single) <- outcome_vars

multiv_results_single <- bind_rows(results_list_single)

#------------------------------------------------------------
# 6) 엑셀로 저장
#------------------------------------------------------------
write_xlsx(
  multiv_results_single,
  "C:/Users/nyw02/OneDrive/바탕 화면/연구과정/R 코드/multivariate_results_single_us.xlsx"
)
