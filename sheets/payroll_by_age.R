# payroll employees by age
# table: ons.labour_market_employees_age
# time period format: "january 2026"

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

fetch_payroll_by_age <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    res <- DBI::dbGetQuery(conn, 'SELECT age_group, time_period, value
FROM "ons"."labour_market_employees_age"')
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("fetch payroll by age failed: ", e$message)
    tibble::tibble(age_group = character(), time_period = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

# "january 2026" -> 2026-01-01
parse_month_label_to_date <- function(label) {
  if (is.na(label) || !nzchar(label)) return(NA)
  suppressWarnings(as.Date(paste0("01 ", trimws(label)), format = "%d %B %Y"))
}

compute_payroll_by_age <- function(df, manual_mm) {
  if (is.null(df) || nrow(df) == 0) return(list(period = NA_character_, data = tibble()))

  df2 <- df %>%
    mutate(
      value      = suppressWarnings(as.numeric(value)),
      month_date = as.Date(vapply(time_period, parse_month_label_to_date, as.Date(NA)))
    ) %>%
    filter(!is.na(month_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, data = tibble()))

  latest_date  <- max(df2$month_date, na.rm = TRUE)
  latest_label <- df2 %>% filter(month_date == latest_date) %>% slice(1) %>% pull(time_period) %>% trimws()

  list(
    period = latest_label,
    data   = df2 %>% filter(month_date == latest_date) %>% select(age_group, value) %>% arrange(desc(value))
  )
}

calculate_payroll_by_age <- function(manual_mm) {
  compute_payroll_by_age(fetch_payroll_by_age(), manual_mm)
}
