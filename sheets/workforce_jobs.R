# workforce jobs by industry
# table: ons.labour_market_workforce_jobs
# time period format: "mar 98 (r)" (quarterly, end-month)

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

fetch_workforce_jobs <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    res <- DBI::dbGetQuery(conn, 'SELECT industry, sic_section, time_period, value
FROM "ons"."labour_market_workforce_jobs"')
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("fetch workforce jobs failed: ", e$message)
    tibble::tibble(industry = character(), sic_section = character(), time_period = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

# "mar 98 (r)" -> 1998-03-01
parse_wfj_period_to_date <- function(x) {
  if (is.na(x) || !nzchar(x)) return(NA)
  x <- trimws(gsub("\\(.*\\)", "", x))
  parts <- strsplit(x, "\\s+")[[1]]
  if (length(parts) < 2) return(NA)
  mon <- tolower(substr(parts[1], 1, 3))
  yy  <- suppressWarnings(as.integer(parts[2]))
  mm  <- month_map[[mon]]
  if (is.null(mm) || is.na(yy)) return(NA)
  yyyy <- if (yy >= 50) 1900 + yy else 2000 + yy
  as.Date(sprintf("%04d-%02d-01", yyyy, mm))
}

latest_wfj_period_label <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(NA_character_)
  d <- df %>%
    mutate(period_date = as.Date(vapply(time_period, parse_wfj_period_to_date, as.Date(NA)))) %>%
    filter(!is.na(period_date)) %>%
    arrange(desc(period_date))
  if (nrow(d) == 0) return(NA_character_)
  trimws(d$time_period[1])
}

compute_workforce_jobs <- function(df, manual_mm) {
  if (is.null(df) || nrow(df) == 0) return(list(period = NA_character_, data = tibble()))

  df2 <- df %>%
    mutate(
      value       = suppressWarnings(as.numeric(value)),
      period_date = as.Date(vapply(time_period, parse_wfj_period_to_date, as.Date(NA)))
    ) %>%
    filter(!is.na(period_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, data = tibble()))

  latest_date  <- max(df2$period_date, na.rm = TRUE)
  latest_label <- df2 %>%
    filter(period_date == latest_date) %>%
    slice(1) %>%
    pull(time_period) %>%
    trimws()

  latest_data <- df2 %>%
    filter(period_date == latest_date) %>%
    select(industry, sic_section, value) %>%
    arrange(desc(value))

  list(period = latest_label, data = latest_data)
}

calculate_workforce_jobs <- function(manual_mm) {
  compute_workforce_jobs(fetch_workforce_jobs(), manual_mm)
}
