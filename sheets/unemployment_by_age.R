# unemployment by age
# table: ons.labour_market_unemployment

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

fetch_unemployment_by_age <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    res <- DBI::dbGetQuery(conn, 'SELECT age_group, duration, value_type, dataset_identifier_code, time_period, value
FROM "ons"."labour_market_unemployment"')
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("fetch unemployment by age failed: ", e$message)
    tibble::tibble(
      age_group = character(), duration = character(), value_type = character(),
      dataset_identifier_code = character(), time_period = character(), value = numeric()
    )
  },
  finally = DBI::dbDisconnect(conn))
}

# "mar-may 1992" -> 1992-05-01 (end month)
parse_lfs_period_to_end_date <- function(label) {
  if (is.na(label) || !nzchar(label)) return(NA)
  label <- trimws(label)
  mons  <- regmatches(label, gregexpr("[A-Za-z]{3}", label))[[1]]
  yrs   <- regmatches(label, gregexpr("[0-9]{4}", label))[[1]]
  if (length(mons) < 2 || length(yrs) < 1) return(NA)
  mm <- month_map[[tolower(mons[2])]]
  yyyy <- suppressWarnings(as.integer(yrs[1]))
  if (is.null(mm) || is.na(yyyy)) return(NA)
  as.Date(sprintf("%04d-%02d-01", yyyy, mm))
}

compute_unemployment_by_age <- function(df, manual_mm) {
  if (is.null(df) || nrow(df) == 0)
    return(list(period = NA_character_, level = tibble(), rate = tibble()))

  df2 <- df %>%
    mutate(
      value    = suppressWarnings(as.numeric(value)),
      end_date = as.Date(vapply(time_period, parse_lfs_period_to_end_date, as.Date(NA)))
    ) %>%
    filter(!is.na(end_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, level = tibble(), rate = tibble()))

  latest_date  <- max(df2$end_date, na.rm = TRUE)
  latest_label <- df2 %>% filter(end_date == latest_date) %>% slice(1) %>% pull(time_period) %>% trimws()

  latest <- df2 %>% filter(end_date == latest_date)
  if ("All" %in% latest$duration) latest <- latest %>% filter(duration == "All")

  list(
    period = latest_label,
    level  = latest %>% filter(tolower(value_type) == "level") %>%
               select(age_group, duration, dataset_identifier_code, value) %>% arrange(desc(value)),
    rate   = latest %>% filter(grepl("rate", tolower(value_type))) %>%
               select(age_group, duration, dataset_identifier_code, value) %>% arrange(desc(value))
  )
}

calculate_unemployment_by_age <- function(manual_mm) {
  compute_unemployment_by_age(fetch_unemployment_by_age(), manual_mm)
}
