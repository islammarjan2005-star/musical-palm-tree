# wages cpi module - x09 cpi-adjusted wages
# table: ons.labour_market__weekly_earnings_cpi

CPI_CURRENT <- list(
  TOTAL_EARNINGS_TYPE = "Total pay, seasonally adjusted",
  REG_EARNINGS_TYPE   = "Regular Pay, seasonally adjusted",
  EARNINGS_METRIC     = "% changes year on year",
  TIME_BASIS          = "3 month average"
)

CPI_CHANGE <- list(
  TOTAL_EARNINGS_TYPE = "Total pay, seasonally adjusted",
  REG_EARNINGS_TYPE   = "Regular Pay, seasonally adjusted",
  EARNINGS_METRIC     = "Real AWE(2015 £)",
  TIME_BASIS          = "nan"
)

fetch_wages_cpi <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, earnings_metric, earnings_type, time_basis, value
FROM "ons"."labour_market__weekly_earnings_cpi"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch cpi wages failed: ", e$message)
    tibble::tibble(
      time_period = character(), earnings_metric = character(),
      earnings_type = character(), time_basis = character(), value = numeric()
    )
  },
  finally = DBI::dbDisconnect(conn))
}

# current yoy % (3mo avg)
val_cpi_current <- function(pg_data, period_label, earnings_type) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>%
    filter(
      earnings_type == !!earnings_type,
      earnings_metric == CPI_CURRENT$EARNINGS_METRIC,
      time_basis == CPI_CURRENT$TIME_BASIS,
      trimws(time_period) == trimws(period_label)
    )
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

# real awe £ (single month)
val_cpi_raw <- function(pg_data, period_label, earnings_type) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>%
    filter(
      earnings_type == !!earnings_type,
      earnings_metric == CPI_CHANGE$EARNINGS_METRIC,
      time_basis == CPI_CHANGE$TIME_BASIS,
      trimws(time_period) == trimws(period_label)
    )
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

get_cpi_raw_avg <- function(pg_data, dates, earnings_type) {
  vals <- sapply(dates, function(d) val_cpi_raw(pg_data, make_datetime_label(d), earnings_type))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

compute_wages_cpi <- function(pg_data, manual_mm) {
  cm       <- parse_manual_month(manual_mm)
  anchor_m <- cm %m-% months(2)

  win3      <- seq(anchor_m,                 by = "-1 month", length.out = 3)
  prev3     <- seq(anchor_m %m-% months(3),  by = "-1 month", length.out = 3)
  yago3     <- seq(anchor_m %m-% months(12), by = "-1 month", length.out = 3)
  covid3    <- seq(COVID_DATE,               by = "-1 month", length.out = 3)
  election3 <- seq(ELEC24_DATE,              by = "-1 month", length.out = 3)
  dec2007   <- as.Date("2007-12-01")
  pandemic3 <- c(as.Date("2019-12-01"), as.Date("2020-01-01"), as.Date("2020-02-01"))

  # current yoy % (3mo avg)
  latest_total <- val_cpi_current(pg_data, make_datetime_label(anchor_m), CPI_CURRENT$TOTAL_EARNINGS_TYPE)
  latest_reg   <- val_cpi_current(pg_data, make_datetime_label(anchor_m), CPI_CURRENT$REG_EARNINGS_TYPE)

  # annualised £ changes: 3mo avg real awe difference * 52
  calc_change <- function(dates_a, dates_b, earnings_type) {
    a <- get_cpi_raw_avg(pg_data, dates_a, earnings_type)
    b <- get_cpi_raw_avg(pg_data, dates_b, earnings_type)
    if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
  }

  # current 3mo avg real awe (for historical comparisons)
  cur_total_awe <- get_cpi_raw_avg(pg_data, win3, CPI_CHANGE$TOTAL_EARNINGS_TYPE)
  cur_reg_awe   <- get_cpi_raw_avg(pg_data, win3, CPI_CHANGE$REG_EARNINGS_TYPE)

  dec2007_total  <- val_cpi_raw(pg_data, make_datetime_label(dec2007), CPI_CHANGE$TOTAL_EARNINGS_TYPE)
  dec2007_reg    <- val_cpi_raw(pg_data, make_datetime_label(dec2007), CPI_CHANGE$REG_EARNINGS_TYPE)
  pandemic_total <- get_cpi_raw_avg(pg_data, pandemic3, CPI_CHANGE$TOTAL_EARNINGS_TYPE)
  pandemic_reg   <- get_cpi_raw_avg(pg_data, pandemic3, CPI_CHANGE$REG_EARNINGS_TYPE)

  pct_above <- function(cur, base) {
    if (!is.na(cur) && !is.na(base) && base != 0) ((cur - base) / base) * 100 else NA_real_
  }

  list(
    total = list(
      cur             = latest_total,
      dq              = calc_change(win3, prev3,     CPI_CHANGE$TOTAL_EARNINGS_TYPE),
      dy              = calc_change(win3, yago3,     CPI_CHANGE$TOTAL_EARNINGS_TYPE),
      dc              = calc_change(win3, covid3,    CPI_CHANGE$TOTAL_EARNINGS_TYPE),
      de              = calc_change(win3, election3, CPI_CHANGE$TOTAL_EARNINGS_TYPE),
      pct_vs_dec2007  = pct_above(cur_total_awe, dec2007_total),
      pct_vs_pandemic = pct_above(cur_total_awe, pandemic_total)
    ),
    regular = list(
      cur             = latest_reg,
      dq              = calc_change(win3, prev3,     CPI_CHANGE$REG_EARNINGS_TYPE),
      dy              = calc_change(win3, yago3,     CPI_CHANGE$REG_EARNINGS_TYPE),
      dc              = calc_change(win3, covid3,    CPI_CHANGE$REG_EARNINGS_TYPE),
      de              = calc_change(win3, election3, CPI_CHANGE$REG_EARNINGS_TYPE),
      pct_vs_dec2007  = pct_above(cur_reg_awe, dec2007_reg),
      pct_vs_pandemic = pct_above(cur_reg_awe, pandemic_reg)
    ),
    anchor = anchor_m
  )
}

calculate_wages_cpi <- function(manual_mm) {
  compute_wages_cpi(fetch_wages_cpi(), manual_mm)
}
