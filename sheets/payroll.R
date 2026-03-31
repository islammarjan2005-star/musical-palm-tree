# payroll module - rtisa payrolled employees
# table: ons.labour_market__payrolled_employees

PAYROLL_UNIT_TYPE <- "Payrolled employees"

fetch_payroll <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, unit_type, value
FROM "ons"."labour_market__payrolled_employees"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch payroll failed: ", e$message)
    tibble::tibble(time_period = character(), unit_type = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

val_payroll <- function(pg_data, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>%
    filter(unit_type == PAYROLL_UNIT_TYPE, startsWith(time_period, period_label))
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

get_payroll_avg <- function(pg_data, dates) {
  vals <- sapply(dates, function(d) val_payroll(pg_data, make_payroll_label(d)))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

compute_payroll <- function(pg_data, manual_mm, mode = c("latest", "aligned")) {
  payroll_data <- pg_data %>%
    filter(unit_type == PAYROLL_UNIT_TYPE) %>%
    mutate(parsed_date = as.Date(paste0("01 ", time_period), format = "%d %B %Y")) %>%
    filter(!is.na(parsed_date)) %>%
    arrange(desc(parsed_date))

  empty <- list(
    cur = NA_real_, dq = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_,
    anchor = NA,
    flash_cur = NA_real_, flash_dy = NA_real_, flash_de = NA_real_, flash_dm = NA_real_,
    flash_anchor = NA
  )
  if (nrow(payroll_data) < 3) return(empty)

  mode <- match.arg(mode)

  # anchor: latest available, or aligned to dashboard reference quarter end
  if (mode == "aligned") {
    target_end <- tryCatch(parse_manual_month(manual_mm) %m-% months(2), error = function(e) NA)
    if (!is.na(target_end)) {
      idx_le <- which(payroll_data$parsed_date <= target_end)
      anchor <- if (length(idx_le) >= 1) payroll_data$parsed_date[idx_le[1]] else payroll_data$parsed_date[1]
    } else {
      anchor <- payroll_data$parsed_date[1]
    }
  } else {
    anchor <- payroll_data$parsed_date[1]
  }

  # 3-month windows
  win3     <- seq(anchor,                    by = "-1 month", length.out = 3)
  prev3    <- seq(anchor %m-% months(3),     by = "-1 month", length.out = 3)
  yago3    <- seq(anchor %m-% months(12),    by = "-1 month", length.out = 3)
  covid3   <- seq(COVID_DATE,                by = "-1 month", length.out = 3)
  election3 <- seq(ELEC24_DATE,              by = "-1 month", length.out = 3)

  cur   <- get_payroll_avg(pg_data, win3)
  val_q <- get_payroll_avg(pg_data, prev3)
  val_y <- get_payroll_avg(pg_data, yago3)
  val_c <- get_payroll_avg(pg_data, covid3)
  val_e <- get_payroll_avg(pg_data, election3)

  # changes in thousands
  dq <- if (!is.na(cur) && !is.na(val_q)) (cur - val_q) / 1000 else NA_real_
  dy <- if (!is.na(cur) && !is.na(val_y)) (cur - val_y) / 1000 else NA_real_
  dc <- if (!is.na(cur) && !is.na(val_c)) (cur - val_c) / 1000 else NA_real_
  de <- if (!is.na(cur) && !is.na(val_e)) (cur - val_e) / 1000 else NA_real_

  # flash estimate - latest single month
  flash_anchor <- payroll_data$parsed_date[1]
  flash_cur    <- suppressWarnings(as.numeric(payroll_data$value[1]))

  flash_val_y <- val_payroll(pg_data, make_payroll_label(flash_anchor %m-% months(12)))
  flash_val_e <- val_payroll(pg_data, make_payroll_label(ELEC24_DATE))
  flash_val_m <- val_payroll(pg_data, make_payroll_label(flash_anchor %m-% months(1)))

  flash_dy <- if (!is.na(flash_cur) && !is.na(flash_val_y)) (flash_cur - flash_val_y) / 1000 else NA_real_
  flash_de <- if (!is.na(flash_cur) && !is.na(flash_val_e)) (flash_cur - flash_val_e) / 1000 else NA_real_
  flash_dm <- if (!is.na(flash_cur) && !is.na(flash_val_m)) (flash_cur - flash_val_m) / 1000 else NA_real_

  list(
    cur          = cur / 1000,       # thousands for dashboard
    dq = dq, dy = dy, dc = dc, de = de,
    anchor       = anchor,
    flash_cur    = flash_cur / 1e6,  # millions for narrative
    flash_dy     = flash_dy,
    flash_de     = flash_de,
    flash_dm     = flash_dm,
    flash_anchor = flash_anchor
  )
}

calculate_payroll <- function(manual_mm, mode = c("latest", "aligned")) {
  compute_payroll(fetch_payroll(), manual_mm, mode = mode)
}
