# wages nominal module - a01 sheet 13 (total) + sheet 15 (regular)

WAGES_NOM_CODES <- list(
  WEEKLY_TOTAL      = "KAB9",
  YOY_TOTAL         = "KAC3",
  WEEKLY_REG        = "KAI7",
  YOY_REG           = "KAI9",
  YOY_TOTAL_PUBLIC  = "KAC9",
  YOY_TOTAL_PRIVATE = "KAC6",
  YOY_REG_PUBLIC    = "KAJ7",
  YOY_REG_PRIVATE   = "KAJ4"
)

fetch_wages_total <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__weekly_earnings_total"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch total wages failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

fetch_wages_regular <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__weekly_earnings_regular"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch regular wages failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

get_nom_val <- function(pg_data, date, code) {
  val_by_code(pg_data, code, make_ymd_label(date))
}

get_nom_avg <- function(pg_data, dates, code) {
  vals <- sapply(dates, function(d) get_nom_val(pg_data, d, code))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

compute_wages_nominal <- function(pg_total, pg_regular, manual_mm) {
  cm       <- parse_manual_month(manual_mm)
  anchor_m <- cm %m-% months(2)

  win3      <- seq(anchor_m,                 by = "-1 month", length.out = 3)
  prev3     <- seq(anchor_m %m-% months(3),  by = "-1 month", length.out = 3)
  yago3     <- seq(anchor_m %m-% months(12), by = "-1 month", length.out = 3)
  covid3    <- seq(COVID_DATE,               by = "-1 month", length.out = 3)
  election3 <- seq(ELEC24_DATE,              by = "-1 month", length.out = 3)

  # current yoy % (single month)
  latest_total <- get_nom_val(pg_total, anchor_m, WAGES_NOM_CODES$YOY_TOTAL)
  latest_reg   <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG)

  # public/private sector yoy %
  total_public  <- get_nom_val(pg_total,   anchor_m, WAGES_NOM_CODES$YOY_TOTAL_PUBLIC)
  total_private <- get_nom_val(pg_total,   anchor_m, WAGES_NOM_CODES$YOY_TOTAL_PRIVATE)
  reg_public    <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG_PUBLIC)
  reg_private   <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG_PRIVATE)

  # quarter-on-quarter change in yoy %
  prev_q_anchor <- anchor_m %m-% months(3)
  total_qchange <- {
    v1 <- get_nom_val(pg_total,   anchor_m,     WAGES_NOM_CODES$YOY_TOTAL)
    v2 <- get_nom_val(pg_total,   prev_q_anchor, WAGES_NOM_CODES$YOY_TOTAL)
    if (!is.na(v1) && !is.na(v2)) v1 - v2 else NA_real_
  }
  reg_qchange <- {
    v1 <- get_nom_val(pg_regular, anchor_m,     WAGES_NOM_CODES$YOY_REG)
    v2 <- get_nom_val(pg_regular, prev_q_anchor, WAGES_NOM_CODES$YOY_REG)
    if (!is.na(v1) && !is.na(v2)) v1 - v2 else NA_real_
  }

  # annualised £ changes: 3mo avg weekly difference * 52
  calc_change <- function(pg, dates_a, dates_b, code) {
    a <- get_nom_avg(pg, dates_a, code)
    b <- get_nom_avg(pg, dates_b, code)
    if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
  }

  list(
    total = list(
      cur     = latest_total,
      dq      = calc_change(pg_total, win3, prev3,     WAGES_NOM_CODES$WEEKLY_TOTAL),
      dy      = calc_change(pg_total, win3, yago3,     WAGES_NOM_CODES$WEEKLY_TOTAL),
      dc      = calc_change(pg_total, win3, covid3,    WAGES_NOM_CODES$WEEKLY_TOTAL),
      de      = calc_change(pg_total, win3, election3, WAGES_NOM_CODES$WEEKLY_TOTAL),
      public  = total_public,
      private = total_private,
      qchange = total_qchange
    ),
    regular = list(
      cur     = latest_reg,
      dq      = calc_change(pg_regular, win3, prev3,     WAGES_NOM_CODES$WEEKLY_REG),
      dy      = calc_change(pg_regular, win3, yago3,     WAGES_NOM_CODES$WEEKLY_REG),
      dc      = calc_change(pg_regular, win3, covid3,    WAGES_NOM_CODES$WEEKLY_REG),
      de      = calc_change(pg_regular, win3, election3, WAGES_NOM_CODES$WEEKLY_REG),
      public  = reg_public,
      private = reg_private,
      qchange = reg_qchange
    ),
    anchor = anchor_m
  )
}

calculate_wages_nominal <- function(manual_mm) {
  compute_wages_nominal(fetch_wages_total(), fetch_wages_regular(), manual_mm)
}
