# sector payroll module - rtisa sheet 23
# sic sections: I = hospitality, G = retail, Q = health

SECTOR_CODES <- list(
  HOSPITALITY = "I",
  RETAIL      = "G",
  HEALTH      = "Q"
)

fetch_sector_payroll <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, sic_section, value
FROM "ons"."labour_market__employees_industry"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch sector payroll failed: ", e$message)
    tibble::tibble(time_period = character(), sic_section = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

val_sector <- function(pg_data, sic_section, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>%
    filter(sic_section == !!sic_section, trimws(time_period) == trimws(period_label))
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

compute_sector <- function(pg_data, manual_mm, sic_section) {
  anchor <- parse_manual_month(manual_mm) %m-% months(1)

  lab_cur <- make_payroll_label(anchor)
  lab_m   <- make_payroll_label(anchor %m-% months(1))
  lab_y   <- make_payroll_label(anchor %m-% months(12))
  lab_c   <- make_payroll_label(COVID_DATE)
  lab_e   <- make_payroll_label(ELEC24_DATE)

  cur   <- val_sector(pg_data, sic_section, lab_cur)
  val_m <- val_sector(pg_data, sic_section, lab_m)
  val_y <- val_sector(pg_data, sic_section, lab_y)
  val_c <- val_sector(pg_data, sic_section, lab_c)
  val_e <- val_sector(pg_data, sic_section, lab_e)

  list(
    cur    = cur / 1000,
    dm     = if (!is.na(cur) && !is.na(val_m)) (cur - val_m) / 1000 else NA_real_,
    dy     = if (!is.na(cur) && !is.na(val_y)) (cur - val_y) / 1000 else NA_real_,
    dc     = if (!is.na(cur) && !is.na(val_c)) (cur - val_c) / 1000 else NA_real_,
    de     = if (!is.na(cur) && !is.na(val_e)) (cur - val_e) / 1000 else NA_real_,
    anchor = anchor
  )
}

calculate_sector_payroll <- function(manual_mm) {
  pg_data <- fetch_sector_payroll()
  list(
    hospitality = compute_sector(pg_data, manual_mm, SECTOR_CODES$HOSPITALITY),
    retail      = compute_sector(pg_data, manual_mm, SECTOR_CODES$RETAIL),
    health      = compute_sector(pg_data, manual_mm, SECTOR_CODES$HEALTH)
  )
}
