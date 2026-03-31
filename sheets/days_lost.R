# days lost module - a01 sheet 18 (working days lost)

DAYS_LOST_CODE <- "BBFW"

fetch_days_lost <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__disputes"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch days lost failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

compute_days_lost <- function(pg_data, manual_mm) {
  cm <- parse_manual_month(manual_mm)
  anchor <- cm %m-% months(2)
  lab_cur <- make_payroll_label(anchor)

  match_row <- pg_data %>%
    filter(dataset_identifier_code == DAYS_LOST_CODE, startsWith(time_period, lab_cur))

  cur <- if (nrow(match_row) == 0) NA_real_ else suppressWarnings(as.numeric(match_row$value[1]))

  list(cur = cur, label = lab_cur, anchor = anchor)
}

calculate_days_lost <- function(manual_mm) {
  compute_days_lost(fetch_days_lost(), manual_mm)
}
