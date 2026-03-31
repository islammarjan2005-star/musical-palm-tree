# redundancy module - a01 sheet 10

REDUND_CODE <- "BEIR"

fetch_redundancy <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__redundancies"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch redundancy failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

compute_redundancy <- function(pg_data, manual_mm,
                               covid_label = COVID_LFS_LABEL,
                               election_label = ELECTION_LABEL) {
  cm <- parse_manual_month(manual_mm)
  end_cur <- cm %m-% months(2)

  lab_cur <- make_lfs_label(end_cur)
  lab_q   <- make_lfs_label(end_cur %m-% months(3))
  lab_y   <- make_lfs_label(end_cur %m-% months(12))

  cur   <- val_by_code(pg_data, REDUND_CODE, lab_cur)
  val_q <- val_by_code(pg_data, REDUND_CODE, lab_q)
  val_y <- val_by_code(pg_data, REDUND_CODE, lab_y)
  val_c <- val_by_code(pg_data, REDUND_CODE, covid_label)
  val_e <- val_by_code(pg_data, REDUND_CODE, election_label)

  list(
    cur = cur,
    dq  = if (!is.na(cur) && !is.na(val_q)) cur - val_q else NA_real_,
    dy  = if (!is.na(cur) && !is.na(val_y)) cur - val_y else NA_real_,
    dc  = if (!is.na(cur) && !is.na(val_c)) cur - val_c else NA_real_,
    de  = if (!is.na(cur) && !is.na(val_e)) cur - val_e else NA_real_,
    end = end_cur
  )
}

calculate_redundancy <- function(manual_mm) {
  compute_redundancy(fetch_redundancy(), manual_mm)
}
