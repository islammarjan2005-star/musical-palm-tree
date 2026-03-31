# hr1 module - redundancy notifications
# table: ons.labour_market__redundancies_region

HR1_REGION <- "GB"

fetch_hr1 <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, region, value
FROM "ons"."labour_market__redundancies_region"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch hr1 failed: ", e$message)
    tibble::tibble(time_period = character(), region = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

val_hr1 <- function(pg_data, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>% filter(region == HR1_REGION, startsWith(time_period, period_label))
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

compute_hr1 <- function(pg_data) {
  gb_data <- pg_data %>%
    filter(region == HR1_REGION) %>%
    mutate(parsed_date = as.Date(substr(time_period, 1, 10))) %>%
    filter(!is.na(parsed_date)) %>%
    arrange(desc(parsed_date))

  if (nrow(gb_data) == 0)
    return(list(cur = NA_real_, dm = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_, anchor = NA))

  anchor <- gb_data$parsed_date[1]
  cur    <- suppressWarnings(as.numeric(gb_data$value[1]))

  lab_m <- format(anchor %m-% months(1), "%Y-%m-%d")
  lab_y <- format(anchor %m-% months(12), "%Y-%m-%d")
  lab_c <- format(COVID_DATE,  "%Y-%m-%d")
  lab_e <- format(ELEC24_DATE, "%Y-%m-%d")

  val_m <- val_hr1(pg_data, lab_m)
  val_y <- val_hr1(pg_data, lab_y)
  val_c <- val_hr1(pg_data, lab_c)
  val_e <- val_hr1(pg_data, lab_e)

  list(
    cur    = cur,
    dm     = if (!is.na(cur) && !is.na(val_m)) cur - val_m else NA_real_,
    dy     = if (!is.na(cur) && !is.na(val_y)) cur - val_y else NA_real_,
    dc     = if (!is.na(cur) && !is.na(val_c)) cur - val_c else NA_real_,
    de     = if (!is.na(cur) && !is.na(val_e)) cur - val_e else NA_real_,
    anchor = anchor
  )
}

calculate_hr1 <- function() {
  compute_hr1(fetch_hr1())
}
