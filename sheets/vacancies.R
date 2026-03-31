# vacancies module - a01 sheet 19

VAC_CODES <- list(
  VAC     = "AP2Y",
  QCHANGE = "AP3K"
)

fetch_vacancies <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__vacancies_business"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch vacancies failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

# "jul-sep 2025" -> 2025-09-01 (end month)
parse_lfs_label_to_date <- function(label) {
  matches <- regmatches(label, gregexpr("[A-Za-z]{3}", label))[[1]]
  year    <- regmatches(label, gregexpr("[0-9]{4}", label))[[1]]
  if (length(matches) >= 2 && length(year) >= 1) {
    end_month <- month_map[tolower(matches[2])]
    yr <- as.integer(year[1])
    if (!is.na(end_month) && !is.na(yr))
      return(as.Date(sprintf("%04d-%02d-01", yr, end_month)))
  }
  NA
}

compute_vacancies <- function(pg_data, manual_mm,
                              mode = c("latest", "aligned"),
                              covid_label = COVID_VAC_LABEL,
                              election_label = ELECTION_LABEL) {
  vac_data <- pg_data %>%
    filter(dataset_identifier_code == VAC_CODES$VAC) %>%
    mutate(parsed_date = sapply(time_period, parse_lfs_label_to_date)) %>%
    mutate(parsed_date = as.Date(parsed_date, origin = "1970-01-01")) %>%
    filter(!is.na(parsed_date)) %>%
    arrange(desc(parsed_date))

  if (nrow(vac_data) == 0)
    return(list(cur = NA_real_, dq = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_, end = NA))

  mode <- match.arg(mode)

  # aligned: pick latest period <= dashboard reference quarter end
  if (mode == "aligned") {
    target_end <- tryCatch(parse_manual_month(manual_mm) %m-% months(2), error = function(e) NA)
    if (!is.na(target_end)) {
      idx_exact <- which(vac_data$parsed_date == target_end)
      if (length(idx_exact) >= 1) {
        pick <- idx_exact[1]
      } else {
        idx_le <- which(vac_data$parsed_date <= target_end)
        pick <- if (length(idx_le) >= 1) idx_le[1] else 1
      }
    } else {
      pick <- 1
    }
  } else {
    pick <- 1
  }

  end_cur <- vac_data$parsed_date[pick]
  lab_cur <- trimws(vac_data$time_period[pick])
  lab_y   <- make_lfs_label(end_cur %m-% months(12))

  cur   <- val_by_code(pg_data, VAC_CODES$VAC, lab_cur)
  val_y <- val_by_code(pg_data, VAC_CODES$VAC, lab_y)
  val_c <- val_by_code(pg_data, VAC_CODES$VAC, covid_label)
  val_e <- val_by_code(pg_data, VAC_CODES$VAC, election_label)
  dq    <- val_by_code(pg_data, VAC_CODES$QCHANGE, lab_cur)

  list(
    cur = cur,
    dq  = dq,
    dy  = if (!is.na(cur) && !is.na(val_y)) cur - val_y else NA_real_,
    dc  = if (!is.na(cur) && !is.na(val_c)) cur - val_c else NA_real_,
    de  = if (!is.na(cur) && !is.na(val_e)) cur - val_e else NA_real_,
    end = end_cur
  )
}

calculate_vacancies <- function(manual_mm, mode = c("latest", "aligned")) {
  compute_vacancies(fetch_vacancies(), manual_mm, mode = mode)
}
