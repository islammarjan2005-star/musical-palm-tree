# inactivity reasons module - a01 sheet 11

INACT_REASON_CODES <- list(
  STUDENT        = "LF63",
  FAMILY_HOME    = "LF65",
  TEMP_SICK      = "LF67",
  LONG_TERM_SICK = "LF69",
  DISCOURAGED    = "LFL8",
  RETIRED        = "LF6B",
  OTHER          = "LF6D"
)

INACT_REASON_NAMES <- list(
  LF63 = "students",
  LF65 = "those looking after family or home",
  LF67 = "temporary sickness",
  LF69 = "long-term sickness",
  LFL8 = "discouraged workers",
  LF6B = "retirees",
  LF6D = "other reasons"
)

fetch_inactivity_reasons <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  tryCatch({
    result <- DBI::dbGetQuery(conn, 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__inactivity"')
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("fetch inactivity reasons failed: ", e$message)
    tibble::tibble(time_period = character(), dataset_identifier_code = character(), value = numeric())
  },
  finally = DBI::dbDisconnect(conn))
}

compute_inactivity_reasons <- function(pg_data, manual_mm) {
  cm      <- parse_manual_month(manual_mm)
  end_cur <- cm %m-% months(2)

  lab_cur   <- make_lfs_label(end_cur)
  lab_y     <- make_lfs_label(end_cur %m-% months(12))
  lab_covid <- COVID_LFS_LABEL

  reasons <- c("LF63", "LF65", "LF67", "LF69", "LFL8", "LF6B", "LF6D")
  results <- list()

  for (code in reasons) {
    cur      <- val_by_code(pg_data, code, lab_cur)
    val_y    <- val_by_code(pg_data, code, lab_y)
    val_covid <- val_by_code(pg_data, code, lab_covid)

    results[[code]] <- list(
      cur = cur,
      dy  = if (!is.na(cur) && !is.na(val_y))     cur - val_y     else NA_real_,
      dc  = if (!is.na(cur) && !is.na(val_covid))  cur - val_covid else NA_real_
    )
  }

  results$labels <- list(cur = lab_cur, y = lab_y, covid = lab_covid)
  results$end    <- end_cur
  results
}

find_top_n_drivers <- function(inact_data, comparison = "dc", n = 2) {
  reasons <- c("LF63", "LF65", "LF67", "LF69", "LFL8", "LF6B", "LF6D")
  changes <- sapply(reasons, function(code) inact_data[[code]][[comparison]])
  df <- data.frame(
    code   = reasons,
    name   = sapply(reasons, function(x) INACT_REASON_NAMES[[x]]),
    change = changes,
    stringsAsFactors = FALSE
  )
  head(df[order(-abs(df$change)), ], n)
}

generate_inactivity_driver_text <- function(inact_data) {
  top_drivers <- find_top_n_drivers(inact_data, "dc", 2)
  if (nrow(top_drivers) == 0 || all(is.na(top_drivers$change))) return("changes across various reasons")
  increases <- top_drivers[top_drivers$change > 0, ]
  if (nrow(increases) == 0) {
    "decreases across various reasons"
  } else if (nrow(increases) == 1) {
    paste0("increases in ", increases$name[1])
  } else {
    paste0("increases in ", increases$name[1], " and ", increases$name[2])
  }
}

calculate_inactivity_reasons <- function(manual_mm) {
  compute_inactivity_reasons(fetch_inactivity_reasons(), manual_mm)
}
