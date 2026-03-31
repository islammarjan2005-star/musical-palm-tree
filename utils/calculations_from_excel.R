# calculations_from_excel.R — reads LMS CSV + optional Excel files and populates dashboard variables

suppressPackageStartupMessages({
  library(readxl)
  library(lubridate)
})

if (!exists("parse_manual_month", inherits = TRUE)) source("utils/helpers.R")
if (!exists("COVID_DATE",         inherits = TRUE)) source("utils/config.R")

# ---- Excel helpers (kept for HR1 and RTISA) ----

.read_sheet <- function(path, sheet) {
  tryCatch(
    suppressMessages(readxl::read_excel(path, sheet = sheet, col_names = FALSE)),
    error = function(e) {
      warning("failed to read sheet '", sheet, "' from ", basename(path), ": ", e$message)
      data.frame()
    }
  )
}

.find_row <- function(tbl, label) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(NA_integer_)
  col1 <- trimws(as.character(tbl[[1]]))
  label <- trimws(label)
  idx <- which(tolower(col1) == tolower(label))
  if (length(idx) == 0) return(NA_integer_)
  idx[1]
}

.cell_num <- function(tbl, row, col) {
  if (is.na(row) || row < 1 || row > nrow(tbl) || col > ncol(tbl)) return(NA_real_)
  x <- as.character(tbl[[col]][row])
  suppressWarnings(as.numeric(gsub("[^0-9.eE+-]", "", x)))
}

.lfs_label <- function(end_date) {
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

.detect_dates <- function(x) {
  if (inherits(x, "Date")) return(floor_date(as.Date(x), "month"))
  if (inherits(x, c("POSIXct", "POSIXt"))) return(floor_date(as.Date(x), "month"))
  s <- as.character(x)
  num <- suppressWarnings(as.numeric(s))
  is_num <- !is.na(num) & grepl("^[0-9]+\\.?[0-9]*$", s)
  out <- rep(as.Date(NA), length(s))
  if (any(is_num)) out[is_num] <- as.Date(num[is_num], origin = "1899-12-30")
  if (any(!is_num)) {
    out[!is_num] <- suppressWarnings(
      lubridate::parse_date_time(
        s[!is_num],
        orders = c("ymd", "mdy", "dmy", "bY", "BY", "Y b", "b Y", "Ym", "my")
      )
    )
  }
  floor_date(as.Date(out), "month")
}

.val_by_date <- function(df_m, df_v, target_date) {
  idx <- which(df_m == target_date)
  if (length(idx) == 0) return(NA_real_)
  df_v[idx[1]]
}

.avg_by_dates <- function(df_m, df_v, target_dates) {
  vals <- vapply(target_dates, function(d) .val_by_date(df_m, df_v, d), numeric(1))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

.safe_last <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) return(NA_real_)
  x[length(x)]
}

.find_col_by_code <- function(tbl, code, fallback_col = NA_integer_, search_rows = 1:min(10, nrow(tbl))) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(fallback_col)
  for (r in search_rows) {
    for (c in seq_len(ncol(tbl))) {
      cell <- as.character(tbl[[c]][r])
      if (!is.na(cell) && grepl(code, cell, fixed = TRUE)) return(c)
    }
  }
  fallback_col
}

# ---- LMS CSV reader ----

.read_lms_csv <- function(path) {
  raw <- read.csv(path, header = FALSE, stringsAsFactors = FALSE, check.names = FALSE)
  cdids <- as.character(raw[2, -1])
  titles <- as.character(raw[1, -1])
  data_rows <- raw[-(1:7), ]
  date_labels <- trimws(as.character(data_rows[[1]]))
  month_map_upper <- c(JAN=1, FEB=2, MAR=3, APR=4, MAY=5, JUN=6,
                        JUL=7, AUG=8, SEP=9, OCT=10, NOV=11, DEC=12)
  dates <- vapply(date_labels, function(lbl) {
    parts <- strsplit(trimws(lbl), "\\s+")[[1]]
    if (length(parts) == 2) {
      yr <- suppressWarnings(as.integer(parts[1]))
      mon <- month_map_upper[toupper(parts[2])]
      if (!is.na(yr) && !is.na(mon)) return(as.numeric(as.Date(sprintf("%04d-%02d-01", yr, mon))))
    } else if (length(parts) == 1) {
      yr <- suppressWarnings(as.integer(parts[1]))
      if (!is.na(yr) && yr > 1800 && yr < 2100) return(as.numeric(as.Date(sprintf("%04d-01-01", yr))))
    }
    NA_real_
  }, numeric(1))
  dates <- as.Date(dates, origin = "1970-01-01")
  vals <- as.matrix(data_rows[, -1])
  storage.mode(vals) <- "character"
  vals_num <- matrix(suppressWarnings(as.numeric(vals)), nrow = nrow(vals), ncol = ncol(vals))
  list(dates = dates, cdids = cdids, titles = titles, values = vals_num)
}

.lms_col <- function(lms, cdid) {
  idx <- match(cdid, lms$cdids)
  if (is.na(idx)) return(rep(NA_real_, length(lms$dates)))
  lms$values[, idx]
}

.lms_val <- function(lms, cdid, target_date) {
  vals <- .lms_col(lms, cdid)
  .val_by_date(lms$dates, vals, target_date)
}

.lms_avg <- function(lms, cdid, target_dates) {
  vals <- .lms_col(lms, cdid)
  .avg_by_dates(lms$dates, vals, target_dates)
}

.lms_metric <- function(lms, cdid, target_dates) {
  vals <- vapply(target_dates, function(d) .lms_val(lms, cdid, d), numeric(1))
  names(vals) <- c("cur", "q", "y", "covid", "elec")
  list(
    cur = vals["cur"],
    dq  = vals["cur"] - vals["q"],
    dy  = vals["cur"] - vals["y"],
    dc  = vals["cur"] - vals["covid"],
    de  = vals["cur"] - vals["elec"]
  )
}

.lms_last_valid <- function(lms, cdid) {
  vals <- .lms_col(lms, cdid)
  ok <- which(!is.na(lms$dates) & !is.na(vals))
  if (length(ok) == 0) return(list(date = as.Date(NA), value = NA_real_))
  last <- ok[length(ok)]
  list(date = lms$dates[last], value = vals[last])
}

# detect reference month from LMS csv
.detect_manual_month_from_lms <- function(file_lms) {
  if (is.null(file_lms)) return(NULL)
  tryCatch({
    lms <- .read_lms_csv(file_lms)
    last <- .lms_last_valid(lms, "MGRZ")
    if (is.na(last$date)) return(NULL)
    lfs_end <- last$date
    cm_date <- lfs_end %m+% months(2)
    tolower(paste0(format(cm_date, "%b"), format(cm_date, "%Y")))
  }, error = function(e) NULL)
}

# keep backward compat
.detect_manual_month_from_a01 <- .detect_manual_month_from_lms


# ---- OECD data from LMS ----

.read_oecd_from_lms <- function(file_lms) {
  lms <- .read_lms_csv(file_lms)

  oecd_cdids <- list(
    "United Kingdom" = list(unemp = "ZXDW", emp = "ANZ6", inact = "LF2S"),
    "United States"  = list(unemp = "ZXDX", emp = "A48Q", inact = NULL),
    "France"         = list(unemp = "ZXDN", emp = "YXSR", inact = NULL),
    "Germany"        = list(unemp = "ZXDK", emp = "YXSS", inact = NULL),
    "Italy"          = list(unemp = "ZXDP", emp = "YXSV", inact = NULL),
    "Spain"          = list(unemp = "ZXDM", emp = "YXSZ", inact = NULL),
    "Canada"         = list(unemp = "ZXDZ", emp = "A48O", inact = NULL),
    "Japan"          = list(unemp = "ZXDY", emp = "A48P", inact = NULL),
    "Euro area"      = list(unemp = "A493", emp = "A496", inact = NULL)
  )

  .extract_latest <- function(metric) {
    countries <- character(0); periods <- character(0); values <- numeric(0)
    for (country in names(oecd_cdids)) {
      cdid <- oecd_cdids[[country]][[metric]]
      if (is.null(cdid) || !cdid %in% lms$cdids) next
      last <- .lms_last_valid(lms, cdid)
      if (is.na(last$date) || is.na(last$value)) next
      countries <- c(countries, country)
      periods <- c(periods, format(last$date, "%b %Y"))
      values <- c(values, last$value)
    }
    if (length(countries) == 0) return(NULL)
    data.frame(country = countries, period = periods, value = values, stringsAsFactors = FALSE)
  }

  list(
    unemp = .extract_latest("unemp"),
    emp   = .extract_latest("emp"),
    inact = .extract_latest("inact")
  )
}

# ---- main entry point ----

run_calculations_from_excel <- function(manual_month = NULL,
                                        file_lms = NULL,
                                        file_hr1 = NULL,
                                        file_rtisa = NULL,
                                        vac_end_override = NULL,
                                        payroll_end_override = NULL,
                                        target_env = globalenv()) {

  if (is.null(manual_month)) {
    manual_month <- .detect_manual_month_from_lms(file_lms)
    if (is.null(manual_month)) {
      stop("cannot detect reference period from LMS file", call. = FALSE)
    }
  }

  cm       <- parse_manual_month(manual_month)
  anchor_m <- cm %m-% months(2)

  lfs_end_cur   <- anchor_m
  lfs_end_q     <- anchor_m %m-% months(3)
  lfs_end_y     <- anchor_m %m-% months(12)
  lfs_end_covid <- COVID_DATE
  lfs_end_elec  <- ELEC24_DATE

  all_dates <- c(lfs_end_cur, lfs_end_q, lfs_end_y, lfs_end_covid, lfs_end_elec)

  # read LMS
  lms <- if (!is.null(file_lms)) .read_lms_csv(file_lms) else NULL

  # ---- LFS metrics from LMS ----
  .safe_lms_metric <- function(cdid) {
    if (is.null(lms)) {
      return(list(cur = NA_real_, dq = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_))
    }
    .lms_metric(lms, cdid, all_dates)
  }

  m_emp16    <- .safe_lms_metric("MGRZ")   # employment level 16+
  m_emprt    <- .safe_lms_metric("LF24")   # employment rate 16-64
  m_unemp16  <- .safe_lms_metric("MGSC")   # unemployment level 16+
  m_unemprt  <- .safe_lms_metric("MGSX")   # unemployment rate 16+
  m_inact    <- .safe_lms_metric("LF2M")   # inactivity level 16-64
  m_inactrt  <- .safe_lms_metric("LF2S")   # inactivity rate 16-64

  for (prefix in c("emp16", "emp_rt", "unemp16", "unemp_rt", "inact", "inact_rt")) {
    m <- switch(prefix,
                emp16 = m_emp16, emp_rt = m_emprt,
                unemp16 = m_unemp16, unemp_rt = m_unemprt,
                inact = m_inact, inact_rt = m_inactrt)
    assign(paste0(prefix, "_cur"), m$cur, envir = target_env)
    assign(paste0(prefix, "_dq"),  m$dq,  envir = target_env)
    assign(paste0(prefix, "_dy"),  m$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  m$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  m$de,  envir = target_env)
  }

  # inactivity 50-64
  m_5064    <- .safe_lms_metric("LF2A")
  m_5064rt  <- .safe_lms_metric("LF2W")

  for (prefix in c("inact5064", "inact5064_rt")) {
    m <- if (prefix == "inact5064") m_5064 else m_5064rt
    assign(paste0(prefix, "_cur"), m$cur, envir = target_env)
    assign(paste0(prefix, "_dq"),  m$dq,  envir = target_env)
    assign(paste0(prefix, "_dy"),  m$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  m$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  m$de,  envir = target_env)
  }

  # redundancy (LFS)
  m_redund       <- .safe_lms_metric("BEIR")
  m_redund_level <- .safe_lms_metric("BEAO")

  assign("redund_cur", m_redund$cur, envir = target_env)
  assign("redund_dq",  m_redund$dq,  envir = target_env)
  assign("redund_dy",  m_redund$dy,  envir = target_env)
  assign("redund_dc",  m_redund$dc,  envir = target_env)
  assign("redund_de",  m_redund$de,  envir = target_env)

  # ---- Wages nominal from LMS ----
  if (!is.null(lms)) {
    latest_wages <- .lms_val(lms, "KAC3", anchor_m)

    win3      <- seq(anchor_m %m-% months(2), by = "month", length.out = 3)
    prev3     <- seq(anchor_m %m-% months(5), by = "month", length.out = 3)
    yago3     <- win3 %m-% months(12)
    covid3    <- seq(COVID_DATE  %m-% months(2), by = "month", length.out = 3)
    election3 <- seq(ELEC24_DATE %m-% months(2), by = "month", length.out = 3)

    .wage_change_lms <- function(a_months, b_months) {
      a <- .lms_avg(lms, "KAB9", a_months)
      b <- .lms_avg(lms, "KAB9", b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    wages_change_q        <- .wage_change_lms(win3, prev3)
    wages_change_y        <- .wage_change_lms(win3, yago3)
    wages_change_covid    <- .wage_change_lms(win3, covid3)
    wages_change_election <- .wage_change_lms(win3, election3)

    prev_q_pct <- .lms_val(lms, "KAC3", anchor_m %m-% months(3))
    wages_total_qchange <- if (!is.na(latest_wages) && !is.na(prev_q_pct)) latest_wages - prev_q_pct else NA_real_

    latest_regular_cash <- .lms_val(lms, "KAI9", anchor_m)
    prev_q_reg <- .lms_val(lms, "KAI9", anchor_m %m-% months(3))
    wages_reg_qchange <- if (!is.na(latest_regular_cash) && !is.na(prev_q_reg)) latest_regular_cash - prev_q_reg else NA_real_

    wages_total_public  <- .lms_val(lms, "KAC9", anchor_m)
    wages_total_private <- .lms_val(lms, "KAC6", anchor_m)
    wages_reg_public    <- .lms_val(lms, "KAJ7", anchor_m)
    wages_reg_private   <- .lms_val(lms, "KAJ4", anchor_m)
  } else {
    latest_wages <- wages_change_q <- wages_change_y <- NA_real_
    wages_change_covid <- wages_change_election <- wages_total_qchange <- NA_real_
    latest_regular_cash <- wages_reg_qchange <- NA_real_
    wages_total_public <- wages_total_private <- NA_real_
    wages_reg_public <- wages_reg_private <- NA_real_
    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
  }

  assign("latest_wages",          latest_wages,          envir = target_env)
  assign("wages_change_q",        wages_change_q,        envir = target_env)
  assign("wages_change_y",        wages_change_y,        envir = target_env)
  assign("wages_change_covid",    wages_change_covid,    envir = target_env)
  assign("wages_change_election", wages_change_election, envir = target_env)
  assign("wages_total_public",    wages_total_public,    envir = target_env)
  assign("wages_total_private",   wages_total_private,   envir = target_env)
  assign("wages_total_qchange",   wages_total_qchange,   envir = target_env)
  assign("latest_regular_cash",   latest_regular_cash,   envir = target_env)
  assign("wages_reg_public",      wages_reg_public,      envir = target_env)
  assign("wages_reg_private",     wages_reg_private,     envir = target_env)
  assign("wages_reg_qchange",     wages_reg_qchange,     envir = target_env)

  # ---- Real wages CPI from LMS ----
  if (!is.null(lms)) {
    cpi_last <- .lms_last_valid(lms, "A3WW")
    cpi_anchor <- if (!is.na(cpi_last$date)) cpi_last$date else anchor_m

    latest_wages_cpi   <- .lms_val(lms, "A3WW", cpi_anchor)
    latest_regular_cpi <- .lms_val(lms, "A2FA", cpi_anchor)

    .cpi_change_lms <- function(a_months, b_months) {
      a <- .lms_avg(lms, "A3WX", a_months)
      b <- .lms_avg(lms, "A3WX", b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    cpi_win3      <- seq(cpi_anchor %m-% months(2), by = "month", length.out = 3)
    prev3_cpi     <- seq(cpi_anchor %m-% months(5), by = "month", length.out = 3)
    yago3_cpi     <- cpi_win3 %m-% months(12)
    covid3_cpi    <- seq(COVID_DATE  %m-% months(2), by = "month", length.out = 3)
    election3_cpi <- seq(ELEC24_DATE %m-% months(2), by = "month", length.out = 3)

    wages_cpi_change_q        <- .cpi_change_lms(cpi_win3, prev3_cpi)
    wages_cpi_change_y        <- .cpi_change_lms(cpi_win3, yago3_cpi)
    wages_cpi_change_covid    <- .cpi_change_lms(cpi_win3, covid3_cpi)
    wages_cpi_change_election <- .cpi_change_lms(cpi_win3, election3_cpi)

    dec2007_val  <- .lms_val(lms, "A3WX", as.Date("2007-12-01"))
    cur_cpi_real <- .lms_avg(lms, "A3WX", cpi_win3)
    wages_cpi_total_vs_dec2007 <- if (!is.na(cur_cpi_real) && !is.na(dec2007_val) && dec2007_val != 0) {
      ((cur_cpi_real - dec2007_val) / dec2007_val) * 100
    } else NA_real_

    pandemic_avg <- .lms_avg(lms, "A3WX", covid3_cpi)
    wages_cpi_total_vs_pandemic <- if (!is.na(cur_cpi_real) && !is.na(pandemic_avg) && pandemic_avg != 0) {
      ((cur_cpi_real - pandemic_avg) / pandemic_avg) * 100
    } else NA_real_
  } else {
    latest_wages_cpi <- latest_regular_cpi <- NA_real_
    wages_cpi_change_q <- wages_cpi_change_y <- wages_cpi_change_covid <- wages_cpi_change_election <- NA_real_
    wages_cpi_total_vs_dec2007 <- wages_cpi_total_vs_pandemic <- NA_real_
  }

  assign("latest_wages_cpi",           latest_wages_cpi,           envir = target_env)
  assign("latest_regular_cpi",         latest_regular_cpi,         envir = target_env)
  assign("wages_cpi_change_q",         wages_cpi_change_q,         envir = target_env)
  assign("wages_cpi_change_y",         wages_cpi_change_y,         envir = target_env)
  assign("wages_cpi_change_covid",     wages_cpi_change_covid,     envir = target_env)
  assign("wages_cpi_change_election",  wages_cpi_change_election,  envir = target_env)
  assign("wages_cpi_total_vs_dec2007", wages_cpi_total_vs_dec2007, envir = target_env)
  assign("wages_cpi_total_vs_pandemic", wages_cpi_total_vs_pandemic, envir = target_env)

  # ---- Vacancies from LMS ----
  if (!is.null(lms)) {
    vac_vals <- .lms_col(lms, "AP2Y")
    vac_ok <- which(!is.na(lms$dates) & !is.na(vac_vals))
    if (length(vac_ok) > 0) {
      vac_latest_date <- lms$dates[vac_ok[length(vac_ok)]]
      if (!is.null(vac_end_override) && vac_end_override %in% lms$dates[vac_ok]) {
        vac_end <- vac_end_override
      } else {
        vac_end <- vac_latest_date
      }
      vac_cur <- .lms_val(lms, "AP2Y", vac_end)
      vac_dq  <- vac_cur - .lms_val(lms, "AP2Y", vac_end %m-% months(3))
      vac_dy  <- vac_cur - .lms_val(lms, "AP2Y", vac_end %m-% months(12))
      vac_dc  <- vac_cur - .lms_val(lms, "AP2Y", COVID_DATE)
      vac_de  <- vac_cur - .lms_val(lms, "AP2Y", ELEC24_DATE)
    } else {
      vac_end <- if (!is.null(vac_end_override)) vac_end_override else lfs_end_cur
      vac_cur <- vac_dq <- vac_dy <- vac_dc <- vac_de <- NA_real_
    }
  } else {
    vac_end <- if (!is.null(vac_end_override)) vac_end_override else lfs_end_cur
    vac_cur <- vac_dq <- vac_dy <- vac_dc <- vac_de <- NA_real_
  }

  assign("vac_cur", vac_cur, envir = target_env)
  assign("vac_dq",  vac_dq,  envir = target_env)
  assign("vac_dy",  vac_dy,  envir = target_env)
  assign("vac_dc",  vac_dc,  envir = target_env)
  assign("vac_de",  vac_de,  envir = target_env)
  assign("vac", list(cur = vac_cur, dq = vac_dq, dy = vac_dy,
                     dc = vac_dc, de = vac_de, end = vac_end), envir = target_env)

  # ---- Days lost from LMS ----
  if (!is.null(lms)) {
    dl_last <- .lms_last_valid(lms, "BBFW")
    days_lost_cur   <- dl_last$value
    days_lost_label <- if (!is.na(dl_last$date)) format(dl_last$date, "%B %Y") else ""
  } else {
    days_lost_cur <- NA_real_
    days_lost_label <- ""
  }

  assign("days_lost_cur",   days_lost_cur,   envir = target_env)
  assign("days_lost_label", days_lost_label, envir = target_env)

  # ---- RTISA payroll (unchanged - separate Excel file) ----
  rtisa_pay <- if (!is.null(file_rtisa)) {
    .read_sheet(file_rtisa, "1. Payrolled employees (UK)")
  } else data.frame()

  rtisa_latest <- anchor_m

  if (nrow(rtisa_pay) > 0 && ncol(rtisa_pay) >= 2) {
    rtisa_text   <- trimws(as.character(rtisa_pay[[1]]))
    rtisa_parsed <- suppressWarnings(lubridate::parse_date_time(rtisa_text, orders = c("B Y", "bY", "BY")))
    rtisa_months <- floor_date(as.Date(rtisa_parsed), "month")
    rtisa_vals   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_pay[[2]]))))

    pay_df <- data.frame(m = rtisa_months, v = rtisa_vals, stringsAsFactors = FALSE)
    pay_df <- pay_df[!is.na(pay_df$m) & !is.na(pay_df$v), ]
    pay_df <- pay_df[order(pay_df$m), ]

    rtisa_latest <- if (!is.null(payroll_end_override) && payroll_end_override %in% pay_df$m) {
      payroll_end_override
    } else if (nrow(pay_df) > 0) {
      pay_df$m[nrow(pay_df)]
    } else anchor_m
    rtisa_cm <- rtisa_latest %m+% months(2)

    months_cur  <- seq(rtisa_cm %m-% months(4), by = "month", length.out = 3)
    months_prev <- seq(rtisa_cm %m-% months(7), by = "month", length.out = 3)
    months_yago <- months_cur %m-% months(12)

    pay_cur_raw <- .avg_by_dates(pay_df$m, pay_df$v, months_cur)
    pay_prev3   <- .avg_by_dates(pay_df$m, pay_df$v, months_prev)
    pay_yago3   <- .avg_by_dates(pay_df$m, pay_df$v, months_yago)

    payroll_cur <- if (!is.na(pay_cur_raw)) pay_cur_raw / 1000 else NA_real_
    payroll_dq  <- if (!is.na(pay_cur_raw) && !is.na(pay_prev3)) (pay_cur_raw - pay_prev3) / 1000 else NA_real_
    payroll_dy  <- if (!is.na(pay_cur_raw) && !is.na(pay_yago3)) (pay_cur_raw - pay_yago3) / 1000 else NA_real_

    covid_base <- .avg_by_dates(pay_df$m, pay_df$v, seq(COVID_DATE  %m-% months(2), by = "month", length.out = 3))
    payroll_dc <- if (!is.na(pay_cur_raw) && !is.na(covid_base)) (pay_cur_raw - covid_base) / 1000 else NA_real_

    elec_base  <- .avg_by_dates(pay_df$m, pay_df$v, seq(ELEC24_DATE %m-% months(2), by = "month", length.out = 3))
    payroll_de <- if (!is.na(payroll_cur) && !is.na(elec_base)) payroll_cur - (elec_base / 1000) else NA_real_

    flash_anchor <- pay_df$m[nrow(pay_df)]
    flash_val    <- .val_by_date(pay_df$m, pay_df$v, flash_anchor)
    flash_prev_m <- .val_by_date(pay_df$m, pay_df$v, flash_anchor %m-% months(1))
    flash_prev_y <- .val_by_date(pay_df$m, pay_df$v, flash_anchor %m-% months(12))
    flash_elec   <- .val_by_date(pay_df$m, pay_df$v, ELEC24_DATE)

    payroll_flash_cur <- if (!is.na(flash_val)) flash_val / 1e6 else NA_real_
    payroll_flash_dm  <- if (!is.na(flash_val) && !is.na(flash_prev_m)) (flash_val - flash_prev_m) / 1000 else NA_real_
    payroll_flash_dy  <- if (!is.na(flash_val) && !is.na(flash_prev_y)) (flash_val - flash_prev_y) / 1000 else NA_real_
    payroll_flash_de  <- if (!is.na(flash_val) && !is.na(flash_elec)) (flash_val - flash_elec) / 1000 else NA_real_
  } else {
    payroll_cur <- payroll_dq <- payroll_dy <- payroll_dc <- payroll_de <- NA_real_
    payroll_flash_cur <- payroll_flash_dm <- payroll_flash_dy <- payroll_flash_de <- NA_real_
    flash_anchor <- anchor_m
  }

  assign("payroll_cur", payroll_cur, envir = target_env)
  assign("payroll_dq",  payroll_dq,  envir = target_env)
  assign("payroll_dy",  payroll_dy,  envir = target_env)
  assign("payroll_dc",  payroll_dc,  envir = target_env)
  assign("payroll_de",  payroll_de,  envir = target_env)
  assign("payroll_flash_cur", payroll_flash_cur, envir = target_env)
  assign("payroll_flash_dm",  payroll_flash_dm,  envir = target_env)
  assign("payroll_flash_dy",  payroll_flash_dy,  envir = target_env)
  assign("payroll_flash_de",  payroll_flash_de,  envir = target_env)

  # ---- RTISA sector payroll (unchanged) ----
  rtisa_sec <- if (!is.null(file_rtisa)) {
    .read_sheet(file_rtisa, "23. Employees (Industry)")
  } else data.frame()

  if (nrow(rtisa_sec) > 0 && ncol(rtisa_sec) >= 18) {
    sec_text   <- trimws(as.character(rtisa_sec[[1]]))
    sec_parsed <- suppressWarnings(lubridate::parse_date_time(sec_text, orders = c("B Y", "bY", "BY")))
    sec_months <- floor_date(as.Date(sec_parsed), "month")

    sec_retail <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[8]]))))
    sec_hosp   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[10]]))))
    sec_health <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[18]]))))

    sec_valid  <- which(!is.na(sec_months) & !is.na(sec_retail))
    sec_anchor <- if (length(sec_valid) > 0) sec_months[sec_valid[length(sec_valid)]] else cm %m-% months(1)

    .sector_full <- function(vals) {
      now   <- .val_by_date(sec_months, vals, sec_anchor)
      prev  <- .val_by_date(sec_months, vals, sec_anchor %m-% months(1))
      yago  <- .val_by_date(sec_months, vals, sec_anchor %m-% months(12))
      covid <- .val_by_date(sec_months, vals, COVID_DATE)
      elec  <- .val_by_date(sec_months, vals, ELEC24_DATE)
      list(
        cur = if (!is.na(now)) now / 1000 else NA_real_,
        dm  = if (!is.na(now) && !is.na(prev)) (now - prev) / 1000 else NA_real_,
        dy  = if (!is.na(now) && !is.na(yago)) (now - yago) / 1000 else NA_real_,
        dc  = if (!is.na(now) && !is.na(covid)) (now - covid) / 1000 else NA_real_,
        de  = if (!is.na(now) && !is.na(elec)) (now - elec) / 1000 else NA_real_
      )
    }

    s_retail <- .sector_full(sec_retail)
    s_hosp   <- .sector_full(sec_hosp)
    s_health <- .sector_full(sec_health)
  } else {
    na_sector <- list(cur = NA_real_, dm = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_)
    s_retail <- s_hosp <- s_health <- na_sector
    sec_anchor <- cm %m-% months(1)
  }

  for (prefix in c("hosp", "retail", "health")) {
    s <- switch(prefix, hosp = s_hosp, retail = s_retail, health = s_health)
    assign(paste0(prefix, "_cur"), s$cur, envir = target_env)
    assign(paste0(prefix, "_dm"),  s$dm,  envir = target_env)
    assign(paste0(prefix, "_dy"),  s$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  s$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  s$de,  envir = target_env)
  }

  # ---- HR1 (unchanged - separate Excel file) ----
  hr1_tbl <- if (!is.null(file_hr1)) .read_sheet(file_hr1, "1a") else data.frame()

  if (nrow(hr1_tbl) > 0 && ncol(hr1_tbl) >= 13) {
    hr1_dates <- .detect_dates(hr1_tbl[[1]])
    hr1_vals  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(hr1_tbl[[13]]))))

    valid_hr1 <- which(!is.na(hr1_dates) & !is.na(hr1_vals))
    if (length(valid_hr1) > 0) {
      last_hr1 <- valid_hr1[length(valid_hr1)]
      hr1_cur  <- hr1_vals[last_hr1]
      hr1_month_label <- format(hr1_dates[last_hr1], "%B %Y")

      prev_hr1 <- if (length(valid_hr1) >= 2) hr1_vals[valid_hr1[length(valid_hr1) - 1]] else NA_real_
      hr1_dm   <- if (!is.na(hr1_cur) && !is.na(prev_hr1)) hr1_cur - prev_hr1 else NA_real_

      hr1_cur_date <- hr1_dates[last_hr1]
      hr1_yago  <- .val_by_date(hr1_dates, hr1_vals, hr1_cur_date %m-% months(12))
      hr1_covid <- .val_by_date(hr1_dates, hr1_vals, COVID_DATE)
      hr1_elec  <- .val_by_date(hr1_dates, hr1_vals, ELEC24_DATE)

      hr1_dy <- if (!is.na(hr1_cur) && !is.na(hr1_yago)) hr1_cur - hr1_yago else NA_real_
      hr1_dc <- if (!is.na(hr1_cur) && !is.na(hr1_covid)) hr1_cur - hr1_covid else NA_real_
      hr1_de <- if (!is.na(hr1_cur) && !is.na(hr1_elec)) hr1_cur - hr1_elec else NA_real_
    } else {
      hr1_cur <- NA_real_
      hr1_dm <- hr1_dy <- hr1_dc <- hr1_de <- NA_real_
      hr1_month_label <- ""
    }
  } else {
    hr1_cur <- NA_real_
    hr1_dm <- hr1_dy <- hr1_dc <- hr1_de <- NA_real_
    hr1_month_label <- ""
  }

  assign("hr1_cur", hr1_cur, envir = target_env)
  assign("hr1_dm",  hr1_dm,  envir = target_env)
  assign("hr1_dy",  hr1_dy,  envir = target_env)
  assign("hr1_dc",  hr1_dc,  envir = target_env)
  assign("hr1_de",  hr1_de,  envir = target_env)
  assign("hr1_month_label", hr1_month_label, envir = target_env)

  # ---- Labels ----
  lfs_period_label       <- lfs_label_narrative(lfs_end_cur)
  lfs_period_short_label <- make_lfs_label(lfs_end_cur)
  vacancies_period_label <- lfs_label_narrative(vac_end)
  vacancies_period_short_label <- make_lfs_label(vac_end)
  payroll_flash_label_val <- format(flash_anchor, "%B %Y")
  sector_month_label <- format(sec_anchor, "%B %Y")

  assign("lfs_period_label",             lfs_period_label,             envir = target_env)
  assign("lfs_period_short_label",       lfs_period_short_label,       envir = target_env)
  assign("vacancies_period_label",       vacancies_period_label,       envir = target_env)
  assign("vacancies_period_short_label", vacancies_period_short_label, envir = target_env)
  assign("payroll_flash_label",          payroll_flash_label_val,      envir = target_env)
  assign("payroll_month_label",          format(rtisa_latest, "%B %Y"),    envir = target_env)
  assign("payroll_period_short_label",  make_lfs_label(rtisa_latest),     envir = target_env)
  assign("hr1_month_label",             hr1_month_label,              envir = target_env)
  assign("sector_month_label",          sector_month_label,           envir = target_env)
  assign("manual_month",                manual_month,                 envir = target_env)

  assign("inact_driver_text", "", envir = target_env)

  assign("payroll", list(
    cur = payroll_cur, dq = payroll_dq, dy = payroll_dy,
    dc = payroll_dc, de = payroll_de,
    flash_cur = payroll_flash_cur, flash_dm = payroll_flash_dm,
    flash_dy = payroll_flash_dy, flash_de = payroll_flash_de,
    flash_anchor = flash_anchor, anchor = anchor_m
  ), envir = target_env)

  assign("wages_nom", list(
    total = list(cur = latest_wages, dq = wages_change_q,
                 dy = wages_change_y, dc = wages_change_covid,
                 de = wages_change_election,
                 public = wages_total_public, private = wages_total_private,
                 qchange = wages_total_qchange),
    regular = list(cur = latest_regular_cash,
                   public = wages_reg_public, private = wages_reg_private,
                   qchange = wages_reg_qchange),
    anchor = anchor_m
  ), envir = target_env)

  invisible(manual_month)
}
