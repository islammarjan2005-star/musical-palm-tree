# auto-detect reference month from database (no manual entry needed)

source("utils/helpers.R")

detected <- auto_detect_manual_month()
if (!is.null(detected)) {
  manual_month <- detected
} else {
  # fallback: derive from current date
  manual_month <- tolower(paste0(format(Sys.Date(), "%b"), format(Sys.Date(), "%Y")))
  message("[config] Could not auto-detect from database; using current date: ", manual_month)
}

# baseline dates used across sheets for change calculations
COVID_DATE  <- as.Date("2020-02-01")
ELEC24_DATE <- as.Date("2024-06-01")

# narrative labels for comparison columns
COVID_LFS_LABEL <- "Dec-Feb 2020"
COVID_VAC_LABEL <- "Jan-Mar 2020"
ELECTION_LABEL  <- "Apr-Jun 2024"
