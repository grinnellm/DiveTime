##### Header #####
# Author:       Matthew H. Grinnell
# Affiliation:  Pacific Biological Station, Fisheries and Oceans Canada (DFO)
# Group:        Herring, Marine Ecosystems and Aquaculture Division
# Address:      3190 Hammond Bay Road, Nanaimo, BC, Canada, V9T 6N7
# Contact:      matthew.grinnell@dfo-mpo.gc.ca | 778.268.1026
# Code name:    CalcDiveTime.R
# Version:      2.0
# Date started: May 08, 2015
# Date edited:  June 05, 2024
#
# Overview:
# Read in a Microsoft Excel file (*.xls, *.xlsx) with Date, Transect, as well as
# start and end times for each diver. Determine daily payable dive time via:
#   1. Minimum of 2 hrs (variable 'minTime') if the diver was active; and
#   2. Round up to the nearest 15 mins (variable 'roundHr') if more than 2 hrs.
#
# Requirements:
# Various packages listed below.
#
# Notes:
# Dive times may be wrong due to Issues #2 and #4;
# see https://github.com/grinnellm/DiveTime/issues

# Issue #2
message(
  "This script does not account for additional dive time when dives\n",
  "\tspawn more than an 8 hr interval in one day (Issue #2)"
)

# Issue #4
message(
  "This script has not been tested for dives that span midnight (Issue #4)"
)

##### Housekeeping #####

# General options
rm(list = ls()) # Clear the workspace
graphics.off() # Turn graphics off

# Install missing packages and load required packages (if required)
UsePackages <- function(
    pkgs, update = FALSE, locn = "http://cran.rstudio.com/"
) {
  # Identify missing (i.e., not yet installed) packages
  newPkgs <- pkgs[!(pkgs %in% installed.packages()[, "Package"])]
  # Install missing packages if required
  if (length(newPkgs)) install.packages(newPkgs, repos = locn)
  # Loop over all packages
  for (i in 1:length(pkgs)) {
    # Load required packages using 'library'
    eval(parse(text = paste("library(", pkgs[i], ")", sep = "")))
  } # End i loop over package names
  # Update packages if requested
  if (update) update.packages(ask = FALSE)
} # End UsePackages function

# Make packages available
UsePackages(pkgs = c("tidyverse", "tools", "scales", "readxl", "xlsx"))

# Keep group function quiet
options(dplyr.summarise.inform = FALSE)

# Note: re-install if weird error message re WithCallingHandlers
# require( "Rcpp" )

##### Controls #####

# Allowable file types
filesOK <- c("xls", "xlsx")

# Choose input file interactively
excelPath <- choose.files(
  caption = "Select excel file with dive times",
  multi = FALSE, filters = matrix(c("MS Excel", paste("*.", filesOK,
    sep = "",
    collapse = ";"
  )), nrow = 1)
)

# Superceeded
warning("`DiveTime` is superseded; use the `BubbleTime` package instead.")

##### Parameters #####

# Rule 1: minimum daily dive time if any diving (mins)
minTime <- 120

# Rule 2: round up to nearest portion of an hour (i.e., 4 = quarter hour)
roundHr <- 4

# Figure resolution: dots per inch
figDPI <- 600

##### Data #####

# Get the file name
excelIn <- basename(excelPath)

# Deconstruct the input file name: extension
inExt <- file_ext(excelIn)

# Deconstruct the input file name: base
inBase <- file_path_sans_ext(excelIn)

# Stop if it's not xlsx
if (!inExt %in% filesOK) {
  stop("Input file must be one of ", filesOK, call. = FALSE)
}

# Get path to output files to same directory as input
outDir <- dirname(excelPath)

# Name of the excel file to export results
excelOut <- paste(inBase, "Output", inExt, sep = ".")

# Load the dive data (i.e., start and end times by person and day)
diveDat <- read_excel(path = file.path(outDir, excelIn), sheet = 1)

##### Main #####

# Grab date and transect columns
dt <- diveDat %>%
  select(Date, Transect)

# Error if NAs
if (any(is.na(dt))) stop("Missing Date or Transect values", call. = FALSE)

# Check for divers with no dive times
diveTimes <- diveDat %>%
  select(-Date, -Transect) %>%
  apply(MARGIN = 2, FUN = function(x) all(is.na(x)))

# Warning message if diver(s) have all NAs
if (any(diveTimes)) {
  stop("Diver(s) have no dive times: ",
    paste(names(diveTimes)[diveTimes], collapse = ", "), call. = FALSE
  )
}

# Format dive data: calculate dive time based on start and end times
rawMins <- diveDat %>%
  # Ensure transects are characters
  mutate(
    Transect = as.character(Transect),
    Date = as.Date(Date)
  ) %>%
  # Put times into one column
  gather(key = Diver, value = Time, -Date, -Transect) %>%
  # Split diver and start/end times
  separate(col = Diver, into = c("Diver", "StEnd")) %>%
  # Remove missing values
  filter(!is.na(Time)) %>%
  # Group by date, transect, diver, and start/end
  group_by(Date, Transect, Diver, StEnd) %>%
  # Get a unique number if there is more than one dive on a given transect
  mutate(Number = 1:n()) %>%
  # Ungroup
  ungroup() %>%
  # Spread time into start and end times
  spread(key = StEnd, value = Time) %>%
  # Group by date, transect, and diver
  group_by(Date, Transect, Diver) %>%
  # Calculate dive times (if there is more than one dive on a transect
  mutate(
    Time = difftime(time1 = End, time2 = Start, units = "mins"),
    Time = as.numeric(Time)
  ) %>%
  # Get the total for the transect
  summarise(Time = sum(Time)) %>%
  # Ungroup
  ungroup() %>%
  # Arrange by date and diver
  arrange(Date, Transect, Diver, Time)

# Stop if any dive times are negative or zero
if (any(rawMins$Time <= 0)) stop("Non-positive dive time(s)", call. = FALSE)

# Calculate dive time in minutes (raw) wide
rawMinsWide <- rawMins %>%
  spread(Diver, Time, fill = 0)

# Function to calculate column totals and bind to last row
CalcTotal <- function(dat) {
  # Grab the non-numeric columns, and add a row
  info <- dat %>%
    mutate(Date = as.character(Date)) %>%
    keep(is.character) %>%
    add_row(Date = "Total")
  # Grab the other columns
  totTime <- dat %>%
    keep(is.numeric) %>%
    colSums() %>%
    t() %>%
    as_tibble()
  # Add the 'total' row
  df <- dat %>%
    keep(is.numeric) %>%
    bind_rows(totTime)
  # Re-combine the info and total times
  res <- bind_cols(info, df)
  # Return the data
  return(res)
} # End CalcTotal function

# Calculate dive time in minutes (raw) with total
rawMinsTot <- CalcTotal(dat = rawMinsWide)

# Sum dive time in minutes by day (raw)
rawMinsDay <- rawMins %>%
  group_by(Date, Diver) %>%
  summarise(Time = sum(Time)) %>%
  ungroup()

# Sum dive time in minutes by day (raw) wide
rawMinsDayWide <- rawMinsDay %>%
  spread(Diver, Time, fill = 0)

# Sum dive time in minutes by day (raw) with total
rawMinsDayTot <- CalcTotal(dat = rawMinsDayWide)

# Adjust time based on the two rules
AdjustTime <- function(x) {
  # If time is less than the minimum, make it the minimum (rule 1)
  adj <- ifelse(x < minTime, minTime, x)
  # Convert to hours
  hrs <- adj / 60
  # Round up to nearest portion of hour (rule 2)
  res <- ceiling(hrs * roundHr) / roundHr
  # Return the adjusted hours
  return(res)
} # End AdjustTime function

# Get dive time (adjusted)
adjHrsDay <- rawMinsDay %>%
  mutate(Time = AdjustTime(Time))

# Get dive time (adjusted) cumulative
adjHrsDayCum <- adjHrsDay %>%
  group_by(Diver) %>%
  mutate(CTime = cumsum(Time)) %>%
  ungroup()

# Get dive time (adjusted) wide
adjHrsDayWide <- adjHrsDay %>%
  spread(Diver, Time, fill = 0)

# Get dive time (adjusted) with total
adjHrsDayTot <- CalcTotal(dat = adjHrsDayWide)

##### Figures #####

# Get plot size for multi-panel
xySize <- ceiling(sqrt(length(unique(rawMins$Date))))

# Plot raw dive time
plotRawMinutes <- ggplot(data = rawMins, aes(x = Diver, y = Time)) +
  geom_bar(stat = "identity", aes(fill = Diver)) +
  geom_hline(yintercept = 120, linetype = "dashed") +
  coord_flip() +
  labs(y = "Time (mins)") +
  scale_y_continuous(breaks = seq(from = 0, to = 1000, by = 60)) +
  theme_bw() +
  theme(legend.position = "none") +
  facet_wrap(~Date, ncol = xySize)

# Save the plot
ggsave(
  plot = plotRawMinutes, filename = file.path(outDir, "RawMinutes.png"),
  height = xySize * 2 + 1, width = xySize * 3 + 1, dpi = figDPI
)

# Plot cumulative dive time
plotCumulativeHrs <- ggplot(
  data = adjHrsDayCum, aes(x = Date, y = CTime, colour = Diver)
) +
  geom_point(size = 4) +
  geom_line(linewidth = 1) +
  scale_x_date(labels = date_format("%Y-%m-%d")) +
  labs(y = paste("Cumulative time (hours)", sep = "")) +
  expand_limits(y = 0) +
  theme_bw() +
  theme(
    legend.key = element_rect(colour = NA), legend.position = "top",
    legend.background = element_rect(colour = "black", fill = "white"),
  )

# Save the plot
ggsave(
  plot = plotCumulativeHrs, filename = "CumulativeHrs.png", height = 6,
  width = 9, dpi = figDPI
)

##### Tables #####

# Create a worksheet
saveWorkbook(
  wb = createWorkbook(type = "xlsx"), file = file.path(outDir, excelOut)
)

# Write the raw times (mins)
write.xlsx(
  x = data.frame(rawMinsTot), file = file.path(outDir, excelOut),
  sheetName = "rawMinsTot", row.names = FALSE
)

# Write the raw times by day (mins)
write.xlsx(
  x = data.frame(rawMinsDayTot), file = file.path(outDir, excelOut),
  sheetName = "rawMinsDayTot", row.names = FALSE, append = TRUE
)

# Write the adjusted times by day (hrs)
write.xlsx(
  x = data.frame(adjHrsDayTot), file = file.path(outDir, excelOut),
  sheetName = "adjHrsDayTot", row.names = FALSE, append = TRUE
)

##### Output #####

# Print total time to screen
print(data.frame(adjHrsDayTot))

##### End #####

# Print end of file message
cat("\nEnd of file 'CalcDiveTime.R'\n")
