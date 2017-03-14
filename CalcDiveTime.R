###############################################################################
# 
# Author:       Matthew H. Grinnell
# Affiliation:  Pacific Biological Station, Fisheries and Oceans Canada (DFO) 
# Group:        Shellfish, Marine Ecosystems and Aquaculture Division
# Address:      3190 Hammond Bay Road, Nanaimo, BC, Canada, V9T 6N7
# Contact:      e-mail: matt.grinnell@dfo-mpo.gc.ca | tel: 250.756.7055
# Project:      Sandbox
# Code name:    CalcDiveTime.R
# Version:      1.1
# Date started: May 08, 2015
# Date edited:  May 08, 2016
#
# Overview: 
# Read in a Microsoft excel file (*.xlsx) with columns "Date" in the format 
# "August 14, 2015", as well as start and end times for each dive formatted as 
# "15:12" for each diver, for example "MG.St", and "MG.End". See the example 
# dive times in "Example.xlsx." This data must be in the first worksheet in the 
# excel file. If there are multiple excel sheets with dive times (i.e., for two 
# different surveys), put each one in a separate directory. Determine daily 
# payable dive time via: 
#   1. Minimum of 2 hrs (variable 'minTime') if the diver was active; and 
#   2. Round up to the nearest 15 mins (variable 'roundHr') if more than 2 hrs.
# 
# Requirements: 
# Various packages listed below.
# 
# Notes: 
# This script has been tested on R 3.3.2 using Windows 7; it may or may not 
# work on other platforms. I suspect that time may be wrong for dives that go
# past midnight, but this has not been tested.
#
###############################################################################


########################
##### Housekeeping #####
########################

# General options
rm( list=ls( ) )      # Clear the workspace
sTime <- Sys.time( )  # Start the timer
graphics.off( )       # Turn graphics off

# Install missing packages and load required packages (if required)
UsePackages <- function( pkgs, update=FALSE, locn="http://cran.rstudio.com/" ) {
  # Identify missing (i.e., not yet installed) packages
  newPkgs <- pkgs[!(pkgs %in% installed.packages( )[, "Package"])]
  # Install missing packages if required
  if( length(newPkgs) )  install.packages( newPkgs, repos=locn )
  # Loop over all packages
  for( i in 1:length(pkgs) ) {
    # Load required packages using 'library'
    eval( parse(text=paste("library(", pkgs[i], ")", sep="")) )
  }  # End i loop over package names
  # Update packages if requested
  if( update ) update.packages( ask=FALSE )
}  # End UsePackages function

# Make packages available
UsePackages( pkgs=c("rJava", "plyr", "xlsx", "reshape2", "ggplot2", "lubridate", 
        "tools", "scales") )

## Note: Re-install if weird error message re WithCallingHandlers
#require( "Rcpp" )  


#################### 
##### Controls ##### 
####################     

# Choose input file interactively
excelPath <- choose.files( caption="Select excel file with dive times", 
    multi=FALSE, filters=matrix(c("MS Excel", "*.xlsx"), nrow=1) )


######################
##### Parameters #####
######################

# Rule 1: minimum dive time if any diving (mins)
minTime <- 120

# Rule 2: round up to nearest portion of an hour (i.e., 4 = quarter hour)
roundHr <- 4


################
##### Data #####
################

# Get the file name
excelIn <- basename( excelPath )

# Stop if it's not xlsx
if( file_ext(excelIn) != "xlsx" )  
  stop( "Input file must be xlsx", call.=FALSE )

# Get path to output files to same directory as input
outDir <- dirname( excelPath )

# Deconstruct the input file name
inputSplit <- strsplit( excelIn, split=".", fixed=TRUE )[[1]]

# Name of the excel file to export results
excelOut <- paste( inputSplit[1], "Output", inputSplit[2], sep="." )

# Load the dive data (i.e., start and end times by person and day)
diveDat <- read.xlsx( file=file.path(outDir, excelIn), sheetIndex=1 )


################ 
##### Main ##### 
################     

# Format dive data: calculate dive time based on start and end times
CalcRawTime <- function( dat ) {
  # Get columns names
  colNames <- colnames( diveDat )
  # Ensure that Date and Transect columns exist
  if( !all(c("Date", "Transect") %in% colNames) )
    stop( "Excel file requires columns named 'Date' and 'Transect'", 
        call.=FALSE )
  # Get columns with non-time info
  info <- data.frame( Date=dat$Date, Transect=dat$Transect )
  # Ensure date is date
  info$Date <- as.Date( info$Date )
  # Ensure transect is a factor
  info$Transect <- as.factor( info$Transect )
  # Check for NAs, error
  if( any(is.na(info)) )  stop( "Missing Date or Transect values", call.=FALSE )
  # Get columns with time
  times <- dat[, colnames(dat) != c("Date", "Transect")]
  # Order by diver
  times <- times[, order(colnames(times))]
  # Get indices for start times
  stCols <- grep( pattern=".St", x=colnames(times) )
  # Get indices for end times
  endCols <- grep( pattern=".End", x=colnames(times) )
  # Start an empty matrix to hold results
  timeMat <- matrix( NA, nrow=nrow(info), ncol=length(stCols) )
  # Start an empty vector of names
  divers <- 0
  # For each start/end pair, calculate time
  for( i in 1:length(stCols) ) {
    # Get the diver name (start)
    iNameSt <- gsub( pattern=".St", replacement="", 
        x=colnames(times[stCols[i]]) )
    # Get the dive name (end)
    iNameEnd <- gsub( pattern=".End", replacement="", 
        x=colnames(times[endCols[i]]) )
    # Error if names don't match
    if( iNameSt != iNameEnd ) 
      stop( "Diver name mismatch: ", iNameSt, " vs. ", iNameEnd, call.=FALSE )
    # Start times
    iStart <- times[, stCols[i]]
    # Stop if start times are not actually times
    if( !"POSIXct" %in% class(iStart) ) 
      stop( "Non-time data in start times for ", iNameEnd, call.=FALSE )
    # Update start times with correct day
    iStart <- update( iStart, year=year(info$Date), month=month(info$Date),
        mday=mday(info$Date) )
    # End times
    iEnd <- times[, endCols[i]]
    # Stop if end times are not actually times
    if( !"POSIXct" %in% class(iEnd) ) 
      stop( "Non-time data in end times for ", iNameEnd, call.=FALSE )
    # Update end times with correct day
    iEnd <- update( iEnd, year=year(info$Date), month=month(info$Date),
        mday=mday(info$Date) )
    # Check NAs: one time is NA, and not the other
    oneNA <- xor( x=is.na(iStart), y=is.na(iEnd) )
    # Stop if one is NA
    if( any(oneNA) )  stop( "Missing dive time for ", iNameEnd, call.=FALSE )
    # Calculate difference
    iDiff <- difftime( time1=iEnd, time2=iStart, units="mins" )
    # Stop if any dive times are negative or zero
    if( any(na.omit(iDiff) <= 0) ) 
      stop( "Non-positive dive time(s) for ", iNameSt, call.=FALSE )
    # Get diver name
    divers[i] <- iNameSt
    # Add the time info
    timeMat[,i] <- iDiff
  }  # End i loop over start times 
  # Name the matrix of dive times
  colnames(timeMat) <- divers
  # Count number of divers per transect
  buddies <- apply( X=timeMat, MARGIN=1, FUN=function(x) 
        length(na.omit(x)) )
  # Check if each transect had at least two divers
  noBuddy <- which( buddies==1 )
  # Warning for no dive buddy
  if( length(noBuddy) >= 1 ) 
    warning( "No dive buddy on transect(s): ", 
        paste(dat$Transect[noBuddy], collapse=", "), call.=FALSE )
  # Add date info to time mat
  timeMat <- data.frame( info, timeMat )
  # Replace NAs with zeros
  timeMat[is.na(timeMat)] <- 0
  # Get time by transect and date
  timeMatGrp <- aggregate( x=timeMat[-c(1, 2)], by=list(Date=timeMat$Date, 
          Transect=timeMat$Transect), FUN=sum )
  # Order by date and transect
  tmgOrd <- timeMatGrp[order(timeMatGrp$Date, timeMatGrp$Transect), ]
  # Remove transect
  res <- subset( x=tmgOrd, select=-c(Transect) )
  # Return the date, and raw dive times in minutes
  return( res )
}  # End CalcRawTime function

# Calculate dive time in minutes (raw)
rawMins <- CalcRawTime( dat=diveDat )

# Calculate average time per transect, by day
minsPerTrans <- aggregate( x=rawMins[-1], by=list(Date=rawMins$Date),
    FUN=function(x)  mean(x[x>0]) )

# Calculate number of transects per day
transPerDay <- aggregate( x=rawMins[-1], by=list(Date=rawMins$Date), 
    FUN=function(x)  length(x[x>0]) )

# Sum dive time in minutes by day (raw)
rawMinsDay <- aggregate( x=rawMins[-1], by=list(Date=rawMins$Date), FUN=sum )

# Sum dive time by day, and adjust as per rules
AdjustTime <- function( dat ) {
  # Get columns with non-time info
  info <- data.frame( Date=dat$Date )
  # Get columns with time
  times <- dat[, colnames(dat) != "Date"]
  # Start loop over rows
  for( i in 1:nrow(times) ) {
    # Start loop over columns
    for( j in 1:ncol(times) ) {
      # Adjust time by day: minimum 2 hours if any diving (rule 1)
      if( times[i, j] > 0 )  times[i, j] <- max( minTime, times[i, j] )
    }  # End j loop over columns
  }  # End i loop over rows
  # Convert to hours
  times <- times / 60
  # Round up to nearest portion of hour (rule 2)
  times <- ceiling( times*roundHr ) / roundHr
  # Return the data
  return( cbind(info, times) )
}  # End AdjustTime function

# Get dive time (adjusted)
adjHrsDay <- AdjustTime( dat=rawMinsDay )

# Calculate totals for times
CalcTotalTime <- function( dat ) {
  # Get columns with time
  times <- dat[, colnames(dat) != "Date"]
  # Change date to character
  dat$Date <- as.character( dat$Date )
  # Column totals
  totTimes <- colSums( times )
  # Add an empty row
  dat <- rbind( dat, rep(NA, times=ncol(dat)) )
  # Get bottom row index
  lastRow <- nrow(  dat )
  # Insert total
  dat$Date[lastRow] <- "Total"
  # Insert totals
  dat[lastRow, 2:ncol(dat)] <- totTimes
  # Return the updated data, with totals
  return( dat )
}  # End CalcTotalTime function

# Raw times (minutes) with total
rawMinsTot <- CalcTotalTime( dat=rawMins )

# Raw times by day (minutes) with total
rawMinsDayTot <- CalcTotalTime( dat=rawMinsDay )

# Adjusted time by day (hours) with total
adjHrsDayTot <- CalcTotalTime( dat=adjHrsDay )


###################
##### Figures #####
###################

# Plot raw dive time
PlotMins <- function( dat ) {
  # Reshape to long
  lDat <- melt( dat, id.vars=c("Date"), value.name="Time" )
  # Rename divers
  colnames(lDat)[colnames(lDat) == "variable"] <- "Diver"
  # Remove zeros
  lDat <- lDat[lDat$Time > 0, ]
  # Determine size
  xySize <- ceiling( sqrt(length(unique(lDat$Date))) )
  # Plot
  plt <- ggplot( data=lDat, aes(x=Diver, y=Time) ) +
#     geom_text( aes(label=Time), vjust=0 ) + 
      geom_bar( stat="identity", aes(fill=Diver) ) +
      geom_hline( yintercept=120, linetype="dashed" ) + 
      coord_flip( ) +
      labs( y="Time (mins)") +
      theme_bw( ) +
      theme( legend.position="none" ) +
      facet_wrap( ~ Date ) +
      ggsave( filename=file.path(outDir, "RawMinutes.pdf"), height=xySize*2+1, 
          width=xySize*3+1 )
}  # End PlotMins function

# Plot dive time (raw minutes)
PlotMins( dat=rawMinsDay )

# Plot time per transect
PlotMinsPerTrans <- function( dat ) {
  # Reshape to long
  lDat <- melt( dat, id.vars=c("Date"), value.name="Time" )
  # Rename divers
  colnames(lDat)[colnames(lDat) == "variable"] <- "Diver"
  # Determine size
  xySize <- ceiling( sqrt(length(unique(lDat$Date))) )
  # Plot
  plt <- ggplot( data=lDat, aes(x=Date, y=Time) ) +
      geom_density( stat="identity", fill="lightgrey" ) +
      labs( y="Average daily time per transect (mins)") +
      theme_bw( ) +
      scale_x_date( labels=date_format("%Y-%m-%d") ) +
      theme( legend.position="none", 
          axis.text.x=element_text(angle=22.5, hjust=1) ) +
      facet_wrap( ~ Diver ) +
      ggsave( filename=file.path(outDir, "MinsPerTrans.pdf"), height=xySize*2+1,
          width=xySize*3+1 )
}  # End PlotMinsPerTrans function

# Plot minutes per transect
PlotMinsPerTrans( dat=minsPerTrans )

# Plot number of transects per day
PlotTransPerDay <- function( dat ) {
  # Reshape to long
  lDat <- melt( dat, id.vars=c("Date"), value.name="Number" )
  # Rename divers
  colnames(lDat)[colnames(lDat) == "variable"] <- "Diver"
  # Determine size
  xySize <- ceiling( sqrt(length(unique(lDat$Date))) )
  # Plot
  plt <- ggplot( data=lDat, aes(x=Date, y=Number) ) +
      geom_density( stat="identity", fill="lightgrey" ) +
      labs( y="Number of transects per day") +
      theme_bw( ) +
      scale_x_date( labels=date_format("%Y-%m-%d") ) +
      theme( legend.position="none", 
          axis.text.x=element_text(angle=22.5, hjust=1) ) +
      facet_wrap( ~ Diver ) +
      ggsave( filename=file.path(outDir, "TransPerDay.pdf"), height=xySize*2+1,
          width=xySize*3+1 )
}  # End PlotTransPerDay function

# Plot number of transects per day
PlotTransPerDay( dat=transPerDay )

# Plot cumulative dive time
PlotCumulativeTime <- function( dat, dTime, fName ) {
  # Get cumulative time
  cDat <- cbind( dat[1], cumsum(dat[-1]) )
  # Reshape to long
  lDat <- melt( cDat, id.vars=c("Date"), value.name="Time" )
  # Rename divers
  colnames(lDat)[colnames(lDat) == "variable"] <- "Diver"
  # Plot
  plt <- ggplot( data=lDat, aes(x=Date, y=Time) ) +
      geom_point( aes(colour=Diver) ) + 
      geom_line( aes(colour=Diver) ) +
      scale_x_date( labels=date_format("%Y-%m-%d") ) +
      labs( y=paste("Cumulative time (", dTime, ")", sep="" ) ) +
      theme_bw( ) +
      theme( legend.position=c(0, 1), legend.justification=c(0, 1),
          legend.background=element_rect(colour="black", fill="white"),
          legend.key=element_rect(colour=NA) ) +
      ggsave( filename=file.path(outDir, fName), height=6, width=9 )
}  # End PlotCumulativeTime function

# Plot cumulative time (raw minutes)
PlotCumulativeTime( dat=rawMinsDay, dTime="mins", fName="CumulativeMins.pdf" )

# Plot cumulative time (adjusted hours)
PlotCumulativeTime( dat=adjHrsDay, dTime="hrs", fName="CumulativeHrs.pdf" )


##################
##### Tables #####
##################

# Create a worksheet
saveWorkbook( wb=createWorkbook(type="xlsx"), file=file.path(outDir, excelOut) )

# Write the raw times (mins)
write.xlsx( x=rawMinsTot, file=file.path(outDir, excelOut), 
    sheetName="rawMinsTot", row.names=FALSE )

# Write the raw times by day (mins)
write.xlsx( x=rawMinsDayTot, file=file.path(outDir, excelOut), 
    sheetName="rawMinsDayTot", row.names=FALSE, append=TRUE )

# Write the adjusted times by day (hrs)
write.xlsx( x=adjHrsDayTot, file=file.path(outDir, excelOut), 
    sheetName="adjHrsDayTot", row.names=FALSE, append=TRUE )


##################
##### Output #####
##################

# Print total time to screen
print( adjHrsDayTot )


############### 
##### End ##### 
############### 

# Print end of file message
cat( "\nEnd of file 'CalcDiveTime.R'\n" )
