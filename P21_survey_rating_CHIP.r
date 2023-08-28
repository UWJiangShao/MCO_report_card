## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


## Define the custom functions.
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\RetainReliable.R")
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\PercentileRate.R")

## Read in source file
outdir <- "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2022-2023\\3. Survey\\Data\\Temp_Data"
infile <- "CH23_out.xlsx"
CH23_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("GCQ", "HWDC", "PDRat", "HPRat")
CH23_surv <- as.data.frame(read.xlsx(CH23_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   CH23_surv <- merge(CH23_surv, as.data.frame(read.xlsx(CH23_wb), sheet = invar))
# }

## Omit total and LD (if necessary)
CH23_surv <- subset(CH23_surv, !(CH23_surv$PHi_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  CH23_surv[[match(svar, names(CH23_surv))]][which(CH23_surv[match(paste0(svar, "_den"), names(CH23_surv))] < 30)] <- NA
}


## calculate reliability of each score wrt distribution of all scores
CH23_surv <- RetainReliable(CH23_surv, "GCQ", 0.6, varlist = "GCQ")
CH23_surv <- RetainReliable(CH23_surv, "HWDC", 0.6, varlist = "HWDC")
CH23_surv <- RetainReliable(CH23_surv, "PDRat", 0.6, varlist ="PDRat")
CH23_surv <- RetainReliable(CH23_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
CH23_surv <- PercentileRate(CH23_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = CH23_wb, sheetName = "CH23")
# worksheetOrder(CH23_wb) <- c(length(names(CH23_wb)), 1:length(surv_var))
writeData(wb = CH23_wb, sheet = "CH23", x = CH23_surv)
saveWorkbook(CH23_wb, paste(outdir, "CH23_out.xlsx", sep = "/"), overwrite = TRUE)

