## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


## Define the custom functions.
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\RetainReliable.R")
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\PercentileRate.R")

## Read in source file
outdir <- "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2022-2023\\3. Survey\\Data\\Temp_Data"
infile <- "SA23_out.xlsx"
SA23_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("AtC", "HWDC", "PDRat", "HPRat")
SA23_surv <- as.data.frame(read.xlsx(SA23_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SA23_surv <- merge(SA23_surv, as.data.frame(read.xlsx(paste(outdir, infile, sep = "/")), sheet = invar))
# }

## Omit total and LD (if necessary)
SA23_surv <- subset(SA23_surv, !(SA23_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SA23_surv[[match(svar, names(SA23_surv))]][which(SA23_surv[match(paste0(svar, "_den"), names(SA23_surv))] < 30)] <- NA
}


## calculate reliability of each score wrt distribution of all scores
SA23_surv <- RetainReliable(SA23_surv, "AtC", 0.6, varlist = "AtC")
SA23_surv <- RetainReliable(SA23_surv, "HWDC", 0.6, varlist = "HWDC")
SA23_surv <- RetainReliable(SA23_surv, "PDRat", 0.6, varlist ="PDRat")
SA23_surv <- RetainReliable(SA23_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SA23_surv <- PercentileRate(SA23_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SA23_wb, sheetName = "SA23")
# worksheetOrder(SA23_wb) <- c(length(names(SA23_wb)), 1:length(surv_var))
writeData(wb = SA23_wb, sheet = "SA23", x = SA23_surv)
saveWorkbook(SA23_wb, paste(outdir, "SA23_out.xlsx", sep = "/"), overwrite = TRUE)