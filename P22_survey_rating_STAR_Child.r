## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


## Define the custom functions.
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\RetainReliable.R")
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\PercentileRate.R")

## Read in source file
outdir <- "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2022-2023\\3. Survey\\Data\\Temp_Data"
infile <- "SC23_out.xlsx"
SC23_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("GCQ", "HWDC", "PDRat", "HPRat")
SC23_surv <- as.data.frame(read.xlsx(SC23_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SC23_surv <- merge(SC23_surv, as.data.frame(read.xlsx(SC23_wb), sheet = invar))
# }

## Omit total and LD (if necessary)
SC23_surv <- subset(SC23_surv, !(SC23_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SC23_surv[[match(svar, names(SC23_surv))]][which(SC23_surv[match(paste0(svar, "_den"), names(SC23_surv))] < 30)] <- NA
}


## calculate reliability of each score wrt distribution of all scores
SC23_surv <- RetainReliable(SC23_surv, "GCQ", 0.6, varlist = "GCQ")
SC23_surv <- RetainReliable(SC23_surv, "HWDC", 0.6, varlist = "HWDC")
SC23_surv <- RetainReliable(SC23_surv, "PDRat", 0.6, varlist ="PDRat")
SC23_surv <- RetainReliable(SC23_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SC23_surv <- PercentileRate(SC23_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SC23_wb, sheetName = "SC23")
# worksheetOrder(SC23_wb) <- c(length(names(SC23_wb)), 1:length(surv_var))
writeData(wb = SC23_wb, sheet = "SC23", x = SC23_surv)
saveWorkbook(SC23_wb, paste(outdir, "SC23_out.xlsx", sep = "/"), overwrite = TRUE)