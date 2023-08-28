## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


## Define the custom functions.
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\RetainReliable.R")
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\PercentileRate.R")

## Read in source file
outdir <- "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2022-2023\\3. Survey\\Data\\Temp_Data"
infile <- "SP23_out.xlsx"
SP23_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("ATC", "HWDC", "PDRat", "HPRat")
SP23_surv <- as.data.frame(read.xlsx(SP23_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SP23_surv <- merge(SP23_surv, as.data.frame(read.xlsx(SP23_wb), sheet = invar))
# }

## Omit total and LD (if necessary)
SP23_surv <- subset(SP23_surv, !(SP23_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SP23_surv[[match(svar, names(SP23_surv))]][which(SP23_surv[match(paste0(svar, "_den"), names(SP23_surv))] < 30)] <- NA
}

## calculate reliability of each score wrt distribution of all scores
SP23_surv <- RetainReliable(SP23_surv, "ATC", 0.6, varlist = "ATC")
SP23_surv <- RetainReliable(SP23_surv, "HWDC", 0.6, varlist = "HWDC")
SP23_surv <- RetainReliable(SP23_surv, "PDRat", 0.6, varlist ="PDRat")
SP23_surv <- RetainReliable(SP23_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SP23_surv <- PercentileRate(SP23_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SP23_wb, sheetName = "SP23")
# worksheetOrder(SP23_wb) <- c(length(names(SP23_wb)), 1:length(surv_var))
writeData(wb = SP23_wb, sheet = "SP23", x = SP23_surv)
saveWorkbook(SP23_wb, paste(outdir, "SP23_out.xlsx", sep = "/"), overwrite = TRUE)

