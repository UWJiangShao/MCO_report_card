## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


## Define the custom functions.
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\RetainReliable.R")
source("C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\Experiment\\PercentileRate.R")

## Read in source file
outdir <- "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2022-2023\\3. Survey\\Data\\Temp_Data"
infile <- "SK23_out.xlsx"
SK23_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("HPRat", "AtC", "SpecTher", "APM", "coord", "GNI", "transit", "BHCoun")
SK23_surv <- as.data.frame(read.xlsx(SK23_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SK23_surv <- merge(SK23_surv, as.data.frame(read.xlsx(SK23_wb), sheet = invar))
# }

# read in eligible survey population
# SK23_pop <- as.data.frame(read.xlsx(paste(outdir, "SK23_weights.xlsx", sep = "/"), sheet = 5)) %>% select(PHI_Plan_Code = PLAN_CD, sample_pool_members)
# SK23_surv <- merge(SK23_surv, SK23_pop)
# rm(SK23_pop)

## Omit total and LD (if necessary)
SK23_surv <- subset(SK23_surv, !(SK23_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SK23_surv[[match(svar, names(SK23_surv))]][which(SK23_surv[match(paste0(svar, "_den"), names(SK23_surv))] < 30)] <- NA
}

  
## calculate reliability of each score wrt distribution of all scores
SK23_surv <- RetainReliable(SK23_surv, "HPRat", 0.6, varlist = "HPRat")
SK23_surv <- RetainReliable(SK23_surv, "AtC", 0.6, varlist = "AtC")
SK23_surv <- RetainReliable(SK23_surv, "SpecTher", 0.6, varlist = "SpecTher")
SK23_surv <- RetainReliable(SK23_surv, "APM", 0.6, varlist ="APM")
SK23_surv <- RetainReliable(SK23_surv, "coord", 0.6, varlist ="coord")
SK23_surv <- RetainReliable(SK23_surv, "GNI", 0.6, varlist ="GNI")
SK23_surv <- RetainReliable(SK23_surv, "transit", 0.6, varlist ="transit")
SK23_surv <- RetainReliable(SK23_surv, "BHCoun", 0.6, varlist ="BHCoun")


## Rate according to percentile band, adjusted for reliability and significance
SK23_surv <- PercentileRate(SK23_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SK23_wb, sheetName = "SK23")
# worksheetOrder(SK23_wb) <- c(length(names(SK23_wb)), 1:length(surv_var))
writeData(wb = SK23_wb, sheet = "SK23", x = SK23_surv)
saveWorkbook(SK23_wb, paste(outdir, "SK23_out.xlsx", sep = "/"), overwrite = TRUE)
