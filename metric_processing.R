# process metric wookbook to extract data

library(openxlsx)
library(dplyr)
library(janitor)
library(readr)

metric <- "Palmer Park Biodiversity Metric 2.0 Calculation Tool Beta Test - December 2019 Update.xlsm"

m <- loadWorkbook(metric)
m

results <- read.xlsx(m, sheet = 6, skipEmptyRows = TRUE, skipEmptyCols = TRUE)
results

# create summary of baseline units
baseline <- read.xlsx(m, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
  select(2:5, 7, 9, 13, 16) %>% 
  slice(2:6) %>% 
  row_to_names(row_number = 1, remove_row = TRUE) %>% 
  parse_number(`Total habitat units`)

baseline$`Area (hectares)` <-  as.numeric(baseline$`Area (hectares)`)
baseline$`Total habitat units` <- as.numeric(baseline$`Total habitat units`)

baseline <- adorn_totals(baseline)



proposed <- read.xlsx(m, sheet = 10, skipEmptyRows = TRUE, skipEmptyCols = TRUE)
proposed
