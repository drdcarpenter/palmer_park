# process metric workbook to extract data

library(openxlsx)
library(dplyr)
library(janitor)
library(readr)

metric <- "Palmer Park Biodiversity Metric 2.0 Calculation Tool Beta Test - December 2019 Update.xlsm"

m <- loadWorkbook(metric)
m

results <- read.xlsx(m, sheet = 6, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
  row_to_names(row_number = 1, remove_row = TRUE)
results

# create summary of baseline units
baseline <- read.xlsx(m, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
  select(2:5, 7, 9, 13, 16) %>% 
  slice(2:6) %>% 
  row_to_names(row_number = 1, remove_row = TRUE)

baseline$`Area (hectares)` <-  as.numeric(baseline$`Area (hectares)`)
baseline$`Total habitat units` <- as.numeric(baseline$`Total habitat units`)

baseline <- adorn_totals(baseline)



proposed <- read.xlsx(m, sheet = 10, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
  select(1:3, 5, 7, 11, 17) %>% 
  slice(4, 6:7) %>% 
  row_to_names(row_number = 1, remove_row = TRUE) %>% 
  rename(`Strategic Significance` = 6)

proposed$`Area (hectares)` <-  as.numeric(proposed$`Area (hectares)`)
proposed$`Habitat units delivered` <- as.numeric(proposed$`Habitat units delivered`)

proposed <- adorn_totals(proposed)

# as functions?

baseline <- function(metric){
  
  b <- read.xlsx(metric, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(2:5, 7, 9, 13, 16) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 2, remove_row = TRUE)
  
  return(b)
  
}

baseline(m)

proposed <- function(metric){
  
  p <- read.xlsx(metric, sheet = 10, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(1:3, 5, 7, 11, 17) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 4, remove_row = TRUE) %>% 
    rename(`Strategic Significance` = 6) %>% 
    slice(-1)
  
  return(p)
  
}

proposed(m)
