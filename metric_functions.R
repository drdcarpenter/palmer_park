# process metric workbook to extract data

library(openxlsx)
library(dplyr)
library(janitor)
library(readr)

metric <- "Palmer Park Biodiversity Metric 2.0 Calculation Tool Beta Test - December 2019 Update.xlsm"

m <- loadWorkbook(metric)

baseline <- function(metric){
  
  b <- read.xlsx(metric, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(2:5, 7, 9, 13, 16) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 2, remove_row = TRUE)
  
  b$`Area (hectares)` <-  as.numeric(as.character(b$`Area (hectares)`))
  b$`Total habitat units` <- as.numeric(as.character(b$`Total habitat units`))
  
  return(b)
  
}

b <- baseline(m)

proposed <- function(metric){
  
  p <- read.xlsx(metric, sheet = 10, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(1:3, 5, 7, 11, 17) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 4, remove_row = TRUE) %>% 
    rename(`Strategic Significance` = 6) %>% 
    slice(-1)
  
  p$`Area (hectares)` <- as.numeric(as.character(p$`Area (hectares)`))
  p$`Habitat units delivered` <- as.numeric(as.character(p$`Habitat units delivered`))
  
  return(p)
  
}


p <- proposed(m)
