# process metric workbook to extract data

library(openxlsx)
library(dplyr)
library(janitor)
library(readr)

metric <- "Palmer Park Biodiversity Metric 2.0 Calculation Tool Beta Test - December 2019 Update.xlsm"

m <- loadWorkbook(metric)


# results table
results <- read.xlsx(m, sheet = 6, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
  row_to_names(row_number = 1, remove_row = TRUE) %>% 
  clean_names() %>% 
  rename(`Headline Results` = headline_results," " = na, Units = na_2)

print(results, na.print = "")

# calculate baseline

baseline <- function(metric){
  
  b <- read.xlsx(metric, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(3:5, 7, 9, 13, 16) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 2, remove_row = TRUE) %>% 
    rename(Habitat = 1, Area = 2, Connectivity = 5, `Biodiversity units` = 7) %>% 
    slice(1:4)
  
  b$Area <-  as.numeric(as.character(b$Area))
  b$`Biodiversity units` <- as.numeric(as.character(b$`Biodiversity units`))
  
  b <- adorn_totals(b)
  
  return(b)
  
}

b <- baseline(m)

# create summary table of proposed units
# this only does the created habitats, not the retained or enhanced habitats

proposed <- function(metric){
  
  p <- read.xlsx(metric, sheet = 10, skipEmptyRows = TRUE, skipEmptyCols = TRUE) %>% 
    select(1:3, 5, 7, 11, 17) %>% 
    remove_empty(which = c("rows", "cols")) %>% 
    row_to_names(row_number = 4, remove_row = TRUE) %>% 
    rename(Habitat = 1, Area = 2, Connectivity = 5, `Strategic significance` = 6, `Biodiversity units` = 7) %>% 
    slice(2:3)
  
  p$Area <- as.numeric(as.character(p$Area))
  p$`Biodiversity units` <- as.numeric(as.character(p$`Biodiversity units`))
  
  return(p)
  
}


p <- proposed(m)



# retained habitats

r <- read.xlsx(m, sheet = 9, startRow = 4, skipEmptyRows = TRUE, skipEmptyCols = TRUE)
retained <- r %>% 
  select(3,17,5,7,10,13,20) %>% 
  slice(2:5) %>% 
  row_to_names(row_number = 1, remove_row = TRUE) %>% 
  rename(Habitat = 1, Area = 2, `Biodiversity units` = 7)

retained$Area <- as.numeric(as.character(retained$Area))
retained$`Biodiversity units` <- as.numeric(as.character(retained$`Biodiversity units`))



# proposed units table

proposed <- bind_rows(retained, p) %>% 
  slice(-6) %>% 
  adorn_totals()

