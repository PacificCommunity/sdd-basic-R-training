#Load libraries needed to process the primary productions data
library(dplyr)
library(tidyverse)
library(haven) #Reading in Stata files library
library(officer) #Creating a word document report
library(openxlsx) #Creating Excel files
library(pivottabler) #Running pivot or cross tab tables

#You can also create a 'setup' scrip to load all your libraries & functions
source("setup.R")


# ********************* IMPORTING DATAFILE *************************************#

# Local copy of the dataset on your computer

#Setting up path environment
repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path))
setwd(repository)

lforce <- read_dta("data/SPC_REG-13_2011-2019_ILO.dta")

# Data set from OneDrive

#Import the data directly from one drive:
#You need to add a link on your local computer: "add a shortcut to My Files"
#and to create a value with your user's name

#It is better to use the library(foreign) to import the stata database

lforce2 <-read.dta(glue("C:/Users/{user}/OneDrive - SPC/Regional dataset for dotStat/SPC_REG-13_2011-2019_ILO.dta")) 

#with haven library

lforce3 <- read_dta(glue("C:/Users/{user}/OneDrive - SPC/Regional dataset for dotStat/SPC_REG-13_2011-2019_ILO.dta")) %>% 
   mutate(across(where(labelled::is.labelled), haven::as_factor))

# ********************* Define the theme of your table ************************#
# you can find color codes here: https://r-charts.com/colors/
# + spc style here : https://github.com/PacificCommunity/sdd-spcstyle-r-package/tree/main

tableTheme <- list(
  fontName="Helvetica, arial",
  fontSize="0.75em",
  headerBackgroundColor = "#800000",
  headerColor = "#FFF8DC",
  cellBackgroundColor="#FFF8DC",
  cellColor="#800000",
  outlineCellBackgroundColor = "#800000",
  outlineCellColor = "#000000",
  totalBackgroundColor = "grey",
  totalColor="#FFF8DC",
  totalStyle = "bold",
  borderColor = "black"
)

#### ****************** Create our sub-sidiary tables in the form of dataframes ************************* ####
#Create urban Rural dataframe
urbanrural <- data.frame(
  rururb = c(1, 2, 3),
  rururbDesc = c("urban", "Rural", "National")
)

#Create Sex dataframe
sex <- data.frame(
  sex = c(1, 2),
  sexDesc = c("Male", "Female")
)

country <- data.frame(
  country = c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15),
  countryDesc = c("Cook Islands", "Fed. States of Micronesia", "Fiji", "Kiribati", "Marshall Islands", "Nauru", "Niue", "Palau", "Samoa", "Solomon Islands", "Tokelau", "Tonga", "Tuvalu", "Vanuatu", "Wallis & Futuna")
  
)

labforce <- data.frame(
  labforce = c(1, 2, 3),
  labforceDesc = c("1. Employed", "2. Unemployed", "3. Outside Labour Force")
)

age_grp1 <- data.frame(
  age_grp1 = c(1, 2, 3, 4),
  age_grp1Desc = c("15 - 24", "25 - 54", "55 - 64", "65 - 95")
)


labourFource <- merge(lforce, urbanrural, by = "rururb")
labourFource <- merge(labourFource, sex, by = "sex")
labourFource <- merge(labourFource, country, by = "country")
labourFource <- merge(labourFource, labforce, by = "labforce")
labourFource <- merge(labourFource, age_grp1, by = "age_grp1")


#Create Excel workbook
wb <- createWorkbook(creator = Sys.getenv("USERNAME"))


#Create pivot table (1 X 1)
pt <- PivotTable$new()
pt$addData(labourFource)
pt$addColumnDataGroups("age_grp1Desc")
pt$addRowDataGroups("countryDesc")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(indwt), 0)")
pt$theme <- tableTheme
pt$renderPivot()


#Create pivot table (1 X 1) using foreign package
pt <- PivotTable$new()
pt$addData(lforce2)
pt$addColumnDataGroups("age_grp1")
pt$addRowDataGroups("country")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(indwt), 0)")
pt$theme <- tableTheme
pt$renderPivot()


#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-1X1")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-1X1", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)


#Create pivot table (1 X 2)
pt <- PivotTable$new()
pt$addData(labourFource)
pt$addColumnDataGroups("age_grp1Desc")
pt$addColumnDataGroups("sexDesc")
pt$addRowDataGroups("countryDesc")
pt$defineCalculation(calculationName="Population", summariseExpression="comma(sum(indwt), digits=0)") #add comma for thousands
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-1X2")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-1X2", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)


#Create pivot table (2 X 2)
pt <- PivotTable$new()
pt$addData(labourFource)
pt$addColumnDataGroups("age_grp1Desc")
pt$addColumnDataGroups("sexDesc")
pt$addRowDataGroups("countryDesc")
pt$addRowDataGroups("rururbDesc")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(indwt), 0)")
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-2X2")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-2X2", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)


#Create pivot table (2 X 3)
pt <- PivotTable$new()
pt$addData(labourFource)
pt$addColumnDataGroups("age_grp1Desc")
pt$addColumnDataGroups("labforceDesc")
pt$addColumnDataGroups("sexDesc")
pt$addRowDataGroups("countryDesc")
pt$addRowDataGroups("rururbDesc")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(indwt), 0)")
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-2X3")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-2X3", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)



#Save Final Excel Workbook to output folder
saveWorkbook(wb, file="output/Labour_Force_Tables.xlsx", overwrite = TRUE)


#other possibility for cross tabulation without format
ftable(age_grp1 + sex ~ country, lforce2)

#other possibility for table

result2 <- lforce2 %>% filter(age_grp1=="15-24") %>% group_by(country,sex,rururb) %>% 
  mutate(neet=ifelse(neet==2,0,neet),
    perc=round(sum(neet*indwt,na.rm = TRUE)/sum(indwt,na.rm=TRUE)*100,1)) %>% 
  distinct(country,.keep_all = TRUE) %>% 
  select(country,sex,rururb,perc) %>% 
  pivot_wider(names_from=sex, values_from=perc) 
  
library(gt)
  result2 %>% 
  gt(rowname_col = "rururb", groupname_col = "country") %>% 
   #tab_spanner(label = "By sex and age") %>%
  #tab_stubhead(label = "By area") %>%
  tab_style(
    style = cell_text(style = "oblique", weight = "bold"),
    locations = cells_row_groups())  %>%
  tab_style(
    style = cell_text(align = "right", weight = "bold"),
    locations = cells_stub())  %>%
  tab_header(
    title = "NEET by sex & area",
    subtitle = "Pop aged 15-24") %>%
  tab_source_note(md("Source:HIES")) %>% 
    # tab_footnote(
    #   footnote = "This is a footnote.",
    #   locations = cells_body(columns = 1, rows = 1)
    # ) %>% 
    opt_stylize(style = 6, color = "cyan")

  

