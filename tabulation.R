#Load libraries needed to process the primary productions data
library(dplyr)
library(tidyverse)
library(haven) #Reading in Stata files library
library(officer) #Creating a word document report
library(openxlsx) #Creating Excel files
library(pivottabler) #Running pivot or cross tab tables

#Setting up path environment
repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path))
setwd(repository)

lforce <- read_dta("data/SPC_REG-13_2011-2019_ILO.dta")

tableTheme <- list(
  fontName="Helvetica, arial",
  fontSize="0.75em",
  headerBackgroundColor = "#800000",
  headerColor = "#FFF8DC",
  cellBackgroundColor="#FFF8DC",
  cellColor="#800000",
  outlineCellBackgroundColor = "#800000",
  outlineCellColor = "#000000",
  totalBackgroundColor = "#FFF8DC",
  totalColor="#800000",
  totalStyle = "bold",
  borderColor = "#457c4b"
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
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(indwt), 0)")
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


