#Load libraries needed to process the primary productions data


#You can also create a 'setup' scrip to load all your libraries & functions
source("setup.R")


# ********************* IMPORTING DATAFILE *************************************#

# Local copy of the dataset on your computer

#Setting up path environment
#repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path))
#setwd(repository)

#lforce <- read_dta("data/SPC_TON_2021_HIES_8-Labour_v01_PUF.dta")

# Data set from OneDrive

#Import the data directly from one drive:
#You need to add a link on your local computer: "add a shortcut to My Files"
#and to create a value with your user's name

#Use the library(foreign) to import the stata database and all the labels

lforce <-read.dta(glue("C:/Users/{user}/OneDrive - SPC/NADA/Tonga/SPC_TON_2021_HIES_v01_M_v01_A_PUF/Data/Distribute/SPC_TON_2021_HIES_8-Labour_v01_PUF.dta")) 

#with haven library

lforce3 <- read_dta(glue("C:/Users/{user}/OneDrive - SPC/NADA/Tonga/SPC_TON_2021_HIES_v01_M_v01_A_PUF/Data/Distribute/SPC_TON_2021_HIES_8-Labour_v01_PUF.dta")) %>% 
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
# 

#Create Excel workbook
wb <- createWorkbook(creator = Sys.getenv("USERNAME"))


#Create pivot table (1 X 1)
pt <- PivotTable$new()
pt$addData(lforce)
pt$addColumnDataGroups("age_grp5")
pt$addRowDataGroups("island")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(fweight), 0)")
pt$theme <- tableTheme
pt$renderPivot()


#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-1X1")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-1X1", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)


#Create pivot table (1 X 2)
pt <- PivotTable$new()
pt$addData(lforce)
pt$addColumnDataGroups("age_grp5")
pt$addColumnDataGroups("sex")
pt$addRowDataGroups("island")
pt$defineCalculation(calculationName="Population", summariseExpression="comma(sum(fweight), digits=0)") #add comma for thousands
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-1X2")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-1X2", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)


#Create pivot table (2 X 2)
pt <- PivotTable$new()
pt$addData(lforce)
pt$addColumnDataGroups("age_grp5")
pt$addColumnDataGroups("sex")
pt$addRowDataGroups("island")
pt$addRowDataGroups("rururb")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(fweight), 0)")
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-2X2")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-2X2", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)

#names(lforce)

#Create pivot table (2 X 3)
pt <- PivotTable$new()
pt$addData(lforce)
pt$addColumnDataGroups("age_grp5")
pt$addColumnDataGroups("ilo_lfs")
pt$addColumnDataGroups("sex")
pt$addRowDataGroups("island")
pt$addRowDataGroups("rururb")
pt$defineCalculation(calculationName="Population", summariseExpression="round(sum(fweight), 0)")
pt$theme <- tableTheme
pt$renderPivot()

#Writing final results to Excel worksheet
addWorksheet(wb, "Table1A-2X3")
pt$writeToExcelWorksheet(wb=wb, wsName="Table1A-2X3", 
                         topRowNumber=1, leftMostColumnNumber=1, applyStyles=TRUE, mapStylesFromCSS=TRUE)



#Save Final Excel Workbook to output folder
saveWorkbook(wb, file="output/Tonga_HIES_Tables_example.xlsx", overwrite = TRUE)



#other possibilities for cross tabulation without format
ftable(sex ~ island, lforce)
#Créer un tableau croisé intégrant le poids stat = "freq","prop","rprop",cprop"
library(GDAtools)
wtable(lforce$island,lforce$sex,w=lforce$fweight,stat = "rprop")

#other possibility for table

result2 <- lforce %>% filter(age_grp5 %in% c("15-19 years","20-24 years")) %>% group_by(island,sex,rururb) %>% 
  mutate(ilo_neet=ifelse(is.na(ilo_neet),0,ilo_neet),
      perc=round(sum(ilo_neet*fweight,na.rm = TRUE)/sum(fweight,na.rm=TRUE)*100,1)) %>% 
    distinct(island,.keep_all = TRUE) %>% 
  select(island,sex,rururb,perc) %>% 
  pivot_wider(names_from=sex, values_from=perc) 

  
library(gt)
  result2 %>% 
  gt(rowname_col = "rururb", groupname_col = "island") %>% 
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

  

