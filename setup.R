#Load libraries needed to process the primary productions data
library(dplyr)
library(tidyverse)
library(haven) #Reading in Stata files library
library(officer) #Creating a word document report
library(openxlsx) #Creating Excel files
library(pivottabler) #Running pivot or cross tab tables
library(foreign) #Reading stata files with labels

library(glue)
library(formattable)

library(spcstyle)

`%not in%` <- Negate(`%in%`)

user <- Sys.info()['login']
