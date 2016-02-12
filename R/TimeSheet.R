##Importing the time sheet data###
##################################
library(readxl)
TS_data.0 <- read_excel('Data/TS_2016_data.xlsx', sheet = 3)

#Extraction of name staff
#########################
NameStaff <- paste(TS_data.0$`Staff Given Name`, TS_data.0$`Staff Family Name`, sep = ' ')

