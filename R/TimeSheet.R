##Importing the time sheet data###
##################################
library(readxl)
library(stringr)
TS_data.0 <- read_excel('Data/TS_2016_data.xlsx', sheet = 3)

#Extraction of name staff
#########################
NameStaff <- paste(TS_data.0$`Staff Given Name`, TS_data.0$`Staff Family Name`, sep = ' ')

##Extraction days of month
##########################
Days <- na.omit(str_extract(names(TS_data.0), "[0-9]{1,2}"))
ColumnDays <- which(!is.na(str_extract(names(TS_data.0), "[0-9]{1,2}")))

##Extraction all projects
#########################
ProjetAll <- na.omit(str_extract(names(TS_data.0), "^[P|p]roject."))
ColumnProject <- which(!is.na(str_extract(names(TS_data.0), "^[P|p]roject.")))

##Creation table for all programs
#################################
TableProgram <- matrix(0, nrow = length(ProjetAll), ncol = length(Days))
TableProgram <- as.data.frame(TableProgram)

##For one employee
for (i in 1:dim(TableProgram)[1]){
  
  for (j in 1:dim(TableProgram)[2]) {
    TableProgram[i,j] <- ifelse(TS_data.0[22, ColumnDays[j]]=='8', 8, 0)*TS_data.0[22, ColumnProject[i]]/100
    
  }
}



























