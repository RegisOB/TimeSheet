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
    TableProgram[i,j] <- ifelse(TS_data.0[1, ColumnDays[j]]=='8', 8, 0)*TS_data.0[1, ColumnProject[i]]/100
    colnames(TableProgram) <- as.character(1:length(Days))
  }
}

##Modified names of columns/rows
colnames(TableProgram) <- as.character(1:length(Days))
row.names(TableProgram) <- names(TS_data.0)[ColumnProject]

##Added column called total
SumCol <- as.data.frame(apply(TableProgram, 1, sum))
names(SumCol) <- 'Total'
TableProgram.1 <- cbind(TableProgram, SumCol)

##Added row called total
SumRow <- as.data.frame(apply(TableProgram.1, 2, sum))
names(SumRow) <- 'Total'
TableProgram.2 <- rbind(TableProgram.1, t(SumRow))

##Added row empty
TableEmpty1 <- matrix('', 1, ncol = length(Days)+1)
TableEmpty1 <- as.data.frame(TableEmpty1)
colnames(TableEmpty1) <- c(as.character(1:length(Days)), 'Total')
row.names(TableEmpty1) <- ' '


###Creation table for Week(W), Annual leave (A), Public hollidays (P), Illness (I), Other absence (O)
#####################################################################################################
TableAbsence <- matrix(0, nrow = 5, ncol = length(Days))
TableAbsence <- as.data.frame(TableAbsence)

row.names(TableAbsence) <- c('Weekends', 'Annual leave', 'Public holidays', 'Illness', 'Other absence')
colnames(TableAbsence) <- as.character(1:length(Days))

##For one employee
List.absence <- c('W','A','P','I','O')
for (i in 1:dim(TableAbsence)[1]){
  
  for (j in 1:dim(TableAbsence)[2]) {
    TableAbsence[i,j] <- ifelse(TS_data.0[1, ColumnDays[j]]==List.absence[i], 8, 0)
  }
}

##Added column called total
SumCol <- as.data.frame(apply(TableAbsence, 1, sum))
names(SumCol) <- 'Total'
TableAbsence.1 <- cbind(TableAbsence, SumCol)

##Added row called total
SumRow <- as.data.frame(apply(TableAbsence.1, 2, sum))
names(SumRow) <- 'Total'
TableAbsence.2 <- rbind(TableAbsence.1, t(SumRow))

##Total productive hours
TotProd_Hr <- TableProgram.2['Total', ]
row.names(TotProd_Hr) <- 'Total productive hours'

##Total hours
#Total absence
Tot_Abs <- TableAbsence.2['Total', ]
Tot_Hrs <- rbind(TotProd_Hr, Tot_Abs)

SumRow <- as.data.frame(apply(Tot_Hrs, 2, sum))
names(SumRow) <- 'Total hours'
Tot_Hrs.1 <- rbind(Tot_Hrs, t(SumRow))
Tot_Hrs.2 <- Tot_Hrs.1['Total hours', ]













