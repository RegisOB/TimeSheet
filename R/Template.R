###Reading R function to generate automatically the timesheet for all employees
###############################################################################

TimeSheet <- function(Year=2016, Month=1, IndexSheet=3){
  
  #Defining interactivly the directory of main folder
  ###################################################
  directory <- choose.dir(default = "", caption = "Choisir le Dossier principal")
  setwd(directory)
  
  
  ##Importing the time sheet data###
  ##################################
  library(readxl)
  library(stringr)
  library(chron)
  TS_data.0 <- read_excel('Data/TS_2016_data.xlsx', sheet = IndexSheet)
  NameSup <-  read_excel('Data/TS_2016_data.xlsx', sheet = 2)
  
  ##Fill name of supervisor
  TS_data.0$Supervisor <- NA
  for (i in 1:nrow(TS_data.0)){
    for (j in 1:nrow(NameSup)){
      if (TS_data.0$Project[i]==NameSup$Project[j]){
        TS_data.0$Supervisor[i] <- NameSup$`Project supervisor`[j]
      }
    }
  }
  
 
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
  
  ##Management of dates
  date0 <- as.POSIXlt(paste(Year, Month, 1, sep = '-'), tz = "UTC")
  date1 <- as.POSIXlt(paste(Year, Month, length(Days), sep = '-'), tz = "UTC")
  MonthDay <- seq.POSIXt(date0, date1, 'day')
  list.week <- which(is.weekend(MonthDay))
  
  dataPubday <- read_excel('Data/TS_2016_data.xlsx', sheet = 1)
  list.pubday <- which(is.holiday(dates(as.character(MonthDay), format= 'y-m-d'), 
                                  dates(as.character(dataPubday$Date),format= 'y-m-d')))
  
  list.final <- union(list.week, list.pubday)
  
  
  ##For all employees
  ###################
  List.timesheet <- vector('list', nrow(TS_data.0))
  
  for (k in 1:nrow(TS_data.0)){
  
  ##Creation table for all programs
  #################################
  TableProgram <- matrix(0, nrow = length(ProjetAll), ncol = length(Days))
  TableProgram <- as.data.frame(TableProgram)
  
  ##For one employee
  for (i in 1:dim(TableProgram)[1]){
    
    for (j in setdiff(1:dim(TableProgram)[2], list.final)) {
      TableProgram[i,j] <- ifelse(TS_data.0[k, ColumnDays[j]]=='8', 8, 0)*TS_data.0[k, ColumnProject[i]]/100
      #TableProgram[i,j] <- ifelse(TS_data.0[k, ColumnDays[j]]=='8', 8, 0)
      colnames(TableProgram) <- as.character(1:length(Days))
    }
  }
  
  ##Modified names of columns/rows
  colnames(TableProgram) <- as.character(1:length(Days))
  row.names(TableProgram) <- names(TS_data.0)[ColumnProject]
  row.names(TableProgram)[1] <- TS_data.0$Project[k]
  
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
  
  TableEmpty2 <- matrix('', 1, ncol = length(Days)+1)
  TableEmpty2 <- as.data.frame(TableEmpty2)
  colnames(TableEmpty2) <- c(as.character(1:length(Days)), 'Total')
  row.names(TableEmpty2) <- '  '
  
  TableEmpty3 <- matrix('', 1, ncol = length(Days)+1)
  TableEmpty3 <- as.data.frame(TableEmpty3)
  colnames(TableEmpty3) <- c(as.character(1:length(Days)), 'Total')
  row.names(TableEmpty3) <- '   '
  
  
  
  ###Creation table for Week(W), Annual leave (A), Public hollidays (P), Illness (I), Other absence (O)
  #####################################################################################################
  TableAbsence <- matrix(0, nrow = 5, ncol = length(Days))
  TableAbsence <- as.data.frame(TableAbsence)
  
  row.names(TableAbsence) <- c('Weekends', 'Annual leave', 'Public holidays', 'Illness', 'Other absence')
  #row.names(TableAbsence) <- c('Annual leave', 'Illness', 'Other absence')
  
  colnames(TableAbsence) <- as.character(1:length(Days))
  
  ##For one employee
  List.absence <- c('W','A','P','I','O')
  #List.absence <- c('A','I','O')
  for (i in c(2,4,5)){
    
    for (j in 1:dim(TableAbsence)[2]) {
      TableAbsence[i,j] <- ifelse(TS_data.0[k, ColumnDays[j]]==List.absence[i], 8, 0)
    }
  }
  
  ##Fill the weekend and public days
  for (j in list.week) {
    TableAbsence[1,j] <- ifelse(TS_data.0[k, ColumnDays[j]]=='8', 8, 0)
  }
  
  for (j in list.pubday) {
    TableAbsence[3,j] <- ifelse(TS_data.0[k, ColumnDays[j]]=='8', 8, 0)
  }

  
  
  ##Added column called total
  SumCol <- as.data.frame(apply(TableAbsence, 1, sum))
  names(SumCol) <- 'Total'
  TableAbsence.1 <- cbind(TableAbsence, SumCol)
  
  ##Added row called total
  SumRow <- as.data.frame(apply(TableAbsence.1, 2, sum))
  names(SumRow) <- 'Total '
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
  
  
  ###Merge all the tables
  timeshee_table <- rbind.data.frame(TableProgram.2, TableEmpty1, TableAbsence.2, TableEmpty2,
                                     TotProd_Hr, TableEmpty3, Tot_Hrs.2)
  
  List.timesheet[[k]] <- timeshee_table
  
  }
  ###############################################################################
  
  
  #Generating the template
  #########################
  library(knitr)
  setwd('./TimeSheetTemplate')
  
  ##Generate all TimeSheet_Name.rnw
  for (i in 1:length(unique(NameStaff))){
    knit("Child_Template.Rnw",
             output=paste('TimeSheet_', NameStaff[i], '.Rnw', sep=""))}

  #Generate the final time sheet template all emplyees in one pdf
  knit2pdf('Main_Template.Rnw')
  
  #Rename file template
  file.rename('Main_Template.pdf', paste('TIMESHEET_CERMEL_', Sys.Date(), '.pdf', sep = ''))
  
  #Remove all TimeSheet_Name.rnw
  AllTimeSheet <- list.files(path=getwd(), pattern = "^TimeSheet")
  file.remove(AllTimeSheet)
  
  #Remove others files
  OtherFile <- list.files(path=getwd(), pattern = ".tex$|.log$|txt$|.aux$")
  file.remove(OtherFile)
  
  }