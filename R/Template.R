###Reading R function to generate automatically the timesheet for all employees
###############################################################################

Time_Sheet <- function(Year=2016, Month='January', IndexSheet=3){
  
  #Defining interactivly the directory of main folder
  ###################################################
  directory <- choose.dir(default = "", caption = "Choisir le Dossier principal")
  setwd(directory)
  
  
  ##Importing the time sheet data###
  ##################################
  library(readxl)
  library(stringr)
  TS_data.0 <- read_excel('Data/TS_2016_data.xlsx', sheet = IndexSheet)
  
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
    
    for (j in 1:dim(TableProgram)[2]) {
      TableProgram[i,j] <- ifelse(TS_data.0[k, ColumnDays[j]]=='8', 8, 0)*TS_data.0[k, ColumnProject[i]]/100
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
  colnames(TableAbsence) <- as.character(1:length(Days))
  
  ##For one employee
  List.absence <- c('W','A','P','I','O')
  for (i in 1:dim(TableAbsence)[1]){
    
    for (j in 1:dim(TableAbsence)[2]) {
      TableAbsence[i,j] <- ifelse(TS_data.0[k, ColumnDays[j]]==List.absence[i], 8, 0)
    }
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
  #library(tools)
  library(plotflow)
  setwd('./TimeSheetTemplate')

 #knit2pdf("TimeSheet_Template.Rnw")
  
  ##Generate all TimeSheet_Name.tex
  for (i in 1:length(unique(NameStaff))){
    knit("TimeSheet_Template.Rnw",
             output=paste('TimeSheet_', NameStaff[i], '.tex', sep=""))

  }

  ##Generate PDFs from the tex files for each employee
  lapply(unique(NameStaff), function(x)
    texi2pdf(paste0('TimeSheet_', x, '.tex', sep=""), clean = T, quiet = TRUE))

  ##Copy all pdfs from 'TimeSheetTemplate' to 'TimeSheetPDF'
  AllPdf <- dir(path=getwd(), pattern = ".pdf$", full.names = T)
  Allname <- list.files(path=getwd(), pattern = ".pdf$")
  Folder.dest <- paste(paste(directory, '/TimeSheetPDF', sep=''), Allname, sep = '/')
  file.copy(from = AllPdf, to = Folder.dest, overwrite = T)
 #  
 # ##Merge all PDFs into one PDF
 #  setwd('../')
 #  plotflow:::mergePDF(
 #    in.file=paste(file.path("TimeSheetPDF", dir("TimeSheetPDF")), collapse=" "),
 #    file=paste('TimeSheet_CERMEL_', Sys.Date(), '.pdf', sep = '')
 #  )

  
}