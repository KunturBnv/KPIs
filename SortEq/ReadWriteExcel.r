
setwd("GitProjects/KPIs/SortEq/")

#function
SortOf <- function(x) {
  x[is.na(x)] <- "0"
  x[,5] <- "0"
  x[,6] <- "0"
  x[,7] <- "0"
  x[,8] <- "0"
  x[,1] <- as.character(x[,1])
  x[,2] <- as.character(x[,2])
  x[,3] <- as.character(x[,3])
  x[,4] <- as.character(x[,4])
  
  
  colnames(x)[5] <- "Eq.padre"
  colnames(x)[6] <- "Sistema"
  colnames(x)[7] <- "Comp"
  colnames(x)[8] <- "aux"
  #Equipo Padre E,T,V
  x[x[,4] =="0" & (x[,3] =="E" | x[,3] =="V" | x[,3] =="T"),5] <- x[x[,4] =="0" & (x[,3] =="E" | x[,3] =="V" | x[,3] =="T"),2]
  #Sistema F
  x[x[,3]=="F",6] <- x[x[,3]=="F",2]
  x[x[,3]=="F",5] <- x[x[,3]=="F",4]
  #Comp R
  x[,8] <- as.character(x[match(x[,4],x[,2]),3])
  x[,9] <- as.character(x[match(x[,4],x[,2]),4])
  x[is.na(x)] <- "0"
  x[x[,3]=="R",7] <- x[x[,3]=="R",2]
  x[x[,8]=="F",6] <- x[x[,8]=="F",4]
  x[x[,8]=="F",5] <- x[x[,8]=="F",9]
  
  x[x[,3]=="R" & (x[,8]=="T" | x[,8]=="E" | x[,8]=="V") ,5] <- x[x[,3]=="R" & (x[,8]=="T" | x[,8]=="E" | x[,8]=="V"),4]
  x[x[,3]=="T" & x[,4]!="0",5] <- x[x[,3]=="T" & x[,4]!="0",4]
  x[x[,3]=="T" & x[,4]!="0",7] <- x[x[,3]=="T" & x[,4]!="0",2]
  x <- x[,1:7]
  return <- x
}

library(readxl)


dataInExcel <- read_xlsx("database/REPORTE AL 20.08.18.xlsx", col_types = "text")
View(dataInExcel)

library(openxlsx)
options("openxlsx.numFmt" = NULL)

wb1 <- loadWorkbook("database/REPORTE_AL_20.08.18.xlsx")

sheet_names <- names(wb1)

Sheet1 <- read.xlsx(wb1,sheet_names[1])

names_sheet1 <- names(Sheet1)



dataInexcel <- (Sheet1[,c("Ubicac.técnica","Equipo","Tipo.de.equipo","Equipo.superior")])

dataOutexcel <- SortOf(dataInexcel)

sheet1_out <- data.frame(dataOutexcel[,-4:0],Sheet1)

addWorksheet(wb1,"newsheet")

writeData(wb1,"newsheet",sheet1_out,rowNames = FALSE)

saveWorkbook(wb1, "database/REPORTE_AL_20.08.18v2.xlsx", overwrite = TRUE)


