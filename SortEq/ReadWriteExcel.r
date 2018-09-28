
#setwd("GitProjects/KPIs/SortEq/")

#function
SortOf <- function(x) {
  x[is.na(x)] <- "0"
  x[,5] <- "0"
  x[,6] <- "0"
  x[,7] <- "0"
  x[,8] <- "0"
  x[,9] <- "0"
  x[,1] <- as.character(x[,1], stric = TRUE)
  x[,2] <- as.character.numeric_version(x[,2], stric = TRUE)
  x[,3] <- as.character(x[,3], stric = TRUE)
  x[,4] <- as.character.numeric_version(x[,4], stric = TRUE)
  
  x[x=="0        "] <- "0"

  colnames(x)[5] <- "Eq.padre"
  colnames(x)[6] <- "Sistema"
  colnames(x)[7] <- "Comp"
  colnames(x)[8] <- "aux1"
  colnames(x)[9] <- "aux2"
  #Comp R
  x[,8] <- as.character(x[match(x[,4],x[,2]),3], stric = TRUE)
  x[,9] <- as.character.numeric_version(x[match(x[,4],x[,2]),4], stric = TRUE)
  x[x=="NA"] <- "0"
  x[is.na(x)] <- "0"
  x[x=="0        "] <- "0"
  #Equipo Padre E,T,V
  x[x[,4] =="0" & (x[,3] =="E" | x[,3] =="V" | x[,3] =="T"),5] <- x[x[,4] =="0" & (x[,3] =="E" | x[,3] =="V" | x[,3] =="T"),2]
  #Sistema F
  x[x[,3]=="F",6] <- x[x[,3]=="F",2]
  x[x[,3]=="F",5] <- x[x[,3]=="F",4]
  
  
  x[x[,3]=="R",7] <- x[x[,3]=="R",2]
  x[x[,8]=="F",6] <- x[x[,8]=="F",4]
  x[x[,8]=="F",5] <- x[x[,8]=="F",9]
  
  x[x[,3]=="R" & (x[,8]=="T" | x[,8]=="E" | x[,8]=="V") ,5] <- x[x[,3]=="R" & (x[,8]=="T" | x[,8]=="E" | x[,8]=="V"),4]
  x[x[,3]=="T" & x[,4]!="0",5] <- x[x[,3]=="T" & x[,4]!="0",4]
  x[x[,3]=="T" & x[,4]!="0",7] <- x[x[,3]=="T" & x[,4]!="0",2]

  x <- x[,1:7]
  return <- x
}

#install.packages("openxlsx")
library(openxlsx)
options("openxlsx.numFmt" = NULL)

#wb1 <- loadWorkbook("database/REPORTE_AL_20.08.18.xlsx")

flag <- 0

if (flag == 1) {
  ruteIn <- "../../../../Desktop/"
  ExcelInName <- "PRD SAP BV 28.09.2018.xlsx"
  wb1 <- loadWorkbook(ruteIn + ExcelInName)
}
if (flag == 0) {
  ruteIn <- choose.files()
  wb1 <- loadWorkbook(ruteIn)
}

sheet_names <- names(wb1)

Sheet1 <- read.xlsx(wb1,sheet_names[1])
#deleting void columns
Sheet2 <- Sheet1[,colSums(is.na(Sheet1))<nrow(Sheet1)]

names_sheet1 <- names(Sheet1)
dataInexcel <- (Sheet1[,c("Ubicac.técnica","Equipo","Tipo.de.equipo","Equipo.superior")])
#dataInexcel$Equipo <- as.character.numeric_version(dataInexcel$Equipo,strict = TRUE)
dataOutexcel <- SortOf(dataInexcel)
sheet1_out <- data.frame(dataOutexcel[,-4:0],Sheet2)


wb2 <- createWorkbook()
addWorksheet(wb2,"Eq")
writeData(wb2,"Eq",sheet1_out,rowNames = FALSE)

saveWorkbook(wb2, ruteIn, overwrite = TRUE)
