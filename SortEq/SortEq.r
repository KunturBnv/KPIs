setwd("SortEq/")
# change a categorical variable to numeric dummie variable
install.packages("dummies")
library(dummies)

install.packages("tidyr")
library(tidyr)

#getting data
dataIn <- read.csv("dataIn.csv", na.strings = "NA")

dataIn$Eq.Padre <- dataIn$Cod..Equipo

dataIn$Eq.Padre[dataIn$Tipo.de.Objeto=="E"] <- dataIn$Cod..Equipo[dataIn$Tipo.de.Objeto=="E"]
dataIn$Eq.Padre[dataIn$Tipo.de.Objeto=="T"] <- dataIn$Cod..Equipo[dataIn$Tipo.de.Objeto=="T"]

dataIn$Eq.Padre[dataIn$Tipo.de.Objeto=="F"] <- dataIn$Equipo.Superior[dataIn$Tipo.de.Objeto=="F"]
dataIn$Sistema[dataIn$Tipo.de.Objeto=="F"] <- dataIn$Cod..Equipo[dataIn$Tipo.de.Objeto=="F"]

dataIn$Comp[dataIn$Tipo.de.Objeto=="R"] <- dataIn$Cod..Equipo[dataIn$Tipo.de.Objeto=="R"]


for (i in dataIn[,3]) {
  
  dataIn[dataIn[,3]=="E",5] = dataIn[dataIn[,3]=="E",2]
  
}


SortOf <- function(x) {
  x.row.names <- names(x)
  names(x) <- c("UT","Cod.Eq","Tipo.Eq","Eq.sup")
  i = 0
  for (m in x[,3]) {
    i++
    print(i)
  }
  
  
}  
  
#dataOut <- sortOf(dataIn)
