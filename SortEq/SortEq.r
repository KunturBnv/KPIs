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


SortOf <- function()
  
  
#dataOut <- sortOf(dataIn)
