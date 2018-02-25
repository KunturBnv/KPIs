setwd("~/Repository")

da <-read.csv("Avisos 022018 IW29.CSV")
do <-read.csv("Ordenes 022018 IW39.CSV")

da <- da[,colSums(is.na(da))<nrow(da)]
da[da == ""]<-NA
do <- do[,colSums(is.na(do))<nrow(do)]
do[do == ""]<-NA
tot <- merge(do,da,by = "Orden",all = FALSE)

#tot1[,"Fin.de.avería"] = paste(tot1$Fin.de.avería[!is.na(tot1$Fin.de.avería)],tot1$Hora.fin.avería,sep =" ")
#tot1[,"Hora.fin.avería"] = as.POSIXct(tot1$Hora.fin.avería, format="%H:%M:%S")