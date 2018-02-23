setwd("~/Repository")

da <-read.csv("Avisos 022018 IW29.CSV")
do <-read.csv("Ordenes 022018 IW39.CSV")

da <- da[,colSums(is.na(da))<nrow(da)]
da[da == ""]<-NA
do <- do[,colSums(is.na(do))<nrow(do)]
do[do == ""]<-NA
tot <- merge(do,da,by = "Orden",all = FALSE)
