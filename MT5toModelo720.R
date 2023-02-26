library(dplyr)
library(lubridate)
library(gdata)

setwd("/Users/MODELO720")
cuentas<- list.files('/Users/MODELO720/xlsx') #Directorio con uno o más archivos XLSX generados en Metatrader 5

#función para rellenar los NA con el valor previo
rellenanas<- function(x) {
  for(i in 1:nrow(x)) {
    if(is.na(x[i, 3])) {
      x[i, 3] <- x[i - 1, 3]
    }
  }
  x
}

for (i in 1:length(cuentas)) {
  cuenta=cuentas[i]
  index1=0
  index2=0
  index3=0
  index4=0
  start=0
  finish=0
  finish2=0
  len=0
  #Lectura y formato de los datos 
  data<- read.xls(paste0("xlsx/", cuentas[i]))
  index<- which(data == c("Deals"), arr.ind=TRUE)
  start<-as.numeric(index[1,1]) #marca el comienzo de la lectura
  index2<-which(data == c("Balance:"), arr.ind=TRUE)
  finish=index2[1,1] #marca el final de la lectura
  data <- data[(start+2):(finish-2),]
  try(index3<-which(data == c("Open Positions"), arr.ind=TRUE))
  if (dim(index3)[1]==1) {(finish2<-as.numeric(index3[1,1]))}
  if (dim(index3)[1]==1) {data <- data[1:(finish2-2),]} #marca un final alternativo de la lectura cuando existen posiciones abiertas
  data<- data[1:length(na.omit(data$X.11)),] #conserva únicamente los datos válidos
  data$X.11<-gsub(" ", "", data$X.11) #elimina espacios indeseables en la porción numérica del balance
  if(sum(na.omit(as.numeric(data$X.11)))==0) {data<- data[1,]
  data$Trade.History.Report[1]<-"2022-10-01"} #solución breve a cuentas con balance 0 
  
  #Damos formato a los valores
  data$X.11<-as.numeric(gsub(" ", "", data$X.11)) #elimina espacios indeseables en la porción numérica del balance
  data$Trade.History.Report<-gsub("\\.", "-", data$Trade.History.Report) #cambia el formato de la fecha
  data$Trade.History.Report<-as.Date(data$Trade.History.Report) #otorga formato de fecha
  if(is.na(data$Trade.History.Report[1])) {data$Trade.History.Report[1]<-as.Date("2022-10-01")} #cuando existe un único valor será el de la primera fecha
  
  #Creamos un nuevo único data frame con fecha y balance
  Date<-as.Date(data$Trade.History.Report)
  Balance<-data$X.11
  Consolidado<- as.data.frame(cbind(Date, Balance))
  Consolidado$Date <- as.Date(Date)
  
  #Se consigue la media para cada día
  Consolidado1<- Consolidado %>% 
    group_by(Date) %>%
    summarize(SaldoMedio = mean(Balance))
  
  #Creamos cadena de días para un array nuevo
  datestart <- as.Date("2022/10/01")
  # Definimos su longitud (92 días últimos días para el Modelo 720)
  len <- 92
  # Generamos rango de fechas para el modelo 720
  Modelo720days<- seq(datestart, by = "day", length.out = len)
  TablaFinal<- as.data.frame(cbind(Modelo720days, Balance=0))
  TablaFinal$Modelo720days<- as.Date(Modelo720days)
  colnames(TablaFinal) <- c("Date", "SaldoMedio")
  Tablasunidas<- left_join(TablaFinal, Consolidado1, by = c("Date" = "Date"))
  Tablasunidas$SaldoMedio.y[1]<- na.omit(Tablasunidas$SaldoMedio.y)[1]

  #Se rellenan los valores de días sin operaciones con el balance del último día válido
  ResultadoFinal<- rellenanas(Tablasunidas)
  
  #Se imprime el número de cuenta seguido del número final del balance en enteros
  print(paste(cuenta, as.integer(mean(ResultadoFinal$SaldoMedio.y))))
  index1=0
  index2=0
  index3=0
  index4=0
  start=0
  finish=0
  finish2=0
  len=0
  
  #Se guarda el restulado en un archivo csv
  resultados<- data.frame(id = cuenta,
                          saldomedio = mean(ResultadoFinal$SaldoMedio.y))
  write.csv(resultados, paste(cuenta,"resultados.csv"))
}

