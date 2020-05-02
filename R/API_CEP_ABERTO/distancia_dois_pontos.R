library(tidyverse)
library(geosphere)
library(RgoogleMaps)
library(leaflet)
library(cepR)

setwd("C:/Users/marce/Documents/Emovel/Dados/ponto_SP/dump_SP1")

CEP_SP <- read.csv('tabela.csv', header = T, sep = ",", stringsAsFactors = F)


CEP_SP <- CEP_SP %>% 
  select(cep, latitude, longitude, bairro)


coord1 <- c(-23.555113, -46.708136)

dist <- c()

for(i in 1:length(CEP_SP$latitude)){
  dist[i] <- distm(coord1, c(CEP_SP$latitude[i], CEP_SP$longitude[i]))
}

cep_distancia <- cbind(CEP_SP, dist)

cep_alto_p <- cep_distancia %>%
  filter(bairro == 'Alto de Pinheiros')

cep_alto_p <- cep_alto_p %>% 
  select(cep, bairro)

write.table(cep_alto_p, file = 'cep_alto_pinheiros.csv', sep = ';', row.names = F)
