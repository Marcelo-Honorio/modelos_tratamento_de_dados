library(readr)
library(tidyverse)
library(geosphere)
library(RgoogleMaps)
library(leaflet)
library(cepR)

setwd("C:/Users/marce/Documents/Emovel/Dados/ponto_SP")

IPTU_SP <- read.csv('plataform_IPTU.csv', header = T, sep = ";", stringsAsFactors = F)

coord1 <- c(IPTU_SP$location.1[1], IPTU_SP$location.0[1])

coord2 <- c(IPTU_SP$location.1[2], IPTU_SP$location.0[2])

dist <- c()

for(i in 1:1473){
  dist[i] <- distm(coord1,c(IPTU_SP$location.1[i], IPTU_SP$location.0[i]))
}

NOVO_IPTU <- cbind(IPTU_SP, dist)
  
resumo <- NOVO_IPTU %>% 
  group_by(tipo_imovel) %>% 
  summarise(numero = length(tipo_imovel),
            porcent = round(numero/1473, 2),
            dist_m = median(dist)) %>%
  arrange(dist_m)

resumo2 <- NOVO_IPTU %>% 
  group_by(tipo_uso) %>% 
  summarise(numero = length(tipo_imovel),
            porcent = numero/1473,
            dist_m = median(dist)) %>%
  arrange(dist_m)

IPTU_TOTAL <- read.csv('ok_IPTU_SP.csv', header = T, sep = ";", stringsAsFactors = F)

IPTU_TOTAL <- IPTU_TOTAL %>%
  select(X_id, VALOR.DO.M2.DO.TERRENO, VALOR.DO.M2.DE.CONSTRUCAO, TIPO.DE.USO.DO.IMOVEL,
         TIPO.DE.PADRAO.DA.CONSTRUCAO, AREA.DO.TERRENO, AREA.CONSTRUIDA, 
         ANO.DA.CONSTRUCAO.CORRIGIDO, NUMERO.DO.CONTRIBUINTE)

IPTU_TOTAL <- IPTU_TOTAL %>%
  left_join(NOVO_IPTU, by = c('NUMERO.DO.CONTRIBUINTE' = 'inscricao_imovel'))


write.table(IPTU_TOTAL, 'IPTU_FINAL.csv', sep = ';', row.names = F)


resumo3 <- IPTU_TOTAL %>% 
  group_by(tipo_imovel) %>% 
  summarise(numero = length(tipo_imovel),
            porcent = round(numero/1473, 2),
            ano_const = round(mean(ano_construcao)),
            dist_m = median(dist)) %>%
  arrange(dist_m)

str(IPTU_TOTAL)

IPTU_MIL <- NOVO_IPTU %>%
  filter(dist <= 100)

leaflet(IPTU_MIL) %>%
  addTiles() %>%
  addMarkers(~location.0, ~location.1)
