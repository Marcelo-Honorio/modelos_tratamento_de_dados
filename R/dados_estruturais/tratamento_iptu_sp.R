rm(list = ls(all = T))
setwd('C:/Users/marce/Documents/Emovel/Dados/Area SP')
dir()
library(tidyverse)
library(readr)
library(readxl)
library(stringr)

##### instalandos pacotes para mapas
install.packages('rgdal')
install.packages('maptools')
install.packages('rgeos')
install.packages('broom')
library(rgdal)
library(maptools)
library(rgeos)
library(broom)
library(ggplot2)

#carregando os mapas
ogrListLayers('br_municipios/BRMUE250GC_SIR.shp')

mg_mapa <- readOGR('br_municipios/BRMUE250GC_SIR.shp',
                   layer = 'BRMUE250GC_SIR')
mg_mapa_data <- data.frame(mg_mapa)
head(mg_mapa_data)
mg_mapa <- fortify(mg_mapa, region = 'CD_GEOCMU')

#importando arquivos
iptu_sp <- read.csv('localizado2_IPTU.csv', sep = ';')

comercial_sp <- read.csv('comercial_IPTU.csv', sep = ';')

names(comercial_sp)
head(comercial_sp)
reg_comercial_sp <- comercial_sp %>% 
  filter((location.0 >= -46.663976 & location.0 <= -46.654471) & 
  (location.1 <= 23.561278 & location.1 >= -23.569102))
reg_comercial_sp <- reg_comercial_sp %>% 
  filter(area_util >= 500 & area_util <= 1500)

colnames(reg_comercial_sp) <- c('id', 'ano_cons', 'area_cons', 'area_ter', 'area_total', 'area_util', 'bairro',
                                'cep', 'cidade.id', 'nome_cid', 'uf', 'complemento', 'edi', 'endereco', 'estado',
                                'insc_imovel', 'iptu_key', 'location.0', 'location.1', 'logradouro', 'numero', 
                                'numero_doc', 'padrao_constr', 'proprietario', 'quant_pavi', 'status', 'status_geo',
                                'testada', 'tipo_doc', 'tipo_imov', 'tipo_terreno', 'tipo_uso', 'zoneamento')


unique(reg_comercial_sp$logradouro)
reg_comercial_sp[, 'logradouro'] <- str_trim(reg_comercial_sp$logradouro)

write.table(reg_comercial_sp, file = 'reg_comercio_sp.csv', sep = ';')

unique(reg_comercial_sp$proprietario)

##pacote leaflet
install.packages('leaflet')

library(leaflet)
leaflet() %>%
  addTiles()

install.packages('ggmap')
library(ggmap)
names(reg_comercial_sp)

leaflet(reg_comercial_sp) %>% 
  addTiles() %>% 
  addMarkers(lat = ~location.1, lng = ~location.0, popup = ~tipo_uso)
