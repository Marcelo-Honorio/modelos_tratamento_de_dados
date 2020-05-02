library(tidyverse)
setwd("C:/Users/marce/Documents/Emovel/Dados/ponto_SP/dump_SP1")

SP1 <- read.csv("sp.cepaberto_parte_1.csv", header = F, sep = ',')
SP2 <- read.csv("sp.cepaberto_parte_2.csv", header = F, sep = ',')
SP3 <- read.csv("sp.cepaberto_parte_3.csv", header = F, sep = ',')
SP4 <- read.csv("sp.cepaberto_parte_4.csv", header = F, sep = ',')
SP5 <- read.csv("sp.cepaberto_parte_5.csv", header = F, sep = ',')

todos_cep <- rbind(SP1, SP2, SP3, SP4, SP5)

colnames(todos_cep) <- c("CEP", "Rua", "Bairro", "id_cidade", "id_estado")

todos_cep <- todos_cep %>% 
  filter(id_cidade == 8966)

CEP_SP <- todos_cep %>%
  select(CEP, Bairro)

CEP_SP <- CEP_SP %>%
  filter(Bairro %in% c('Vila Madalena', 'Pinheiros', 'Pinheiro', 'Butantã', 'Alto de Pinheiros'))


write.table(CEP_SP, file = "CEP_SP.csv", row.names = F, sep = ',')
