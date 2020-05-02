## Vendas 2019

rm(list = ls(all = T)) # iniciando
setwd('C:/Users/marce/Documents/Emovel')
dir()

library(readxl)
library(dplyr)
library(tidyr)
library(stringr)

# carregando os arquivos
load('casc2017.RData')
load('casc2018.RData')
load('casc2019.RData')


#filtrar os NA
is.na(casc2019[, "inscricaoi"])
casc2019 <- casc2019 %>% 
  mutate(codigopropi = as.numeric(codigoprop))
casc2019 <- casc2019 %>% 
  filter(inscricaoi != 'NA', codigopropi != 'NA')
casc2018 <- casc2018 %>% 
  filter(inscricaoi != 'NA', codigoprop != 'NA') 

# Limapando a coluna proprietar17io 2019
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'LTDA', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '-LTDA', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '- ME', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, ' ME ', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '- EPP', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, ' -EPP', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'EPP', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '& CIA', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'E CIA', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '&', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '@', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'E ESPOSA', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'E ESPOSO', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, 'E OUTROS', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'A')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'E')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'C')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'O')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'A')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'I')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'U')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '�', 'O')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '   *', '')
casc2019[ , 'proprietar17'] <- str_replace(casc2019$proprietar17, '   ***', '')
casc2019[ , 'proprietar17'] <- str_replace_all(casc2019$proprietar17, '\\.|/|-*()', '')
casc2019[ , 'proprietar17'] <- str_trim(casc2019$proprietar17)
casc2019[ , 'inscricaoi'] <- str_replace_all(casc2019$inscricaoi, '\\.|/|-*()', '')
casc2019[ , 'inscricaoi'] <- str_trim(casc2019$inscricaoi)

# Limapando a coluna proprietar17io 2018
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'LTDA', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '-LTDA', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '- ME', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, ' ME ', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '- EPP', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, ' -EPP', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'EPP', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '& CIA', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'E CIA', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '&', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '@', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'E ESPOSA', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'E ESPOSO', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, 'E OUTROS', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'A')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'E')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'C')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'O')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'A')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'I')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'U')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '�', 'O')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '   *', '')
casc2018[ , 'proprietar17'] <- str_replace(casc2018$proprietar17, '   ***', '')
casc2018[ , 'proprietar17'] <- str_replace_all(casc2018$proprietar17, '\\.|/|-*()', '')
casc2018[ , 'proprietar17'] <- str_trim(casc2018$proprietar17)
casc2018[ , 'inscricaoi'] <- str_replace_all(casc2018$inscricaoi, '\\.|/|-*()', '')
casc2018[ , 'inscricaoi'] <- str_trim(casc2018$inscricaoi)

# Criando uma nova coluna
casc2019 <- casc2019 %>% 
  mutate(sobrenome = str_extract(casc2019$proprietar17, '[:alpha:]+$'))
casc2018 <- casc2018 %>% 
  mutate(sobrenome = str_extract(casc2018$proprietar17, '[:alpha:]+$'))

is.na(casc2019$sobrenome)

vendas18 <- data.frame()
vendas19 <- data.frame()
continua19 <- data.frame()

#Lop para verificar as vendas
## 2019 (#codigopropi � diferente em 2019)
for(i in 1:length(casc2018$inscricaoi)) {
  for(j in 1:length(casc2019$inscricaoi)) {
    if(casc2018[i,'inscricaoi'] == casc2019[j,'inscricaoi']) {
      if(casc2018[i,'codigoprop'] != casc2019[j, 'codigopropi']) {
        sobrenome1 <- str_extract(casc2018[i,'proprietar17'], '[:alpha:]+$')
        sobrenome2 <- str_extract(casc2019[j,'proprietar17'], '[:alpha:]+$')
        if(sobrenome1 != sobrenome2) {
          vendas_difere_sbrenome19 <- rbind(vendas_difere_sbrenome19, casc2019[j, ])
        } else if (sobrenome1 == sobrenome2) {
          vendas_mesmo_sbrenome19 <- rbind(vendas_mesmo_sbrenome19, casc2019[j, ])
        }
      } else if(casc2018[i,'codigoprop'] == casc2019[j, 'codigopropi']) {
        continua19 <- rbind(continua19, casc2019[i, ])}
    }
  }
}


# exemplo resumido
for(i in 1:length(casc2018$inscricaoi)) {
  for(j in 1:length(casc2019$inscricaoi)) {
    if(casc2018[i,'inscricaoi'] == casc2019[j,'inscricaoi']) {
      if(casc2018[i,'codigoprop'] != casc2019[j, 'codigopropi']) {
        vendas18 <- rbind(vendas18, casc2018[i, ])
        vendas19 <- rbind(vendas19, casc2019[j, ])
      } else if(casc2018[i,'codigoprop'] == casc2019[j, 'codigopropi']) {
        continua19 <- rbind(continua19, casc2019[j, ])}
    }
  }
}

############### REUNIR OS DADOS DE 2017 E 2018 ############################
load('casc2017.RData')

resumido2017 <- casc2017 %>% 
  select(inscricaoi, codigoprop, proprietar17)

resumido2017 <- resumido2017 %>%
  filter(inscricaoi != 'NA', codigoprop != 'NA', proprietar17 != 'NA')

colnames(resumido2017)[2] <- 'codigoprop17'
colnames(resumido2017)[3] <- 'proprietar1717'

load('casc2018.RData')

# Ordenar
?sort

resumido2017 <- resumido2017 %>% 
  arrange(inscricaoi)

casc2018 <- casc2018 %>% 
  arrange(inscricaoi)

save(resumido2017, file = 'resumido2017.RData')
save(casc2018, file = 'casc2018.RData')

venda2018 <- casc2018 %>% 
  full_join(resumido2017, by = c('inscricaoi' = 'inscricaoi'))

save(venda2018, file = 'venda2018.RData')

#Lop para salvar as vendas efetivas

venda2018 <- venda2018 %>% 
    mutate(vendido18 = 1, codigoprop != codigoprop17)


venda2018 <- venda2018 %>% 
  filter(codigoprop != 'NA' & codigoprop17 != 'NA')

## lop para separar as vendas do continua
venda2018[, 'codigoprop'] <- as.numeric(venda2018$codigoprop)
venda2018[, 'codigoprop17'] <- as.numeric(venda2018$codigoprop17)

str(venda2018$codigoprop)
str(venda2018$codigoprop17)

venda2018_cod_diferente <- data.frame()
continua2018_mesmo_prop <- data.frame()
 
venda2018[, 'codigoprop'] <-   str_trim(venda2018$codigoprop)
venda2018[, 'codigoprop17'] <- str_trim(venda2018$codigoprop17)
venda2018[, 'codigoprop'] <- str_replace_all(venda2018$codigoprop, '\\.|/|-*()', '')
venda2018[, 'codigoprop17'] <- str_replace_all(venda2018$codigoprop17, '\\.|/|-*()', '')

for(i in length(venda2018$inscricaoi)){
  if(venda2018[i,'codigoprop'] != venda2018[i,'codigoprop17']){
    venda2018_cod_diferente <- rbind(venda2018_cod_diferente, venda2018[i, ])
  } else if(venda2018[ i,'codigoprop'] == venda2018[i,'codigoprop17']){
    continua2018_mesmo_prop <- rbind(continua2018_mesmo_prop, venda2018[i,])
  }
}

#limpando o nome
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'LTDA', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '-LTDA', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '- ME', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, ' ME ', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '- EPP', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, ' -EPP', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'EPP', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '& CIA', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'E CIA', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '&', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '@', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'E ESPOSA', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'E ESPOSO', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, 'E OUTROS', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'A')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'E')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'C')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'O')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'A')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'I')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'U')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '�', 'O')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '   *', '')
venda2018[ , 'proprietar'] <- str_replace(venda2018$proprietar, '   ***', '')
venda2018[ , 'proprietar'] <- str_replace_all(venda2018$proprietar, '\\.|/|-*()', '')
venda2018[ , 'proprietar'] <- str_trim(venda2018$proprietar)

# Limapando a coluna proprietar17io 2018
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'LTDA', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '-LTDA', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '- ME', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, ' ME ', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '- EPP', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, ' -EPP', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'EPP', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '& CIA', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'E CIA', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '&', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '@', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'E ESPOSA', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'E ESPOSO', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, 'E OUTROS', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'A')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'E')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'C')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'O')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'A')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'I')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'U')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '�', 'O')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '   *', '')
venda2018[ , 'proprietar17'] <- str_replace(venda2018$proprietar17, '   ***', '')
venda2018[ , 'proprietar17'] <- str_replace_all(venda2018$proprietar17, '\\.|/|-*()', '')
venda2018[ , 'proprietar17'] <- str_trim(venda2018$proprietar17)

# Criando uma nova coluna
venda2018 <- venda2018 %>% 
  mutate(sobrenome18 = str_extract(venda2018$proprietar, '[:alpha:]+$'))
venda2018 <- venda2018 %>% 
  mutate(sobrenome17 = str_extract(venda2018$proprietar17, '[:alpha:]+$'))

save(venda2018, file = 'venda2018.txt') #salvando em excel

#Filtrando dados
dir()
load('Venda2018.RData')

?save
write.table(venda2018, file = 'venda2018.csv', sep = ';')
