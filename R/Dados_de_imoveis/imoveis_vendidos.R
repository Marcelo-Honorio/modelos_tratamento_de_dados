# Emovel

rm(list = ls(all = T)) # r
setwd('C:/Users/marce/Documents/Emovel')
dir()

library(readxl)
library(dplyr)
library(tidyr)

# TRANSFERINDO O DADOS INICIAIS

cas17 <- read_excel('cascavel2017.xlsx') # Imoveis Cascavel 2017
save(cas17, file = 'cas17.RData')
names(cas17)
casc2017 <- cas17 %>% 
  select(arealote, areaunidad, cadastrado, cadastroim, codigolog0, codigologr ,codigolot0, codigolote, codigoprop, codigoquad, codigounid, complement, desativado, inscricaoi, logradour0, logradouro, loteamento, numero, numerologr, padrao, proprietar, tipologia, utilizacao) %>% 
  mutate(ano = 2017)
save(casc2017, file = 'casc2017.RData') #Salvando os dados finais

cas18 <- read_excel('cascavel2018.xlsx') #Imoveis Cascavel 2018
save(cas18, file = 'cas18.RData')
names(cas18)
casc2018 <- cas18 %>% 
  select(arealote, areaunidad, cadastrado, cadastroim, codigolog0, codigologr ,codigolot0, codigolote, codigoprop, codigoquad, codigounid, complement, desativado, inscricaoi, logradour0, logradouro, loteamento, numero, numerologr, padrao, proprietar, tipologia, utilizacao) %>% 
  mutate(ano = 2018)
save(casc2018, file = 'casc2018.RData') #Salvando os dados finais

cas19 <- read_excel('cascavel2019.xlsx') #Imovei Cascavel 2019
save(cas19, file = 'cas19.RData')
names(cas19)
casc2019 <- cas19 %>% 
  select(arealote, areaunidad, cadastrado, cadastroim, codigolog0, codigologr ,codigolot0, codigolote, codigoprop, codigoquad, codigounid, complement, desativado, inscricaoi, logradour0, logradouro, loteamento, numero, numerologr, padrao, proprietar, tipologia, utilizacao) %>% 
  mutate(ano = 2019)
save(casc2019, file = 'casc2019.RData') #Salvando os dados finais


####### 28/06/19 MONTANDO O LOP
load('casc2017.RData')
load('casc2018.RData')
load('casc2019.RData')

colnames(vendas) <-  c('arealote', 'areaunidad' , 'cadastrado', 'cadastroim', 'codigolog0', 'codigologr' ,'codigolot0', 'codigolote', 'codigoprop', 'codigoquad', 'codigounid', 'complement', 'desativado', 'inscricaoi', 'logradour0', 'logradouro', 'loteamento', 'numero', 'numerologr', 'padrao', 'proprietar', 'tipologia', 'utilizacao')
colnames(continua) <-  c('arealote', 'areaunidad' , 'cadastrado', 'cadastroim', 'codigolog0', 'codigologr' ,'codigolot0', 'codigolote', 'codigoprop', 'codigoquad', 'codigounid', 'complement', 'desativado', 'inscricaoi', 'logradour0', 'logradouro', 'loteamento', 'numero', 'numerologr', 'padrao', 'proprietar', 'tipologia', 'utilizacao')

#inscricaoi como ky
for(i in 1:length(casc2019$inscricaoi)) {
  for(j in 1:length(casc2018$inscricaoi)) {
    if(casc2018[i,"inscricaoi"] == casc2019[j,'inscricaoi'] & casc2018[i,'codigoprop'] != casc2019[j, 'codigoprop']) {
      vendas <- rbind(vendas, casc2019[i, ])
    } else if(casc2018[i,"inscricaoi"] == casc2019[j,'inscricaoi'] & casc2018[i,'codigoprop'] == casc2019[j, 'codigoprop']) {
      continua <- rbind(continua, casc2019[i, ])
    } 
  }
}
casc2019[,'codigoprop'] <- as.numeric(casc2019[,'codigoprop'])
str(casc2019)
str(casc2018)
str(casc2017)



#mudando o código prop
casc2019 <- casc2019 %>% 
  mutate(codigopropi = as.numeric(codigoprop))

vendas_difere_sbrenome19 <- data.frame()
vendas_mesmo_sbrenome19 <- data.frame()
continua <- data.frame()
  
## Exemplo 2 (#codigopropi é diferente em 2019)
for(i in 1:length(casc2019$inscricaoi)) {
  for(j in 1:length(casc2018$inscricaoi)) {
    if(casc2018[i,"inscricaoi"] == casc2019[j,'inscricaoi']) {
      if(casc2018[i,'codigoprop'] != casc2019[j, 'codigopropi']) {
        sobrenome1 <- str_extract(casc2018[i,'proprietar'], '[:alpha:]+$')
        sobrenome2 <- str_extract(casc2019[j,'proprietar'], '[:alpha:]+$')
        if(sobrenome1 != sobrenome2) {
          vendas_difere_sbrenome19 <- rbind(vendas, casc2019[i, ])
        } else if (sobrenome1 == sobrenome2) {
          vendas_mesmo_sbrenome19 <- rbind(vendas, casc2019[i, ])
        }
      } else if(casc2018[i,'codigoprop'] == casc2019[j, 'codigopropi']) {
        continua <- rbind(continua, casc2019[i, ])} 
    }
  }
}

library(stringr)

vendas <- s

nomes <- c('Marcelo Honorio', 'Morgana Zwirtes')
nomes2 <- c('Ricardo Honorio', 'Marina Zwirtes')

sobrenome <- str_extract(nomes, '[:alpha:]+$')

colnames(vendas) <-  c('arealote', 'areaunidad' , 'cadastrado', 'cadastroim', 'codigolog0', 'codigologr' ,'codigolot0', 'codigolote', 'codigoprop', 'codigoquad', 'codigounid', 'complement', 'desativado', 'inscricaoi', 'logradour0', 'logradouro', 'loteamento', 'numero', 'numerologr', 'padrao', 'proprietar', 'tipologia', 'utilizacao')

for(i in length(casc2019$inscricaoi)) { 
  for(j in length(casc2018$inscricaoi)) {
    print(casc2018[i,'inscricaoi'])
    print(casc2019[j,'inscricaoi'])
  } 
}

for(i in 1:length(casc2019$inscricaoi)) {
    print(casc2018[i,'inscricaoi'])
    print(casc2019[j,'inscricaoi'])
}

rm(cas17, cas18, cas19)
rm(casc2017, casc2018, casc2019)

#reunir as tabelas
imovEmp <- rbind(casc2019, casc2018, casc2017)
save(imovEmp, file = 'imovEmp.RData')
?rename


colnames(casc2018) <- c('arealote-18', 'areaunidad-18' , 'cadastrado-18', 'cadastroim-18', 'codigolog0-18', 'codigologr-18' ,'codigolot0-18', 'codigolote-18', 'codigoprop-18', 'codigoquad-18', 'codigounid-18', 'complement-18', 'desativado-18', 'inscricaoi-18', 'logradour0-18', 'logradouro-18', 'loteamento-18', 'numero-18', 'numerologr-18', 'padrao-18', 'proprietar-18', 'tipologia-18', 'utilizacao-18')
colnames(casc2017) <- c('arealote-17', 'areaunidad-17' , 'cadastrado-17', 'cadastroim-17', 'codigolog0-17', 'codigologr-17' ,'codigolot0-17', 'codigolote-17', 'codigoprop-17', 'codigoquad-17', 'codigounid-17', 'complement-17', 'desativado-17', 'inscricaoi-17', 'logradour0-17', 'logradouro-17', 'loteamento-17', 'numero-17', 'numerologr-17', 'padrao-17', 'proprietar-17', 'tipologia-17', 'utilizacao-17')
imovlad <- cbind(casc2019, casc2018, casc2017)

save(imoveis, file = 'imoveis.RData')


#relacionar as tabelas
str(casc2019)
as.numeric(casc2019$codigoprop)
casc2019 %>% 
  arrange(codigoprop)
?spread

str(Casc2019)

str(Casc2018)

venda <- Casc2019 %>% 
  full_join(Casc2018, by = c('proprietar' = 'proprietar'))

?full_join

venda <- imoveis %>% 
  ifelse()
  filter()

