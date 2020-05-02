rm(list = ls(all = T))
library(tidyverse)
library(stringr)
library(readr)
library(readxl)
library(lubridate)
library(formattable)
library(knitr)
library(kableExtra)
library(webshot)

setwd('C:/Users/marce/Downloads/amostra/terrenos')
dir()

#lista todos os aquivos
input <- dir(pattern = 'xlsx')
L <- length(input)

#lendo arquivos e salvando em uma lista
dados <- NULL
for(i in 1:L){
  dados[[i]] <- read_excel(input[i])
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Rua', 'R')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Avenida', 'Av')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Estrada', 'Est')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Travessa', 'Tv')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Praca', 'Pc')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, ',', '')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'á', 'a')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ã', 'a')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'à', 'a')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'é', 'e')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ê', 'e')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'í', 'i')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'õ', 'o')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ó', 'o')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ô', 'o')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ú', 'u')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'ç', 'c')
  dados[[i]][, 'Endereco'] <- str_trim(dados[[i]]$Endereco)
  #Eliminar duplicado
  dados[[i]] <- dados[[i]] %>%
    distinct(Endereco, AreaUtil, Valor, .keep_all = T)
  
  #eliminando os dodos menos que o quartil
#  quartil <- quantile(dados[[i]]$ValorMetroUtil, probs = 0.25, na.rm = T)
#  dados[[i]] <- dados[[i]] %>%
#    filter(ValorMetroUtil >= quartil)
  
  #criando uma nova area total
  dados[[i]] <- dados[[i]] %>%
    mutate(Area_Total = ifelse(AreaUtil >= AreaTotal, AreaUtil, AreaTotal))
  
  #selecionando as colunas
  dados[[i]] <- dados[[i]] %>% 
    select('Imobiliaria', 'Bairro', 'Valor', 'Quartos', 'Garagens', 'Area_Total', 'AreaUtil', 'ValorMetroUtil', 'Endereco', 'DataAnuncio', 'Link')
  
  #criando uma coluna
  dados[[i]][, 'Valor'] <- as.numeric(dados[[i]]$Valor)
  dados[[i]] <- dados[[i]] %>% 
    mutate(ValorMetro = Valor/Area_Total)
  #transformando os dados em date
  dados[[i]][, 'DataAnuncio'] <- dmy(dados[[i]]$DataAnuncio)
  
  #tabela para correção
  nome <- input[i]
  nome <- str_replace_all(nome, '.xlsx', '')
  agencia <- str_extract(nome, '\\d+')
  nome <- str_replace(nome, '\\-', '')
  nome <- str_replace(nome, agencia, '')
  nome <- str_trim(nome)
  nome <- str_replace_all(nome, ' ', '_')
  nome_resumo <- paste( agencia, '_',nome, "_tab_limpa", sep = '')
  write.table(dados[[i]], file = paste(nome_resumo, '.csv', sep = ''), row.names = F, sep = ';', dec = ',')
  
  #criando a coluna faixa e dias de anuncio
  dados[[i]] <- dados[[i]] %>% 
    mutate(TempoAnuncio = Sys.Date() - DataAnuncio, 
           faixa = ifelse(Area_Total < 250, 'menos de 250',
                          ifelse(Area_Total>= 250 & Area_Total< 500, '250-499',
                          ifelse(Area_Total>= 500 & Area_Total< 750, '500-749',
                          ifelse(Area_Total>= 750 & Area_Total< 1000, '750-999',
                          ifelse(Area_Total>= 1000 & Area_Total< 1250, '1000-1249',
                          ifelse(Area_Total>= 1250 & Area_Total< 1500, '1250-1499',
                          ifelse(Area_Total>= 1500 & Area_Total< 1750, '1500-1749',
                          ifelse(Area_Total>= 1750 & Area_Total< 2000, '1750-1999',
                          ifelse(Area_Total>= 2000 & Area_Total< 2250, '2000-2249',
                          ifelse(Area_Total>= 2250 & Area_Total< 2500, '2250-2499','acima de 2500'))))))))))
    )
  
  ## Calculando os parametros
  amostra <- length(dados[[i]]$Endereco)
  med_valor_metroTotal <- median(dados[[i]]$ValorMetro)
  med_valor_metroUtil <- median(dados[[i]]$ValorMetroUtil)
  med_valor_imovel <- median(dados[[i]]$Valor)
  med_area_total <- median(dados[[i]]$Area_Total)
  tempo_anuncio <- median(dados[[i]]$TempoAnuncio)
  resumo_calculo <- rbind(amostra, med_valor_metroTotal, med_valor_metroUtil, med_valor_imovel, med_area_total, tempo_anuncio)
  
  ## Salvando os parametros
  nome_resumo <- paste( agencia, '_',nome, "_parametros", sep = '')
  write.table(resumo_calculo, file = paste(nome_resumo, '.txt', sep = ''), sep = ' -> ')
  #criando o resumo
  dados[[i]] <- dados[[i]] %>%
    group_by(faixa) %>%
    summarise(numero = length(Endereco),
              med_areaTotal = median(Area_Total),
              med_areaUtil = median(AreaUtil),
              med_valor = median(Valor),
              med_valor_metroTotal = median(ValorMetro),
              med_valor_metroUtil = median(ValorMetroUtil),
              med_t_anuncio = median(as.numeric(TempoAnuncio))) %>%
    arrange(med_areaTotal)
  
  #TABELA
  colnames(dados[[i]]) <- c('Metragem', 'Amostra', 'Med. Área', 'Med. Area Util', 'Med. Valor', 'Med. Valor m² total', 'Med. Valor m² Útil', 'Med. Tempo de anúncio')
  
  dados[[i]]$`Med. Valor` <- currency(dados[[i]]$`Med. Valor`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor` <- as.character(dados[[i]]$`Med. Valor`)
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\.', ':')
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\,', '.')
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\:', ',')
  
  dados[[i]]$`Med. Valor m² total` <- currency(dados[[i]]$`Med. Valor m² total`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor m² total` <- as.character(dados[[i]]$`Med. Valor m² total`)
  dados[[i]]$`Med. Valor m² total` <- str_replace_all(dados[[i]]$`Med. Valor m² total`, '\\.', ':')
  dados[[i]]$`Med. Valor m² total` <- str_replace_all(dados[[i]]$`Med. Valor m² total`, '\\,', '.')
  dados[[i]]$`Med. Valor m² total` <- str_replace_all(dados[[i]]$`Med. Valor m² total`, '\\:', ',')
  
  dados[[i]]$`Med. Valor m² Útil` <- currency(dados[[i]]$`Med. Valor m² Útil`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor m² Útil` <- as.character(dados[[i]]$`Med. Valor m² Útil`)
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\.', ':')
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\,', '.')
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\:', ',')
  
  dados[[i]] <- dados[[i]] %>%
    select(Metragem, Amostra, `Med. Área`, `Med. Valor`, `Med. Valor m² total`, `Med. Tempo de anúncio`)
  
  dados[[i]]$`Med. Área` <- round(dados[[i]]$`Med. Área`)
  dados[[i]]$`Med. Tempo de anúncio` <- round(dados[[i]]$`Med. Tempo de anúncio`)
  
  
  
  nome_resumo <- paste( agencia, '_',nome, "_resumo", sep = '')
  dados[[i]] %>% 
    kable(align = 'c', format.args = list(decimal.mark = ",", digits = NULL, preserve.width = 'individual'), row_label_position = 'c') %>%
    kable_styling(bootstrap_options = c('striped', 'condensed', 'houver'),  full_width = F, row_label_position = "c") %>%
    row_spec(0, color = 'white', background = '#243654', font_size = 12, monospace = F) %>%
    row_spec(1:length(dados[[i]]$Amostra), color = 'black') %>% 
    save_kable(file = paste(nome_resumo, '.png', sep = ''))
  
  cat(input[i],'/n')
  
}

# Salvando o resumo
# nome_resumo <- paste("resumo_", nome, sep = "")
# nome_resumo <- str_replace_all(nome_resumo, '.xlsx', '')
# write.table(dados[[i]], file = paste(nome_resumo, '.csv', sep = ''), row.names = F, sep = ';', dec = ',')

agencias <- str_extract(input, '\\d+')
agencias <- unique(agencias)
write.table(agencias, file = 'agencias.txt', row.names = F, col.names = F)
