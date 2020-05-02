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


setwd('C:/Users/marce/Downloads/amostra/amostra_venda')
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
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Alameda', 'Al')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'AL', 'Al')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Al.', 'Al')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Travessa', 'Tv')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Praca', 'Pc')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'Quadra', 'Qdr')
  dados[[i]][, 'Endereco'] <- str_replace_all(dados[[i]]$Endereco, 'QUADRA', 'Qdr')
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
  
  #selecionando colunas
  dados[[i]] <- dados[[i]] %>% 
    select('Imobiliaria', 'Bairro', 'Valor', 'Quartos', 'Garagens', 'AreaUtil', 'ValorMetroUtil', 'Endereco', 'DataAnuncio', 'Link')
  #criando colunas
  dados[[i]] <- dados[[i]] %>%  
    mutate(AreaTotal = (AreaUtil+ 17 +(11*Garagens)), ValorMetro = Valor/AreaTotal)
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
           faixa = ifelse(AreaTotal < 50, 'menos de 50',
                          ifelse(AreaTotal>= 50 & AreaTotal < 60, '50-59',
                          ifelse(AreaTotal>= 60 & AreaTotal < 70, '60-69',
                          ifelse(AreaTotal>= 70 & AreaTotal < 80, '70-79',
                          ifelse(AreaTotal>= 80 & AreaTotal < 90, '80-89',
                          ifelse(AreaTotal>= 90 & AreaTotal < 100, '90-99',
                          ifelse(AreaTotal>= 100 & AreaTotal < 110, '100-109',
                          ifelse(AreaTotal>= 110 & AreaTotal < 120, '110-119',
                          ifelse(AreaTotal>= 120 & AreaTotal < 130, '120-129',
                          ifelse(AreaTotal>= 130 & AreaTotal < 140, '130-139',
                          ifelse(AreaTotal>= 140 & AreaTotal < 150, '140-149',
                          ifelse(AreaTotal>= 150 & AreaTotal < 160, '150-159',
                          ifelse(AreaTotal>= 160 & AreaTotal < 170, '160-169',
                          ifelse(AreaTotal>= 170 & AreaTotal < 180, '170-179',
                          ifelse(AreaTotal>= 180 & AreaTotal < 190, '180-189',
                          ifelse(AreaTotal>= 190 & AreaTotal < 200, '190-199',
                          ifelse(AreaTotal>= 200 & AreaTotal < 210, '200-209',
                          ifelse(AreaTotal>= 210 & AreaTotal < 220, '210-219',
                          ifelse(AreaTotal>= 220 & AreaTotal < 230, '220-229',
                          ifelse(AreaTotal>= 230 & AreaTotal < 240, '230-239',
                          ifelse(AreaTotal>= 240 & AreaTotal < 250, '240-249', 'acima de 250'
                                 )))))))))))))))))))))
           )
  
  ## Calculando os parametros
  amostra <- length(dados[[i]]$Endereco)
  med_valor_metroTotal <- median(dados[[i]]$ValorMetro)
  med_valor_metroUtil <- median(dados[[i]]$ValorMetroUtil)
  med_valor_imovel <- median(dados[[i]]$Valor)
  med_area_total <- median(dados[[i]]$AreaTotal)
  tempo_anuncio <- median(dados[[i]]$TempoAnuncio)
  resumo_calculo <- rbind(amostra, med_valor_metroTotal, med_valor_metroUtil, med_valor_imovel, med_area_total, tempo_anuncio)
  
  ## Salvando os parametros
  nome_resumo <- paste(agencia, '_',nome, "_parametros", sep = '')
  write.table(resumo_calculo, file = paste(nome_resumo, '.txt', sep = ''), sep = ' -> ')
  
  #resumo da amostra
  dados[[i]] <- dados[[i]] %>%
    group_by(faixa) %>%
    summarise(amostra = length(Endereco),
              med_area_total = median(AreaTotal),
              med_valor = median(Valor),
              med_valor_metroTotal = median(ValorMetro),
              med_valor_metroUtil = median(ValorMetroUtil),
              med_n_quarto = median(Quartos),
              med_q_garage = median(Garagens),
              med_t_anuncio = median(as.numeric(TempoAnuncio))) %>%
    arrange(med_area_total)
  
  #TABELA
  colnames(dados[[i]]) <- c('Metragem', 'Amostra', 'Med. Área', 'Med. Valor', 'Med. Valor m²', 'Med. Valor m² Útil', 'Quarto', 'Vagas', 'Med. Tempo de anúncio')
  
  dados[[i]]$`Med. Valor` <- currency(dados[[i]]$`Med. Valor`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor` <- as.character(dados[[i]]$`Med. Valor`)
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\.', ':')
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\,', '.')
  dados[[i]]$`Med. Valor` <- str_replace_all(dados[[i]]$`Med. Valor`, '\\:', ',')
  
  dados[[i]]$`Med. Valor m²` <- currency(dados[[i]]$`Med. Valor m²`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor m²` <- as.character(dados[[i]]$`Med. Valor m²`)
  dados[[i]]$`Med. Valor m²` <- str_replace_all(dados[[i]]$`Med. Valor m²`, '\\.', ':')
  dados[[i]]$`Med. Valor m²` <- str_replace_all(dados[[i]]$`Med. Valor m²`, '\\,', '.')
  dados[[i]]$`Med. Valor m²` <- str_replace_all(dados[[i]]$`Med. Valor m²`, '\\:', ',')
  
  dados[[i]]$`Med. Valor m² Útil` <- currency(dados[[i]]$`Med. Valor m² Útil`, 'R$', sep = ' ')
  dados[[i]]$`Med. Valor m² Útil` <- as.character(dados[[i]]$`Med. Valor m² Útil`)
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\.', ':')
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\,', '.')
  dados[[i]]$`Med. Valor m² Útil` <- str_replace_all(dados[[i]]$`Med. Valor m² Útil`, '\\:', ',')
  
  dados[[i]] <- dados[[i]] %>%
    select(Metragem, Amostra, `Med. Área`, `Med. Valor`, `Med. Valor m² Útil`, Vagas, `Med. Tempo de anúncio`)
  
 dados[[i]]$`Med. Área` <- round(dados[[i]]$`Med. Área`)
 dados[[i]]$`Med. Tempo de anúncio` <- round(dados[[i]]$`Med. Tempo de anúncio`)
  
 nome_resumo <- paste( agencia, '_',nome, "_resumo", sep = '')
 nome_resumo <- str_replace_all(nome_resumo, '.xlsx', '') 
 dados[[i]] %>% 
    kable(align = 'c', format.args = list(decimal.mark = ",", digits = NULL, preserve.width = 'individual'), row_label_position = 'c') %>%
    kable_styling(bootstrap_options = c('striped', 'condensed', 'houver'),  full_width = F, row_label_position = "c") %>%
    row_spec(0, color = 'white', background = '#243654', font_size = 12, monospace = F) %>%
    row_spec(1:length(dados[[i]]$Amostra), color = 'black') %>% 
    save_kable(file = paste(nome_resumo, '.png', sep = ''))
  
  
  cat(input[i],'/n')
 
}

#agencias <- str_extract(input, '\\d+')
#agencias <- unique(agencias)
#write.table(agencias, file = 'agencias.txt', row.names = F, col.names = F)





