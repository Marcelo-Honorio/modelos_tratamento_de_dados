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

setwd('C:/Users/marce/Documents/Emovel/Dados/avaliacao_ banco/locais_proximos')

locais <- read_excel('locais_prox_itau.xlsx')


locais <- locais %>% 
  arrange(id_agencia)

agencia <- unique(locais$id_agencia)


for (j in 1:length(agencia)){
    tabela <- locais %>%
      filter(id_agencia == agencia[j])
    tabela <-  tabela %>%
      select(Port, nome_local)
    colnames(tabela) <- c('Segmento', 'Nome')
    nome_tabela <- paste("tabela_", agencia[j], sep = "")
    tabela <- tabela %>%
      arrange(Segmento)
    tabela  %>% 
      kable(align = 'c', format.args = list(decimal.mark = ",", digits = NULL, preserve.width = 'individual'), row_label_position = 'c') %>%
      kable_styling(bootstrap_options = c('striped', 'condensed', 'houver'),  full_width = F, row_label_position = "c") %>%
      row_spec(0, color = 'white', background = '#243654', font_size = 12, monospace = F) %>%
      row_spec(1:length(tabela$Nome), color = 'black') %>% 
      save_kable(file = paste(nome_tabela, '.png', sep = ''))
}

