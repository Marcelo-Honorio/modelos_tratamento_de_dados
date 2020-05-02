# -*- coding: utf-8 -*-
import pandas as pd
from geopy.distance import geodesic

tabela = pd.read_csv("tabela.csv")

tabela = tabela.filter(items=['cep', 'latitude', 'longitude', 'logradouro', 'bairro', 'cidade', 'estado'])
dist = pd.Series()

for i in list(tabela.index.values):
    try:
        ponto_cep = (tabela['latitude'][i], tabela['longitude'][i])
        tab = pd.Series(geodesic(referencia, ponto_cep).meters)
        dist = dist.append(tab)
    except:
        dist = dist.append(pd.Series('NA'))

dist = pd.DataFrame(dist)
dist.columns = ['Distancia']
tabela = pd.concat([tabela.reset_index(drop=True), dist.reset_index(drop=True)], axis=1)

tabela = tabela.drop_duplicates(['cep'])

tabela.to_csv('tabela_sem_dupli.csv',)

