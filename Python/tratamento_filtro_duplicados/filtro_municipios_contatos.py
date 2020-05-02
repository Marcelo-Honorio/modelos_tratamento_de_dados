# -*- coding: utf-8 -*-
"""
Created on Thu Mar  5 08:44:00 2020

@author: marce
"""
import pandas as pd
import numpy as np
import requests
import os
import re

path = 'C:\\Users\\marce\\Documents\\Emovel\\Dados\\dados_hugo\\dados_hugo_novo\\RJ'
os.chdir(path)

dados_rj = pd.read_excel('dados_RJ.xlsx')

cidades = list(dados_rj['city_raw'].drop_duplicates())

for i in cidades:
    dados = dados_rj[dados_rj['city_raw'] == i]
    i = re.sub('Ã¡', 'a', i)
    i = re.sub('Ã£', 'a', i)
    i = re.sub('Ã©', 'e', i)
    i = re.sub('Ã\xad', 'i', i)
    i = re.sub('Ã¡', 'a', i)
    i = re.sub('Ã¼', 'u',i)
    i = re.sub('Ã§', 'c', i)
    i = re.sub('Ãº', 'u', i)
    i = re.sub('Ã¢', 'a', i)
    i = re.sub('Ã\xad', 'i', i)
    i = re.sub(',', '-', i)
    i = re.sub("'", '_', i)
    #Salvando
    dados.to_excel(i + '.xlsx', index=False)






for i in range(len(cidades)):
    cidades[i] = re.sub('Ã¡', 'a', cidades[i])
    cidades[i] = re.sub('Ã£', 'a', cidades[i])
    cidades[i] = re.sub('Ã©', 'e', cidades[i])
    cidades[i] = re.sub('Ã\xad', 'i', cidades[i])
    cidades[i] = re.sub('Ã¡', 'a', cidades[i])
    cidades[i] = re.sub('Ã¼', 'u', cidades[i])
    cidades[i] = re.sub('Ã§', 'c', cidades[i])
    cidades[i] = re.sub('Ãº', 'u', cidades[i])
    cidades[i] = re.sub('Ã¢', 'a', cidades[i])
    cidades[i] = re.sub('Ã\xad', 'i', cidades[i])
    cidades[i] = re.sub(',', '-', cidades[i])
    cidades[i] = re.sub("'", '_', cidades[i])




    







