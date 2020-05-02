# -*- coding: utf-8 -*-
"""
Created on Mon Feb 17 11:06:14 2020

@author: marce
"""
import pandas as pd
import numpy as np
import requests
import os
import re

# DETERMINANDO DIRETORIO
os.getcwd()
path = 'C:\\Users\\marce\\Documents\\Emovel\\Dados\\composicoes\\composicao_4'
os.chdir(path)

##############################################################################
# LISTA DE ARQUIVOS PARA LER
# arquivos = [i for i in os.listdir() if re.search('\\b.xlsx\\b', i, re.IGNORECASE)]

##############################################################################

# LENDO ARQUIVO        
api_eemovel = pd.read_csv("response_4.csv", sep = ';')

#LISTA DE NOMES SEM CPF E CNPJ
api_eemovel_na = api_eemovel.fillna(0).query('CPF_COMPLETO == 0').drop_duplicates('NOME')

if len(list(api_eemovel_na['NOME'])) > 0:
    api_eemovel_na = api_eemovel_na[['NOME']]
    api_eemovel_na.to_excel('nome_sem_cpf.xlsx', index=False)

#NOMES PARA ASSERTIVA
api_eemovel = api_eemovel.query('CPF_COMPLETO != 0').sort_values('NOME')

def cpf_completo(x):
    if len(str(int(x))) == 5:
        return '000000' + str(int(x))
    elif len(str(int(x))) == 6:
        return '00000' + str(int(x))
    elif len(str(int(x))) == 7:
        return '0000' + str(int(x))
    elif len(str(int(x))) == 8:
        return '000' + str(int(x))
    elif len(str(int(x))) == 9:
        return '00' + str(int(x))
    elif len(str(int(x))) == 10:
        return '0' + str(int(x))
    else:
        return str(int(x))

# CRIANDO COLUNA
api_eemovel['CPF_CORRIGIDO'] = [cpf_completo(x) for x in api_eemovel['CPF_COMPLETO']]

#CRIANDO UMA LISTA DE DUPLICADOS
api_eemovel = api_eemovel.drop_duplicates(['NOME', 'CPF_CORRIGIDO'])
grupo_nomes_repetido = api_eemovel[['NOME', 'CPF_COMPLETO']].groupby('NOME').count()
nomes_repetido = list()
for i in range(len(grupo_nomes_repetido)):
    if grupo_nomes_repetido['CPF_COMPLETO'][i] > 1:
        nomes_repetido.append(grupo_nomes_repetido.index[i])

#CRIANDO TABELA DUPLICADOS
if len(nomes_repetido) > 0:
    nomes_duplicados = api_eemovel[api_eemovel.NOME.isin(nomes_repetido)]
    nomes_duplicados = nomes_duplicados.drop_duplicates('CPF_CORRIGIDO')
    nomes_duplicados.to_excel('nomes_duplicados.xlsx', index=False)
    lista_duplicados = list(nomes_duplicados['NOME'])
    nomes_unicos = api_eemovel.drop_duplicates(['NOME', 'CPF_CORRIGIDO'])
    
    for i in lista_duplicados:
        nomes_unicos = nomes_unicos[nomes_unicos['NOME'] != i]
    nomes_unicos.to_excel('nomes_unicos.xlsx', index=False)
elif len(nomes_repetido) == 0:   
    #CRIANDO TABELA NOMES UNICOS
    nomes_unicos = api_eemovel.drop_duplicates(['NOME', 'CPF_CORRIGIDO'])
    nomes_unicos.to_excel('nomes_unicos.xlsx', index=False)

#SUBINDO TABELA DE DADOS BASE


#UNIR AS TABELA

#pd.merge(tabela1, tabela2, how = 'innner', on='chave')

#lista_tabela = list()
#for x in arquivos:
#    lista_tabela.append(pd.read_excel(x))


