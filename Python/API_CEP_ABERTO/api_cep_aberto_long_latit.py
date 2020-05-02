# -*- coding: utf-8 -*-
import requests
import pandas as pd
import time
from geopy.distance import geodesic

CEP_SP = pd.read_csv("CEP_SP.csv")

CEP_cas = CEP_SP


CEP_SP = CEP_SP[(CEP_SP['Bairro'] == 'Alto de Pinheiros')]
referencia = ()


CEP_CAS = []

for i in CEP_SP['CEP']:
    CEP_CAS.append(str(i))

CEP_CORIG = []

for i in CEP_CAS:
    if len(i) < 8:
        i = "0" + i
        CEP_CORIG.append(i)

token = "060ed6306e2fa7b7be63c72024afbaef"
headers = {'Authorization': 'Token token=%s' % token}
j = int(0)
cep_vazio = []
todos_cep = []
tabela = pd.DataFrame()
n = int(0)

for cep in CEP_CORIG:
    def search_by_cep():
        url = "http://www.cepaberto.com/api/v3/cep?cep=" + cep
        response = requests.get(url, headers=headers)
        return response.json()
    try:
        n += 1
        tab = pd.DataFrame(search_by_cep())
        tabela = tabela.append(tab)            
    except:
        j += 1
        cep_vazio.append(i)
        vazio = {'%s CEP Vazio' %j}
        print(vazio)
    time.sleep(4)

#Tratar a tabela

tabela = tabela.filter(items=['cep', 'latitude', 'longitude', 'logradouro', 'bairro'])
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

tabela.to_csv('tabela_sem_dupli.csv')

#tabela = tabela.filter(items=['cep', 'latitude', 'longitude', 'logradouro', 'bairro', 'cidade', 'estado'])
#tabela = tabela.drop_duplicates(['cep'])