# -*- coding: utf-8 -*-
import requests
import pandas as pd
import numpy as np

endereco = []

CEP_SP = pd.read_csv("cep_sp_centro.csv", names=None, index_col=None)

cep_centro = []

for i in CEP_SP['CEP']:
    cep_centro.append(i)

print(CEP_SP)

token = "060ed6306e2fa7b7be63c72024afbaef"
headers = {'Authorization': 'Token token=%s' % token}
j = int(0)
cep_vazio = []
tabela = pd.DataFrame()

CEP_CAS = [85818430, 85819100]

for i in cep_centro:
    cep = str(i)
    def search_by_cep():
        url = "http://www.cepaberto.com/api/v3/cep?cep=" + cep
        response = requests.get(url, headers=headers)
        return response.json()
    try:
       endereco = endereco.append(np.array(search_by_cep()))
    except:
        j += 1
        cep_vazio.append(i)
        vazio = {'%s CEP Vazio' %j}
        print(vazio)


