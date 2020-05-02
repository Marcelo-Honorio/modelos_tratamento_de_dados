import requests
import pandas as pd
import time

token = "060ed6306e2fa7b7be63c72024afbaef"
headers = {'Authorization': 'Token token=%s' % token}
j = int(0)
cep_vazio = []
todos_cep = []
tabela = pd.DataFrame()

CEP_CAS = [29010330, 85818430, 85819100, 85818430, 85819100, 85818430]

for i in CEP_CAS:
    cep = str(i)
    def search_by_cep():
        url = "http://www.cepaberto.com/api/v3/cep?cep=" + cep
        response = requests.get(url, headers=headers)
        return response.json()
    try:
        tab = pd.DataFrame(search_by_cep())
        tabela = tabela.append(tab)            
    except:
        j += 1
        cep_vazio.append(i)
        vazio = {'%s CEP Vazio' %j}
        print(vazio)
    time.sleep(5)

tabela = tabela

#tabela.to_csv('tabela.csv')