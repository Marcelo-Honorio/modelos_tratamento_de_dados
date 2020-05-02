# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 11:07:27 2020

@author: marce
"""
import pandas as pd
import numpy as np
import requests
import os
import re

path = 'C:\\Users\\marce\\Documents\\Emovel\\Dados\\dados_hugo\\dados_hugo_novo'
os.chdir(path)

api_eemovel = pd.read_excel("anuncios.xlsx")

api_eemovel = api_eemovel[['broker_raw','contact_names','category_name','contact_emails','features_suite','listing_title','state_name','sub_category_name','condominium_value','city_raw','contact_phones','area','neighborhood_name','city_uf','transaction_rent','city_id','iptu_value','amenities','features_bedroom','total_area','link','created_at','description','state_raw','city_name','original_address_neighborhood','features_garage','features_bathroom','processed_address_number','processed_address_street','state_uf','location','location_point','transaction_sale']]

grupo_uf = api_eemovel.city_uf.drop_duplicates()

grupo_uf = list(grupo_uf)

for i in grupo_uf:
    dados = api_eemovel[api_eemovel['city_uf'] == i]
    dados.to_excel('dados_' + i + '.xlsx', index=False)
    
    

