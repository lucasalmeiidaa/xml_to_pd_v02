#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xmltodict
import pandas as pd
from os.path import join
from os import listdir
import numpy as np
# with open('bases/13190504501136000136550010009435231024529034_9e11f85798f4e4bc8e8425a2274a217b866de706.xml') as fd:
#     doc = xmltodict.parse(fd.read())


# In[2]:


def busca_xml():
    lst_dados = []
    for nf in doc['nfeProc']['NFe']['infNFe']['det']:
        lst_aux = []
        try:
            cod_prod = nf['prod']['cProd']
        except:
            cod_prod = np.nan
        try:
            desc_prod = nf['prod']['xProd']
        except:
            desc_prod = np.nan
        try:
            nom = nf['prod']['NCM']
        except:
            nom = np.nan
        try:
            tipo = nf['prod']['uCom']
        except:
            tipo = np.nan
        try:    
            qtd = nf['prod']['qCom']
        except:
            qtd = np.nan
        try:
            valor_unit = nf['prod']['vUnCom']
        except:
            valor_unit = np.nan
        try:
            valor_total = nf['prod']['vProd']
        except:
            valor_total = np.nan
        try:
            desc_total = nf['prod']['vDesc']
        except:
            desc_total = np.nan
            
        lst_aux.append(cod_prod)
        lst_aux.append(desc_prod)
        lst_aux.append(nom)
        lst_aux.append(tipo)
        lst_aux.append(qtd)
        lst_aux.append(valor_unit)
        lst_aux.append(valor_total)
        lst_aux.append(desc_total)
        lst_dados.append(lst_aux)
    
    df = pd.DataFrame(data=lst_dados, columns=['cod_produto', 'descricao_prod', 
                                     'nomenclatura', 'tipo', 'quantidade', 
                                     'valor_unit', 'valor_total',
                                     'desc_total'])
    
    return df


# In[25]:


doc['nfeProc']['NFe']['infNFe']['emit']['xNome']


# In[52]:


doc['nfeProc']['NFe']['infNFe']['emit']['CNPJ']


# In[61]:


def busca_xml_v2(doc):
    lst_dados = []      
    
    for nf in doc['nfeProc']['NFe']['infNFe']['det']:
        lst_aux = []
        try:
            cod = nf['prod']['cProd']
        except:
            cod = np.nan
        try:
            desc = nf['prod']['xProd']
        except:
            desc = np.nan
        try:
            ncm = nf['prod']['NCM']
        except:
            ncm = np.nan
        try:
            embalagem = nf['prod']['uCom']
        except:
            embalagem = np.nan
        
        try:
#             try:
            orig = nf['imposto']['ICMS']['ICMS60']['orig']
#             except:
#                 pass
        except:
            try:
                orig = nf['imposto']['ICMS']['ICMS00']['orig']
            except:
                try:
                    orig = nf['imposto']['ICMS']['ICMS60']
                except:
                    try:
                        orig = nf['imposto']['ICMS']['ICMSSN102']['orig']
                    except:
                        try:
                            orig = nf['imposto']['ICMS']['ICMSSN101']['orig']
                        except:
                            try:
                                orig = nf['imposto']['ICMS']['ICMSSN500']['orig']
                            except:
                                orig = np.nan
        
        try:
            cnpj = info['CNPJ']
        except:
            cnpj = np.nan
        try:
            razao_social = info['xNome']
        except:
            razao_social = np.nan

        
        lst_aux.append(cod)
        lst_aux.append(desc)
        lst_aux.append(ncm)
        lst_aux.append(embalagem)
#         lst_aux.append(qtd)
#         lst_aux.append(valor_unit)
#         lst_aux.append(valor_total)
#         lst_aux.append(desc_total)
        lst_aux.append(orig)
        lst_dados.append(lst_aux)
    
    df = pd.DataFrame(data=lst_dados, columns=['Cód.', 'Desc.', 'NCM', 'Embalagem', 'CST Orig.'])
    
    try:
        cnpj =  doc['nfeProc']['NFe']['infNFe']['emit']['CNPJ']
    except:
        cnpj = np.nan
    try:
        razao_social = doc['nfeProc']['NFe']['infNFe']['emit']['xNome']
    except:
        razao_social = np.nan
    
    df['CNPJ'] = cnpj
    df['razão social'] = razao_social
    
    return df


# In[62]:


path_xml = 'bases'

files_xml = []

for file in listdir(path_xml):
    files_xml.append(file)


# In[63]:


files_xml


# In[64]:


df = pd.DataFrame()

writer = pd.ExcelWriter('produtos.xlsx', engine='xlsxwriter')

for file in files_xml:
    # Abrindo xml
    with open(join(path_xml, file)) as fd:
        doc = xmltodict.parse(fd.read())
    
    df_aux = busca_xml_v2(doc)
   
    # Transformando colunas em float
#     df['desc_total'] = df['desc_total'].astype(float)
#     df['quantidade'] = df['quantidade'].astype(float)
#     df['valor_total'] = df['valor_total'].astype(float)
    
#     # Fazendo operações de cálculo
#     df['Coluna1'] = df['desc_total'] / df['quantidade']
#     df['Coluna2'] = df['valor_total'] - df['desc_total']
    
    # Exportando para excel
    df = pd.concat([df, df_aux])
    df.to_excel(writer, index=False)
    
    
writer.save()


# In[4]:


files_xml


# In[58]:


with open(join(path_xml, '13191104501136000136550020000037021036310896.xml')) as fd:
    doc = xmltodict.parse(fd.read())


# In[60]:


doc['nfeProc']['NFe']['infNFe']['emit']['CNPJ']


# In[ ]:


lst_aux = []
   try:
       cod = nf['prod']['cProd']
   except:
       cod = np.nan
   try:
       desc = nf['prod']['xProd']
   except:
       desc = np.nan
   try:
       ncm = nf['prod']['NCM']
   except:
       ncm = np.nan
   try:
       embalagem = nf['prod']['uCom']
   except:
       embalagem = np.nan

   try:
#             try:
       orig = nf['imposto']['ICMS']['ICMS60']['orig']
#             except:
#                 pass
   except:
       try:
           orig = nf['imposto']['ICMS']['ICMS00']['orig']
       except:
           try:
               orig = nf['imposto']['ICMS']['ICMS60']
           except:
               try:
                   orig = nf['imposto']['ICMS']['ICMSSN102']['orig']
               except:
                   try:
                       orig = nf['imposto']['ICMS']['ICMSSN101']['orig']
                   except:
                       try:
                           orig = nf['imposto']['ICMS']['ICMSSN500']['orig']
                       except:
                           orig = np.nan


   lst_aux.append(cod)
   lst_aux.append(desc)
   lst_aux.append(ncm)
   lst_aux.append(embalagem)


# In[6]:


# writer = pd.ExcelWriter('saída.xlsx')

# for df in lista_dfs:
#     df.to_excel(writer)


# In[11]:


# doc['nfeProc']['NFe']['infNFe']['det'][4]


# In[10]:


doc['nfeProc']['NFe']['infNFe']['det'][1]['imposto']['ICMS']['ICMS60']['orig']

