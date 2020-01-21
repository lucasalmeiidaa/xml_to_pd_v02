#!/usr/bin/env python
# coding: utf-8

# In[2]:


import xmltodict
import pandas as pd
from os.path import join
from os import listdir
import numpy as np
# with open('bases/13190504501136000136550010009435231024529034_9e11f85798f4e4bc8e8425a2274a217b866de706.xml') as fd:
#     doc = xmltodict.parse(fd.read())


# In[3]:


def busca_xml_v2():
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
    
    df = pd.DataFrame(data=lst_dados, columns=['Cód.', 'Desc.', 
                                     'NCM', 'Embalagem', 'CST Orig.'])
    
    return df


# In[6]:


path_xml = 'bases'

files_xml = []

for file in listdir(path_xml):
    files_xml.append(file)


# In[7]:


df = pd.DataFrame()

writer = pd.ExcelWriter('produtos.xlsx', engine='xlsxwriter')

for file in files_xml:
    # Abrindo xml
    with open(join(path_xml, file)) as fd:
        doc = xmltodict.parse(fd.read())
    
    df_aux = busca_xml_v2()
   
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

