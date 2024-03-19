import re
import numpy as np
import requests
from PyPDF2 import PdfFileReader
from io import BytesIO
from selenium import webdriver
import openpyxl
import os
import pandas as pd
import warnings

# Carregue as planilhas
xls = pd.ExcelFile('tabela1.xlsx')
df_principal = pd.read_excel(xls, 'Principal', header=2)  # os nomes das colunas estão na linha 3
df_grd = pd.read_excel(xls, 'GRD_OBJ_LOCAIS')
df_bacias = pd.read_excel(xls, 'GRD_BACIAS')

# Carregue o segundo arquivo xlsx
df2 = pd.read_excel('tabela2.xlsx')

# Crie a nova coluna "BACIAS" no df2
df2['BACIAS'] = np.nan

x = 0

for i, empreendimento in df2['EMPREENDIMENTO'].items():
    print(x, end=':')
    # Se o valor da coluna "EMPREENDIMENTO" for uma string vazia, pule para a próxima linha
    if pd.isna(empreendimento):
        x += 1
        continue

    # Se o valor da coluna "ESTADO" for uma string vazia, faça uma busca
    if pd.isna(df2.loc[i, 'ESTADO']):
        # Encontre o valor correspondente nas colunas "Nome/Denominação do Objeto" e "Informações do Processo SEI informado:" da planilha "Principal"
        "matching_row = df_principal[(df_principal['Nome/Denominação do Objeto'] == empreendimento) | (df_principal['Informações do Processo SEI informado:'].str.contains((empreendimento), na=False))]"
        # Ignore UserWarning
        warnings.filterwarnings('ignore', category=UserWarning)
        warnings.filterwarnings('ignore', 'This pattern is interpreted as a regular expression.*', category=UserWarning)
        # Your code
        matches = df_principal.iloc[:, 1:].astype(str).apply(lambda col: col.str.contains(empreendimento, na=False, regex=False))
        matching_row = df_principal[matches.any(axis=1)]
        # Verifica a Tipologia  e se for Petróleo e Gás - Perfuração, Petróleo e Gás - Pesquisa Sísmica ou Petróleo e Gás - Produção, busca na planilha "GRD_BACIAS"        
        tipologia = df2.loc[i, 'TIPOLOGIA']
        if tipologia == 'Petróleo e Gás - Perfuração' or tipologia == 'Petróleo e Gás - Pesquisa Sísmica' or tipologia == 'Petróleo e Gás - Produção':
            if not matching_row.empty:
                matching_bacia = df_bacias[df_bacias['#Processo'] == matching_row['#Processo'].values[0]]
                if not matching_bacia.empty:
                    df2.loc[i, 'BACIAS'] = matching_bacia['Bacia Sedimentar'].values[0]  # Supondo que 'BACIAS' é uma coluna em df_bacias
                    print('Bacia = ' + matching_bacia['Bacia Sedimentar'].values[0])
                    x += 1
            continue
        else:
            # Extraia o valor da coluna "#Processo"
            if not matching_row.empty:
                processso_value = matching_row['#Processo'].values[0]
                
                # Encontre as linhas correspondentes na planilha "GRD_OBJ_LOCAIS"
                matching_rows = df_grd[df_grd['#Processo'] == processso_value]

                # Extraia os valores das colunas "UF" e "Município"
                uf_values = matching_rows['UF'].values
                municipio_values = matching_rows['Município'].values

                # Adicione esses valores ao df2
                if uf_values.size > 0 and municipio_values.size > 0:
                    if uf_values.size > 1:
                        df2.loc[i, 'ESTADO'] = ', '.join(uf_values)
                        df2.loc[i, 'MUNICIPIO'] = ', '.join(municipio_values)
                        print('UF = ' + ', '.join(uf_values))
                        x += 1
                    else:
                        df2.loc[i, 'ESTADO'] = uf_values[0]
                        df2.loc[i, 'MUNICIPIO'] = municipio_values[0]
                        print('UF = ' + uf_values[0])
                        x += 1

# Salve o df2 como um novo arquivo xlsx
df2.to_excel('Test2.xlsx', index=False)