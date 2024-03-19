import pandas as pd
import numpy as np
import warnings
from rapidfuzz import process, fuzz

def match_name(name, list_names, min_score=0):
    # Use the fuzz.token_set_ratio function as the scorer
    return process.extractOne(name, list_names, scorer=fuzz.token_set_ratio, score_cutoff=min_score)

def match_index(name, list_names, min_score=0):
    # Inicialize a melhor pontuação e o melhor índice
    best_score = -1
    best_index = -1

    # Itere sobre list_names
    for i, list_name in enumerate(list_names):
        # Calcule a pontuação
        score = fuzz.token_set_ratio(name, list_name)

        # Se a pontuação for maior que a melhor pontuação e min_score, atualize a melhor pontuação e o melhor índice
        if score > best_score and score >= min_score:
            best_score = score
            best_index = i

    # Se a melhor pontuação é -1, nenhuma correspondência foi encontrada, então retorne None
    if best_score == -1:
        return None

    # Retorne o melhor índice, a melhor correspondência e a melhor pontuação
    return best_index, list_names[best_index], best_score

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
# Itere sobre as linhas do df2
for i, empreendimento in df2['EMPREENDIMENTO'].items():
    print(x, end=', ')
    # Se o valor da coluna "EMPREENDIMENTO" for uma string vazia, pule para a próxima linha
    if pd.isna(empreendimento):
        x += 1
        continue

    # Use rapidfuzz to find matching empreendimento in df_principal
    match = match_name(empreendimento, df_principal['Nome/Denominação do Objeto'], 80)
    if match:
        matching_row = df_principal[df_principal['Nome/Denominação do Objeto'] == match[0]]
    else:
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
    x += 1    


# Salve o df2 como um novo arquivo xlsx
df2.to_excel('IBAMA.xlsx', index=False)


# Carregar o arquivo de saída do programa main.py
df1 = pd.read_excel('Ibama.xlsx')

# Carregar o outro arquivo xlsx
df2 = pd.read_excel('Planilha_2_todos_processos_SISLIC1.xlsx')

# Convert 'Empreendimento' to string
df2['Empreendimento'] = df2['Empreendimento'].astype(str)
df2['Empreendimento'].rename('EMPREENDIMENTO', inplace=True)
# Use a função match_name para encontrar as correspondências
matches = df1['EMPREENDIMENTO'].apply(lambda x: match_name(x, df2['Empreendimento'], 80))

# Filtrar as correspondências para remover None
matches = matches[matches.notnull()]

# Crie uma lista de tuplas onde cada tupla contém o índice da linha que deu match em df1 e df2
indices = [(index, match[0]) for index, match in matches.items()]
print(indices)

for i in range(len(indices)):
    if indices[i][1] in df2.index:
        df1.loc[indices[indices[i][0]], 'Nº PROCESSO'] = df2.loc[indices[i][1], 'Nr Processo']
    else:
        print(f"Índice {indices[i][1]} não encontrado em df2")
df1.to_excel('IBAMA_NEW1.xlsx', index=False)
