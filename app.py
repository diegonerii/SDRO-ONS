import os, os.path
import pandas as pd
from funcoes import *

lista_arquivos = [arquivo[7:-5] for arquivo in os.listdir('files/')]
# lista_arquivos = ['01-01-2017', '02-01-2017']

for files in lista_arquivos:
    """
    O uso do try-except é necessário em virtude de alguns arquivos na pasta files/ poderem estar corrompidos
    """
    try:
        Inicializacao(files)
        BalancoEnergetico(files)
        DespachoTermico(files)
        ENA(files)
        EnergiaArmazenada(files)
    except AttributeError as erro:
        print(f'O dia {files} veio com o arquivo corrompido. Erro encontrado: {erro}.')
        continue

# Define duas listas: a primeira receberá todos os dicionários em estado bruto e a segunda receberá os dicionários transformados
lista_impressao, lista_df = [], []

# Definição do nome final das planilhas
abas = ['01-Balanço de Energia', '12-Motivo do Despacho Térmico', '19-Energia Natural Afluente', '20-Variação Energia Armazenada']

# Criar um objeto ExcelWriter
writer = pd.ExcelWriter('DIARIO - Final.xlsx', engine='xlsxwriter')

# Apenda os dicionários na primeira lista
lista_impressao.append(be_dicionario_campos)
lista_impressao.append(dt_dicionario_campos)
lista_impressao.append(ena_dicionario_campos)
lista_impressao.append(ea_dicionario_campos)

# Transforma cada dicionário
for dicionario in lista_impressao:
    workbook = writer.book
    df = pd.DataFrame.from_dict(dicionario, orient='index')
    df = df.transpose()
    lista_df.append(df)

# Salva cada dicionário em uma planilha de uma pasta Excel
for dataframe in range(0, len(lista_df)):
    lista_df[dataframe].to_excel(
        writer, 
        sheet_name=str(abas[dataframe]),
        startrow=0,
        startcol=0,
        index=False
    )

# Salve o arquivo Excel e fecha o objeto ExcelWriter
writer.save()
