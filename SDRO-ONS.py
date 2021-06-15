import requests
import urllib
import pandas as pd
from datetime import datetime, timedelta
import os.path
import xlrd
import time
import re


tempo_ini = time.time()

url_exemplo = r'http://sdro.ons.org.br/SDRO/DIARIO/2021_05_31/Html/DIARIO_31-05-2021.xlsx' # URL_Exemplo
data = pd.ExcelFile(url_exemplo)
abas = data.sheet_names

url_pt_1 = r'http://sdro.ons.org.br/SDRO/DIARIO/'


########## BAIXAR TODOS OS ARQUIVOS ##########
data_inicial = '01-01-2017'
data_inicial = datetime.strptime(data_inicial, '%d-%m-%Y')

print("Verificando Existência de Novo Arquivo...")
while data_inicial <= datetime.today():
    ano = data_inicial.year
    mes = data_inicial.month
    dia = data_inicial.day

    url_pt_2 = str(ano) + "_" + str(mes).zfill(2) + "_" + str(dia).zfill(2) + "/Html/DIARIO_" + \
               str(dia).zfill(2) + "-" + str(mes).zfill(2) + "-" + str(ano) + ".xlsx"
    url = url_pt_1 + url_pt_2
    data = str(dia) + "-" + str(mes) + "-" + str(ano)

    if os.path.exists('C:/Users/GN6006/Documents/Teste_ONS/' + str(url[-22:])):
        #print("ARQUIVO EXISTE")
        pass
    else:
        #print("ARQUIVO NÃO EXISTE")
        if str(requests.get(url)) == '<Response [200]>':
            urllib.request.urlretrieve(url, 'C:/Users/GN6006/Documents/Teste_ONS/'+str(url[-22:]))
            print(data)
            print("Arquivo Baixado")

    data_inicial = data_inicial + timedelta(days=1)


######################## MANIPULAÇÃO DE ARQUIVO ########################

lista_impressao = []

dicionario_campos = {}
be_dicionario_campos = {}
dt_dicionario_campos = {}
ena_dicionario_campos = {}
ea_dicionario_campos = {}

######################## ABA BALANÇO DE ENERGIA (BE) ########################

carga_programada, carga_verificada = [], []
producao_hidro_programada, producao_hidro_verificada = [], []
producao_itaipu_programada, producao_itaipu_verificada = [], []
producao_nuclear_programada, producao_nuclear_verificada = [], []
producao_termo_programada, producao_termo_verificada = [], []
producao_eolica_programada, producao_eolica_verificada = [], []
producao_solar_programada, producao_solar_verificada = [], []
exportacao_programada, exportacao_verificada = [], []

listas_norte, carga_norte_programada, carga_norte_verificada = [], [], []
listas_nordeste, carga_nordeste_programada, carga_nordeste_verificada = [], [], []
listas_sudeste, carga_sudeste_programada, carga_sudeste_verificada = [], [], []
listas_sul, carga_sul_programada, carga_sul_verificada = [], [], []

be_data_arquivo = []

######################## ABA DESPACHO TÉRMICO (DT) ########################

dt_usina, dt_cod_ONS, dt_potencia, dt_ordem_de_merito, dt_inflex, dt_restricao = [], [], [], [], [], []
dt_geracao_fora_da_ordem, dt_energia_de_reposicao, dt_garantia_energetica = [], [], []
dt_exportacao, dt_verificado, data_arquivo = [], [], []

######################## ABA ENA (ENA) ########################

ena_subsistema, ena_percent_mlt_dia, ena_percent_mlt_acum, ena_bruta_dia, ena_bruta_acum = [], [], [], [], []
ena_data_arquivo = []

######################## ABA ENERGIA ARMAZENADA (EA) ########################

ea, ea_sul, ea_sudeste, ea_nordeste, ea_norte = [], [], [], [], []
ea_data_arquivo = []

######################## DEFINIÇÕES DE MÉTODOS ########################

def definicao_planilha(antigo_nome, nome_da_planilha, novo_nome, numero_da_aba, lista):
    numero_da_aba = int(numero_da_aba)
    antigo_nome = antigo_nome[2:]
    try:
        try:
            novo_nome = nome_da_planilha.sheet_by_name(str(numero_da_aba).zfill(2)+antigo_nome)
        except:
            try:
                novo_nome = nome_da_planilha.sheet_by_name(str(numero_da_aba+1).zfill(2)+antigo_nome)
            except:
                try:
                    novo_nome = nome_da_planilha.sheet_by_name(str(numero_da_aba + 2).zfill(2)+antigo_nome)
                except:
                    novo_nome = nome_da_planilha.sheet_by_name(str(numero_da_aba + 3).zfill(2)+antigo_nome)
    except:
        pass
    lista.append(novo_nome)


def dt_data(aba, campo):

    linha_e_coluna = []
    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if re.findall(campo, str(aba.cell(linha, coluna).value)) != []:
                linha_e_coluna.append(linha)
                linha_e_coluna.append(coluna)
    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    for scan_linha in range(linha+1, aba.nrows):

        data_arquivo.append(date)

    dt_dicionario_campos["Data"] = data_arquivo


def dt_campos(aba, campo, lista, titulo_coluna):

    linha_e_coluna = []
    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if re.findall(campo, str(aba.cell(linha, coluna).value)) != []:
                linha_e_coluna.append(linha)
                linha_e_coluna.append(coluna)
    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    for scan_linha in range(linha+1, aba.nrows):
        valor = str(aba.cell(scan_linha, coluna).value).strip()
        if re.findall(r'[0-9]+\.[0-9]', valor) != []:
            valor = round(float(aba.cell(scan_linha, coluna).value), 1)
            valor = str(valor).replace(".", ",")
            lista.append(valor)
        else:
            lista.append(str(aba.cell(scan_linha, coluna).value).strip())

    dt_dicionario_campos[titulo_coluna] = lista


def be_busca_campo_pelo_nome(aba, submercado, tipo_de_energia, range_coluna, lista, nome_coluna, range_linha):

    linha_e_coluna = []


    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if re.findall(submercado, str(aba.cell(linha, coluna).value)) != []:
                linha_e_coluna.append(linha)
                linha_e_coluna.append(coluna)
    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    vetor_regex = []
    for scan_linha in range(linha, linha+range_linha):
        if re.findall(tipo_de_energia, str(aba.cell(scan_linha, coluna).value)) != []:
            valor = aba.cell(scan_linha, coluna+range_coluna).value
            if valor != '':
                valor = round(aba.cell(scan_linha, coluna+range_coluna).value, 0)
                vetor_regex.append(valor)
            else:
                vetor_regex.append(0)
        else:
            vetor_regex.append(0)
    if min(vetor_regex) >= 0:
        lista.append(max(vetor_regex))
    else:
        lista.append(min(vetor_regex))

    #dicionario_campos[nome_coluna] = lista
    be_dicionario_campos[nome_coluna] = lista


def be_busca_regioes(regiao_maiuscula, regiao_minuscula):

    tipo_de_energia = ['_hidro_', '_termo_', '_nuclear_', '_eolica_', '_solar_', '_exportacao_']
    programada_e_verificada = ['programada', 'verificada']
    lista_tipo_de_energia = []

    for tipo in range(0, len(tipo_de_energia)):
        for pv in range(0, len(programada_e_verificada)):
            name = "producao_"+regiao_minuscula+tipo_de_energia[tipo]+programada_e_verificada[pv]
            name = []

            if regiao_maiuscula == 'Norte':
                listas_norte.append(name)
            elif regiao_maiuscula == 'Nordeste':
                listas_nordeste.append(name)
            elif regiao_maiuscula == 'Sudeste / Centro-Oeste':
                listas_sudeste.append(name)
            elif regiao_maiuscula == 'Sul':
                listas_sul.append(name)

    if regiao_maiuscula == 'Norte':
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 3,
                              listas_norte[0], regiao_maiuscula + ' - Hidro Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 2,
                              listas_norte[1], regiao_maiuscula + ' - Hidro Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 3,
                              listas_norte[2], regiao_maiuscula + ' - Termo Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 2,
                              listas_norte[3], regiao_maiuscula + ' - Termo Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 3,
                              listas_norte[4], regiao_maiuscula + ' - Nuclear Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 2,
                              listas_norte[5], regiao_maiuscula + ' - Nuclear Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 3,
                              listas_norte[6], regiao_maiuscula + ' - Eólica Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 2,
                              listas_norte[7], regiao_maiuscula + ' - Eólica Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 3,
                              listas_norte[8], regiao_maiuscula + ' - Solar Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 2,
                              listas_norte[9], regiao_maiuscula + ' - Solar Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 3,
                              carga_norte_programada, regiao_maiuscula + ' - Carga Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 2,
                              carga_norte_verificada, regiao_maiuscula + ' - Carga Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Exportação', 3,
                              listas_norte[10], regiao_maiuscula + ' - Exportação Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Exportação', 2,
                              listas_norte[11], regiao_maiuscula + ' - Exportação Verif', 11)


    elif regiao_maiuscula == 'Nordeste':
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 3,
                              listas_nordeste[0], regiao_maiuscula + ' - Hidro Prog', 9)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 2,
                              listas_nordeste[1], regiao_maiuscula + ' - Hidro Verif', 9)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 3,
                              listas_nordeste[2], regiao_maiuscula + ' - Termo Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 2,
                              listas_nordeste[3], regiao_maiuscula + ' - Termo Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 3,
                              listas_nordeste[4], regiao_maiuscula + ' - Nuclear Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 2,
                              listas_nordeste[5], regiao_maiuscula + ' - Nuclear Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 3,
                              listas_nordeste[6], regiao_maiuscula + ' - Eólica Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 2,
                              listas_nordeste[7], regiao_maiuscula + ' - Eólica Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 3,
                              listas_nordeste[8], regiao_maiuscula + ' - Solar Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 2,
                              listas_nordeste[9], regiao_maiuscula + ' - Solar Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 3,
                              carga_nordeste_programada, regiao_maiuscula + ' - Carga Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 2,
                              carga_nordeste_verificada, regiao_maiuscula + ' - Carga Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Importação', 3,
                              listas_nordeste[10], regiao_maiuscula + ' - Exportação Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Importação', 2,
                              listas_nordeste[11], regiao_maiuscula + ' - Exportação Verif', 11)


    elif regiao_maiuscula == 'Sudeste / Centro-Oeste':
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 3,
                              listas_sudeste[0], regiao_maiuscula + ' - Hidro Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 2,
                              listas_sudeste[1], regiao_maiuscula + ' - Hidro Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 3,
                              listas_sudeste[2], regiao_maiuscula + ' - Termo Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 2,
                              listas_sudeste[3], regiao_maiuscula + ' - Termo Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 3,
                              listas_sudeste[4], regiao_maiuscula + ' - Nuclear Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 2,
                              listas_sudeste[5], regiao_maiuscula + ' - Nuclear Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 3,
                              listas_sudeste[6], regiao_maiuscula + ' - Eólica Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 2,
                              listas_sudeste[7], regiao_maiuscula + ' - Eólica Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 3,
                              listas_sudeste[8], regiao_maiuscula + ' - Solar Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 2,
                              listas_sudeste[9], regiao_maiuscula + ' - Solar Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 3,
                              carga_sudeste_programada, regiao_maiuscula + ' - Carga Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 2,
                              carga_sudeste_verificada, regiao_maiuscula + ' - Carga Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Intercâmbio', 3,
                              listas_sudeste[10], regiao_maiuscula + ' - Exportação Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Intercâmbio', 2,
                              listas_sudeste[11], regiao_maiuscula + ' - Exportação Verif', 11)


    elif regiao_maiuscula == 'Sul':
        listas_sul.append(name)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 3,
                              listas_sul[0], regiao_maiuscula + ' - Hidro Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Hidro', 2,
                              listas_sul[1], regiao_maiuscula + ' - Hidro Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 3,
                              listas_sul[2], regiao_maiuscula + ' - Termo Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Termo', 2,
                              listas_sul[3], regiao_maiuscula + ' - Termo Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 3,
                              listas_sul[4], regiao_maiuscula + ' - Nuclear Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Nuclear', 2,
                              listas_sul[5], regiao_maiuscula + ' - Nuclear Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 3,
                              listas_sul[6], regiao_maiuscula + ' - Eólica Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Eólica', 2,
                              listas_sul[7], regiao_maiuscula + ' - Eólica Verif', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 3,
                              listas_sul[8], regiao_maiuscula + ' - Solar Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Solar', 2,
                              listas_sul[9], regiao_maiuscula + ' - Solar Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 3,
                              carga_sul_programada, regiao_maiuscula + ' - Carga Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Carga', 2,
                              carga_sul_verificada, regiao_maiuscula + ' - Carga Verif', 11)

        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Intercâmbio', 3,
                              listas_sul[10], regiao_maiuscula + ' - Exportação Prog', 11)
        be_busca_campo_pelo_nome(worksheet_BE, regiao_maiuscula, 'Intercâmbio', 2,
                              listas_sul[11], regiao_maiuscula + ' - Exportação Verif', 11)


def ena_campos(aba, campo, lista, titulo_coluna):

    linha_e_coluna = []

    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if campo == 'Subsistema':
                if re.findall(campo, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
                else:
                    campo2 = 'Submercado'
                    if re.findall(campo2, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)
            else:
                if campo == worksheet_ENA.cell(linha, coluna).value:
                #if re.findall(campo, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)

    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]

    for scan_linha in range(linha + 1, aba.nrows):
        valor = str(worksheet_ENA.cell(scan_linha, coluna).value)
        if re.findall(r'[0-9]+\.[0-9]', valor) != []:
            lista.append(round(float(worksheet_ENA.cell(scan_linha, coluna).value), 0))
        else:
            lista.append(str(aba.cell(scan_linha, coluna).value).strip())

    # dicionario_campos[titulo_coluna] = lista
    ena_dicionario_campos[titulo_coluna] = lista


def ena_data(aba, campo, titulo_coluna):

    linha_e_coluna = []

    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if campo == 'Subsistema':
                if re.findall(campo, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
                else:
                    campo2 = 'Submercado'
                    if re.findall(campo2, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)

    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]

    for scan_linha in range(linha + 1, aba.nrows):
        valor = str(worksheet_ENA.cell(scan_linha, coluna).value)
        ena_data_arquivo.append(date)

    ena_dicionario_campos[titulo_coluna] = ena_data_arquivo


def ea_campos(aba, campo, lista, titulo_coluna, range_linha):

    linha_e_coluna = []
    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if re.findall(campo, str(aba.cell(linha, coluna).value)) != []:
                linha_e_coluna.append(linha)
                linha_e_coluna.append(coluna)
    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    for scan_linha in range(linha+range_linha, aba.nrows):
        valor = str(aba.cell(scan_linha, coluna).value)
        if re.findall(r'[0-9]+\.[0-9]', valor) != []:
            valor = round(float(aba.cell(scan_linha, coluna).value), 1)
            valor = str(valor).replace(".", ",")
            lista.append(valor)
        else:
            lista.append(str(aba.cell(scan_linha, coluna).value).strip())


    ea_dicionario_campos[titulo_coluna] = lista


def ea_data(aba, campo):

    linha_e_coluna = []
    for coluna in range(0, aba.ncols):
        for linha in range(0, aba.nrows):
            if re.findall(campo, str(aba.cell(linha, coluna).value)) != []:
                linha_e_coluna.append(linha)
                linha_e_coluna.append(coluna)
    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    for scan_linha in range(linha+0, aba.nrows):
        ea_data_arquivo.append(date)

    ea_dicionario_campos["Data"] = ea_data_arquivo


######################## ESPAÇO DESTINADO PARA TESTES DO PROGRAMA ########################

subdiretorio = r'C:\Users\GN6006\Documents\Teste_ONS\DIARIO_'

data_inicial_manipulacao = '01-01-2017'
data_inicial_manipulacao = datetime.strptime(data_inicial_manipulacao, '%d-%m-%Y')

lista_datas = ['01-01-2017.xlsx']
for dataa in range(0, len(lista_datas)):
    diretorio2 = subdiretorio+lista_datas[dataa]
    workbook2 = xlrd.open_workbook(diretorio2)
    try:
        worksheet_ENA = workbook2.sheet_by_name('19-Energia Natural Afluente')
    except:
        pass

    """linha_e_coluna = []
    field = 'ENA Bruta (MWmed) no dia'
    campo = field
    print("tamanho do valor da string: ", len(field))
    print(campo == worksheet_ENA.cell(3, 3).value)

    for coluna in range(0, worksheet_ENA.ncols):
        for linha in range(0, worksheet_ENA.nrows):
            if campo == 'Subsistema':
                if re.findall(campo, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
                else:
                    campo2 = 'Submercado'
                    if re.findall(campo2, str(worksheet_ENA.cell(linha, coluna).value)) != []:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)
            else:
                print("Entrou no ELSE")
                if campo == worksheet_ENA.cell(linha, coluna).value:
                    print("Entrou no IF")
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)

    linha = linha_e_coluna[0]
    coluna = linha_e_coluna[1]
    print(linha_e_coluna)

    for scan_linha in range(linha + 1, worksheet_ENA.nrows):
        valor = str(worksheet_ENA.cell(scan_linha, coluna).value)
        if re.findall(r'[0-9]+\.[0-9]', valor) != []:
            ena_bruta_dia.append(round(float(worksheet_ENA.cell(scan_linha, coluna).value), 0))
        else:
            ena_bruta_dia.append(str(worksheet_ENA.cell(scan_linha, coluna).value).strip())"""

######################## PROGRAMA EM SI ########################

print("Processando os Arquivos por Data...")
while data_inicial_manipulacao <= datetime.today():
    #lista_planilhas = []
    ano = str(data_inicial_manipulacao.year)
    mes = str(data_inicial_manipulacao.month).zfill(2)
    dia = str(data_inicial_manipulacao.day).zfill(2)

    date = dia+"-"+mes+"-"+ano
    diretorio = subdiretorio+date+".xlsx"

    if os.path.exists(diretorio):
        workbook = xlrd.open_workbook(diretorio) # carrega o excel

        print(date)


        try:
            worksheet_BE = workbook.sheet_by_name('01-Balanço de Energia')
            try:
                worksheet_DT = workbook.sheet_by_name('12-Motivo do Despacho Térmico')
            except:
                try:
                    worksheet_DT = workbook.sheet_by_name('11-Motivo do Despacho Térmico')
                except:
                    worksheet_DT = workbook.sheet_by_name('10-Motivo do Despacho Térmico')

            try:
                worksheet_ENA = workbook.sheet_by_name('19-Energia Natural Afluente')
            except:
                try:
                    worksheet_ENA = workbook.sheet_by_name('20-Energia Natural Afluente')
                except:
                    worksheet_ENA = workbook.sheet_by_name('21-Energia Natural Afluente')

            try:
                worksheet_EA = workbook.sheet_by_name('18-Variação Energia Armazenada')
            except:
                try:
                    worksheet_EA = workbook.sheet_by_name('19-Variação Energia Armazenada')
                except:
                    worksheet_EA = workbook.sheet_by_name('20-Variação Energia Armazenada')

        except:
            pass


        be_data_arquivo.append(date)
        be_dicionario_campos["Data"] = be_data_arquivo
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Hidro', 1, producao_hidro_programada, 'Hidro Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Hidro', 2, producao_hidro_verificada, 'Hidro Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Itaipu', 1, producao_itaipu_programada, 'Itaipu Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Itaipu', 2, producao_itaipu_verificada, 'Itaipu Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Nuclear', 1, producao_nuclear_programada, 'Nuclear Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Nuclear', 2, producao_nuclear_verificada, 'Nuclear Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Termo', 1, producao_termo_programada, 'Termo Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Termo', 2, producao_termo_verificada, 'Termo Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Eólica', 1, producao_eolica_programada, 'Eólica Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Eólica', 2, producao_eolica_verificada, 'Eólica Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Solar', 1, producao_solar_programada, 'Solar Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Solar', 2, producao_solar_verificada, 'Solar Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Intercâmbio', 1, exportacao_programada, 'Exportação Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Intercâmbio', 2, exportacao_verificada, 'Exportação Verificada', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Carga', 1, carga_programada, 'Carga Programado', 11)
        be_busca_campo_pelo_nome(worksheet_BE, 'Interligado', 'Carga', 2, carga_verificada, 'Carga Verificada', 11)
        be_busca_regioes('Norte', 'norte')
        be_busca_regioes('Nordeste', 'nordeste')
        be_busca_regioes('Sudeste / Centro-Oeste', 'sudeste')
        be_busca_regioes('Sul', 'sul')


        dt_data(worksheet_DT, 'Usina')
        dt_campos(worksheet_DT, 'Usina', dt_usina, 'Usina')
        dt_campos(worksheet_DT, 'Código', dt_cod_ONS, 'Código ONS')
        dt_campos(worksheet_DT, 'Potência', dt_potencia, 'Potência Instalada')
        dt_campos(worksheet_DT, 'Ordem', dt_ordem_de_merito, 'Ordem de Mérito')
        dt_campos(worksheet_DT, 'Inflex.', dt_inflex, 'Inflexibilidade')
        dt_campos(worksheet_DT, 'Restrição', dt_restricao, 'Restrição Elétrica')
        dt_campos(worksheet_DT, 'Geração', dt_geracao_fora_da_ordem, 'Ger. Fora da Ordem')
        dt_campos(worksheet_DT, 'Energia', dt_energia_de_reposicao, 'Energia de Reposição')
        dt_campos(worksheet_DT, 'Garantia', dt_garantia_energetica, 'Garantia Energética')
        dt_campos(worksheet_DT, 'Export.', dt_exportacao, 'Exportação')
        dt_campos(worksheet_DT, 'Verificado', dt_verificado, 'Verificado')


        ena_data(worksheet_ENA, 'Subsistema', 'Data')
        ena_campos(worksheet_ENA, 'Subsistema', ena_subsistema, 'Subsistema')
        ena_campos(worksheet_ENA, '% MLT no dia', ena_percent_mlt_dia, '% MLT no dia')
        ena_campos(worksheet_ENA, '% MLT acumulado no mês até o dia', ena_percent_mlt_acum, '% MLT acumulado no mês até o dia')
        ena_campos(worksheet_ENA, 'ENA Bruta (MWmed) no dia', ena_bruta_dia, 'ENA Bruta (MWmed) no dia')
        ena_campos(worksheet_ENA, 'ENA Bruta (MWmed) acumulada até o dia', ena_bruta_acum, 'ENA Bruta (MWmed) acumulada até o dia')


        ea_data(worksheet_EA, 'Capacidade Máxima')
        ea_campos(worksheet_EA, 'Capacidade Máxima',ea, 'Energia Armazenada', 0)
        ea_campos(worksheet_EA, 'Sul', ea_sul, 'Sul', 1)
        ea_campos(worksheet_EA, 'SE/CO', ea_sudeste, 'SE/CO', 1)
        ea_campos(worksheet_EA, 'Norte', ea_norte, 'Norte', 1)
        ea_campos(worksheet_EA, 'NE', ea_nordeste, 'NE', 1)


    data_inicial_manipulacao = data_inicial_manipulacao + timedelta(days=1)

######################## SALVANDO O ARQUIVO EM EXCEL ########################

print("Salvando os Arquivos em Excel...")
lista_impressao.append(be_dicionario_campos.copy())
lista_impressao.append(dt_dicionario_campos.copy())
lista_impressao.append(ena_dicionario_campos.copy())
lista_impressao.append(ea_dicionario_campos.copy())


arquivo_saida = r'\\br-gve-fs\gvdocs$\01. Gestao Clientes\01.1 ESTUDOS NOVOS CLIENTES\MODELO\Novos modelos\PBI\Banco Boletim da Operação\DIARIO - Final_.xlsx'
writer = pd.ExcelWriter(arquivo_saida, engine='xlsxwriter')

abas = ['01-Balanço de Energia', '12-Motivo do Despacho Térmico', '19-Energia Natural Afluente',
        '20-Variação Energia Armazenada']

lista_df = []

for dicionario in range(0, len(lista_impressao)):
    workbook = writer.book
    df = pd.DataFrame.from_dict(lista_impressao[dicionario], orient='index')
    df = df.transpose()
    lista_df.append(df)


for dataframe in range(0, len(lista_df)):
    lista_df[dataframe].to_excel(writer, sheet_name=str(abas[dataframe]),
                                 startrow=0, startcol=0, index=False)
writer.save()


tempo_fim = time.time()
print("A execução do programa levou {} minutos".format(round((tempo_fim - tempo_ini)/60),2))
