import xlrd
import re
import os.path
from datetime import datetime

"""
importante: versão do xlrd precisa ser menor que a 2.0.0, podendo ser encontrada em: https://pypi.org/project/xlrd/#history
"""

# ######################## MANIPULAÇÃO DE ARQUIVO ########################

dicionario_campos = {}
be_dicionario_campos = {}
dt_dicionario_campos = {}
ena_dicionario_campos = {}
ea_dicionario_campos = {}


class Inicializacao:
    """
    Inicalização vai ser a superclasse. Ela vai oferecer os atributos que serão utilizados pelas classes filhas a seguir
    """
    
    def __init__(self, data_do_arquivo):
        
        self.subdiretorio = r'files\DIARIO_'
        self.data_inicial_manipulacao = data_do_arquivo
        self.data_inicial_manipulacao = datetime.strptime(self.data_inicial_manipulacao, '%d-%m-%Y')

        self.ano = str(self.data_inicial_manipulacao.year)
        self.mes = str(self.data_inicial_manipulacao.month).zfill(2)
        self.dia = str(self.data_inicial_manipulacao.day).zfill(2)

        self.date = self.dia+"-"+self.mes+"-"+self.ano
        self.diretorio = self.subdiretorio+self.date+".xlsx"

        if os.path.exists(self.diretorio):
            self.workbook = xlrd.open_workbook(self.diretorio) # carrega o excel

        print(data_do_arquivo)


class DespachoTermico(Inicializacao):
    def __init__(self, data_do_arquivo):
        
        """
        Herda os atributos da superclasse.
        """
        super().__init__(data_do_arquivo)

        """
        Dentre todas as planilhas da pasta, ele vai procurar através de um regex somente a que tem 'Motivo do Despacho Térmico'.
        """
        for sheet_name in self.workbook.sheet_names():
            if re.findall('Motivo do Despacho Térmico', sheet_name) != []:
                self.worksheet_DT = self.workbook.sheet_by_name(sheet_name) 

        fontes = {
            'Usina':'Usina', 
            'Código':'Código ONS', 
            'Potência':'Potência Instalada', 
            'Ordem':'Ordem de Mérito', 
            'Inflex.':'Inflexibilidade', 
            'Restrição':'Restrição Elétrica',
            'Geração':'Ger. Fora da Ordem', 
            'Energia':'Energia de Reposição',
            'Garantia':'Garantia Energética', 
            'Export.':'Exportação', 
            'Verificado':'Verificado'
        }

        """
        A linha de código abaixo somente capturará a data do arquivo.
        """
        self.dt_data('Usina')

        """
        Através do título de cada key do dicionário acima, vai fazer uma varredura nos valores.
        """
        for fonte in range(0, len(list(fontes.keys()))): 
            self.dt_campos(list(fontes.keys())[fonte], list(fontes.values())[fonte])
                
    def dt_data(self, campo):

        linha_e_coluna = []
        lista = []

        """Laço para achar onde está a célula de referência"""
        for coluna in range(0, self.worksheet_DT.ncols):
            for linha in range(0, self.worksheet_DT.nrows):
                if re.findall(campo, str(self.worksheet_DT.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]
        
        """Vai apendar a data do arquivo em uma lista para cada valor achado"""
        for scan_linha in range(linha+1, self.worksheet_DT.nrows):
            lista.append(self.date)

        dt_dicionario_campos.setdefault("Data", []).extend(lista)

    
    def dt_campos(self, campo, titulo_coluna):

        linha_e_coluna = []
        lista = []

        """Laço para achar onde está a célula de referência"""
        for coluna in range(0, self.worksheet_DT.ncols):
            for linha in range(0, self.worksheet_DT.nrows):
                if re.findall(campo, str(self.worksheet_DT.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]
        
        """
        Vai fazer uma varredura e buscar tudo o que for número separado por ponto.
        Depois, coloca esse valor numa lista.        
        """
        for scan_linha in range(linha+1, self.worksheet_DT.nrows):
            valor = str(self.worksheet_DT.cell(scan_linha, coluna).value).strip()
            if re.findall(r'[0-9]+\.[0-9]', valor) != []:
                valor = round(float(self.worksheet_DT.cell(scan_linha, coluna).value), 1)
                valor = str(valor).replace(".", ",")
                lista.append(valor)
            else:
                lista.append(str(self.worksheet_DT.cell(scan_linha, coluna).value).strip())

        dt_dicionario_campos.setdefault(titulo_coluna, []).extend(lista) 


class BalancoEnergetico(Inicializacao):
    def __init__(self, data_do_arquivo):
        
        """
        Herda os atributos da superclasse.
        """
        super().__init__(data_do_arquivo)

        """
        Dentre todas as planilhas da pasta, ele vai procurar através de um regex somente a que tem 'Balanço de Energia'.
        """
        for sheet_name in self.workbook.sheet_names():
            if (re.findall('Balanço de Energia', sheet_name) != []) and (re.findall('Balanço de Energia Acumulado', sheet_name) == []):
                self.worksheet_BE = self.workbook.sheet_by_name(sheet_name) 
        
        """
        criando uma coluna de data
        """
        be_data_arquivo = []
        be_data_arquivo.append(self.date)
        be_dicionario_campos.setdefault("Data", []).extend(be_data_arquivo)
        
        submercados = ['Interligado', 'Norte', 'Nordeste', 'Sudeste / Centro-Oeste', 'Sul']
        for submercado in submercados:
            self.be_busca_regioes(submercado)


    def __be_busca_valores(self, tipo_de_energia, range_coluna, nome_coluna, submercado='Interligado', range_linha=int(11)):
        
        """
        Este método vai atuar dentro da planilha Balanço de Energia. Ele tem por objetivo 
        sumarizar os dados de produção de energia no SIN e em cada submercado.
        """
        vetor_regex = []
        linha_e_coluna = []
        self.lista = []

        """Laço para achar onde está a célula do Interligado, que servirá de referência"""
        for coluna in range(0, self.worksheet_BE.ncols):
            for linha in range(0, self.worksheet_BE.nrows):
                if re.findall(submercado, str(self.worksheet_BE.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]
        
        """Depois de achar a célula, o programa faz uma varredura de linha a linha e apendando numa lista"""
        for scan_linha in range(linha, linha+range_linha):
            if re.findall(tipo_de_energia, str(self.worksheet_BE.cell(scan_linha, coluna).value)) != []:
                valor = self.worksheet_BE.cell(scan_linha, coluna+range_coluna).value
                # print(valor)
                if valor != '':
                    valor = round(self.worksheet_BE.cell(scan_linha, coluna+range_coluna).value, 0)
                    vetor_regex.append(valor)
                else:
                    vetor_regex.append(0)
            else:
                vetor_regex.append(0)
        """
        o vetor_regex é uma lista com vários valores (match do regex). Ele só apenda o maior valor.
        """
        if min(vetor_regex) >= 0:
            self.lista.append(max(vetor_regex))
        else:
            self.lista.append(min(vetor_regex))

        be_dicionario_campos.setdefault(nome_coluna, []).extend(self.lista) 


    def be_busca_regioes(self, submercado):
        """
        Depois de estabelecido o método de buscar os valores, utiliza-se be_busca_regioes para buscar valores através da palavras de referência. Os condicionais abaixo foram necessário, pois cada submercado tem um range diferente a ser buscado.
        """
        fontes = ['Hidro', 'Itaipu', 'Nuclear', 'Termo', 'Eólica', 'Solar', 'Intercâmbio', 'Exportação', 'Importação', 'Carga']

        for fonte in fontes:
            if submercado == 'Interligado':
                self.__be_busca_valores(fonte, 1, f'SIN {fonte} Programado', submercado='Interligado')
                self.__be_busca_valores(fonte, 2, f'SIN {fonte} verificada', submercado='Interligado')
            
            elif (submercado == 'Norte') | (submercado == 'Nordeste'):
                self.__be_busca_valores(fonte, 2, f'{submercado} {fonte} Programado', submercado=submercado, range_linha=7)
                self.__be_busca_valores(fonte, 3, f'{submercado} {fonte} verificada', submercado=submercado, range_linha=7)

            else:
                self.__be_busca_valores(fonte, 2, f'{submercado} {fonte} Programado', submercado=submercado, range_linha=8)
                self.__be_busca_valores(fonte, 3, f'{submercado} {fonte} verificada', submercado=submercado, range_linha=8)         


class ENA(Inicializacao):
    def __init__(self, data_do_arquivo):

        """
        Herda os atributos da superclasse.
        """
        super().__init__(data_do_arquivo)

        """
        Dentre todas as planilhas da pasta, ele vai procurar através de um regex somente a que tem 'Energia Natural Afluente'.
        """
        for sheet_name in self.workbook.sheet_names():
            if re.findall('Energia Natural Afluente', sheet_name) != []:
                self.worksheet_ENA = self.workbook.sheet_by_name(sheet_name) 

        fontes = {
            'Subsistema':'Subsistema', 
            '% MLT no dia':'% MLT no dia', 
            '% MLT acumulado no mês até o dia':'% MLT acumulado no mês até o dia', 
            'ENA Bruta (MWmed) no dia':'ENA Bruta (MWmed) no dia', 
            'ENA Bruta (MWmed) acumulada até o dia':'ENA Bruta (MWmed) acumulada até o dia'
        }

        self.ena_data('Subsistema', 'Data')

        for fonte in range(0, len(list(fontes.keys()))): 
            self.ena_campos(list(fontes.keys())[fonte], list(fontes.values())[fonte])
               
    def ena_data(self, campo, titulo_coluna):
        
        """
        Este primeiro método tem como objetivo pegar a data do arquivo e apendar
        numa lista de tamanho equivalente ao número de máximo de registros de uma
        coluna. Primeiro ele acha os valores de referência para linha e coluna e depois faz uma varredura.
        """
        linha_e_coluna = []
        lista = []

        for coluna in range(0, self.worksheet_ENA.ncols):
            for linha in range(0, self.worksheet_ENA.nrows):
                if campo == 'Subsistema':
                    if re.findall(campo, str(self.worksheet_ENA.cell(linha, coluna).value)) != []:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)
                    else:
                        campo2 = 'Submercado'
                        if re.findall(campo2, str(self.worksheet_ENA.cell(linha, coluna).value)) != []:
                            linha_e_coluna.append(linha)
                            linha_e_coluna.append(coluna)

        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]

        for scan_linha in range(linha + 1, self.worksheet_ENA.nrows):
            valor = str(self.worksheet_ENA.cell(scan_linha, coluna).value)
            lista.append(self.date)

        ena_dicionario_campos.setdefault(titulo_coluna, []).extend(lista) 

    def ena_campos(self, campo, titulo_coluna):
        
        """
        Seguindo a mesma lógica do método anterior, este método vai buscar
        todos os valores das colunas a partir de uma palavra de referência e dos valores achados para linha e coluna no método anterior.
        """
        linha_e_coluna = []

        for coluna in range(0, self.worksheet_ENA.ncols):
            for linha in range(0, self.worksheet_ENA.nrows):
                if campo == 'Subsistema':
                    if re.findall(campo, str(self.worksheet_ENA.cell(linha, coluna).value)) != []:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)
                    else:
                        campo2 = 'Submercado'
                        if re.findall(campo2, str(self.worksheet_ENA.cell(linha, coluna).value)) != []:
                            linha_e_coluna.append(linha)
                            linha_e_coluna.append(coluna)
                else:
                    if campo == self.worksheet_ENA.cell(linha, coluna).value:
                        linha_e_coluna.append(linha)
                        linha_e_coluna.append(coluna)

        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]

        lista = []

        for scan_linha in range(linha + 1, self.worksheet_ENA.nrows):
            valor = str(self.worksheet_ENA.cell(scan_linha, coluna).value)
            if re.findall(r'[0-9]+\.[0-9]', valor) != []:
                lista.append(round(float(self.worksheet_ENA.cell(scan_linha, coluna).value), 0))
            else:
                lista.append(str(self.worksheet_ENA.cell(scan_linha, coluna).value).strip())

        ena_dicionario_campos.setdefault(titulo_coluna, []).extend(lista) 


class EnergiaArmazenada(Inicializacao):
    def __init__(self, data_do_arquivo):
        
        """
        Herda os atributos da superclasse.
        """
        super().__init__(data_do_arquivo)

        """
        Dentre todas as planilhas da pasta, ele vai procurar através de um regex somente a que tem 'Variação Energia Armazenada'.
        """
        for sheet_name in self.workbook.sheet_names():
            if re.findall('Variação Energia Armazenada', sheet_name) != []:
                self.worksheet_EA = self.workbook.sheet_by_name(sheet_name)

        self.ea_data('Capacidade Máxima')
        self.ea_campos('Capacidade Máxima', 'Energia Armazenada', 0)

        submercados = ['Sul', 'SE/CO', 'Norte', 'NE']
        for submercado in range(0, len(submercados)):
            self.ea_campos(submercados[submercado], submercados[submercado], 1)

    def ea_data(self, campo):
        
        """
        Este primeiro método tem como objetivo pegar a data do arquivo e apendar
        numa lista de tamanho equivalente ao número de máximo de registros de uma
        coluna. Primeiro ele acha os valores de referência para linha e coluna e depois faz uma varredura.
        """
        linha_e_coluna = []
        lista = []

        for coluna in range(0, self.worksheet_EA.ncols):
            for linha in range(0, self.worksheet_EA.nrows):
                if re.findall(campo, str(self.worksheet_EA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]
        for scan_linha in range(linha+0, self.worksheet_EA.nrows):
            lista.append(self.date)

        ea_dicionario_campos.setdefault("Data", []).extend(lista) 

    def ea_campos(self, campo, titulo_coluna, range_linha):
        
        """
        Seguindo a mesma lógica do método anterior, este método vai buscar
        todos os valores das colunas. Primeiro ele acha o valor máximo de cada linha e coluna a partir de uma referência e depois faz uma varredura.
        """
        linha_e_coluna = []
        lista = []

        for coluna in range(0, self.worksheet_EA.ncols):
            for linha in range(0, self.worksheet_EA.nrows):
                if re.findall(campo, str(self.worksheet_EA.cell(linha, coluna).value)) != []:
                    linha_e_coluna.append(linha)
                    linha_e_coluna.append(coluna)
        linha = linha_e_coluna[0]
        coluna = linha_e_coluna[1]

        for scan_linha in range(linha+range_linha, self.worksheet_EA.nrows):
            valor = str(self.worksheet_EA.cell(scan_linha, coluna).value)
            if re.findall(r'[0-9]+\.[0-9]', valor) != []:
                valor = round(float(self.worksheet_EA.cell(scan_linha, coluna).value), 1)
                valor = str(valor).replace(".", ",")
                lista.append(valor)
            else:
                lista.append(str(self.worksheet_EA.cell(scan_linha, coluna).value).strip())

        ea_dicionario_campos.setdefault(titulo_coluna, []).extend(lista) 