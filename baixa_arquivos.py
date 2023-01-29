import requests
import urllib
from datetime import datetime, timedelta
import os.path


class BaixaArquivos:
    """
    Antes de passar para a etapa de transformação, é necessário ter todo o histórico do ONS. 
    Para tal, utilizou-se a biblioteca requests para baixar cada arquivo pelo sua URL.
    ...
    Métodos
    -------
    O único método desta classe é o __baixaArquivos, que não recebe nenhum parâmetro,
    mas reflete a ideia principal de através de uma url fazer um request para baixar o arquivo.
    """
    
    def __init__(self):
        self.__data_inicial__ = '01-01-2017'
        self.__diretorio__ = 'files/'
        self.__url_pt_1__ = r'http://sdro.ons.org.br/SDRO/DIARIO/'

        self.__baixaArquivos()

    def __baixaArquivos(self):

        url_pt_1 = self.__url_pt_1__
        diretorio = self.__diretorio__
        data_inicial = self.__data_inicial__

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

            if os.path.exists(diretorio + str(url[-22:])):
                pass
            else:
                print(data, "ARQUIVO NÃO EXISTE")
                if str(requests.get(url)) == '<Response [200]>':
                    urllib.request.urlretrieve(url, diretorio + str(url[-22:]))
                    print("Arquivo Baixado")

            data_inicial = data_inicial + timedelta(days=1)

if __name__ == "__main__":
    BaixaArquivos()
