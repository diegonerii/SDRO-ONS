# SDRO-ONS

Este projeto tem como objetivo principal a disponibilização de dados do Boletim Diário da Operação do ONS em Excel e de forma colunar.

Para isso, o projeto está dividido em 3 arquivos principais:

    1 - baixa_arquivos.py: através de scraping, vai baixar todos os arquivos históricos do Boletim Diário da Operação e depois armazenar na pasta "files". A data inicial é 01/01/2017.
    2 - Funcoes.py: este vai ser o módulo que vai abrigar todas as classes de transformação do dado, uma para cada aba do arquivo diário. As abas são:
        2.1 - Balanço Energético
        2.2 - Despacho Térmico
        2.3 - Energia Natural Afluente
        2.4 - Energia Armazenada
    3 - app.py: será o arquivo que vai instanciar as classes definidas em (2) e salvar o excel final.

## Observação1: versão do xlrd precisa ser menor que a 2.0.0, podendo ser encontrada em: https://pypi.org/project/xlrd/#history
## Observação2: a ordem de execução sempre tem que seguir (1), (2) e (3). Caso contrário, não existirá arquivos históricos (pasta folder/) para se trabalhar.
