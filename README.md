![Operador Nacional do Sistema](https://user-images.githubusercontent.com/35044363/221372716-a5065a5b-b874-4362-a1bd-e58b162c5c4b.png)
# SDRO-ONS

## Introdução
Este projeto tem como objetivo principal a disponibilização de dados do Boletim Diário da Operação do ONS em Excel e de forma colunar. Atualmente, de todas as planilhas nas pastas excel, somente estão catalogadas: **Balanço Energético**, **Despacho Térmico**, **Energia Natural Afluente** e **Energia Armazenada**.

## Arquivos
Para isso, o projeto está dividido em 3 arquivos principais:

1. baixa_arquivos.py: através de scraping, vai baixar todos os arquivos históricos do Boletim Diário da Operação e depois armazenar na pasta "files". A data inicial é 01/01/2017.
2. funcoes.py: módulo que vai abriga todas as classes de transformação dos dados, uma para cada planilha do arquivo diário:
    2.1. Balanço Energético
    2.2. Despacho Térmico
    2.3. Energia Natural Afluente
    2.4. Energia Armazenada
3. app.py: será o arquivo que vai instanciar as classes definidas em (2) e fazer as etapas para salvar o excel final.

## Como Utilizar

Para utilizar este projeto, é necessário:

1. Instalar uma versão do pacote xlrd menor que a 2.0.0, podendo ser encontrada em: https://pypi.org/project/xlrd/#history
2. Verificar se existe espaço de armazenamento local disponível. Isso porque o baixa_arquivos.py vai fazer o download de milhares de arquivos e colocá-los na pasta local chamada "files".

## Contatos

<div>
<a href = "mailto:diego.holanda.neri@gmail.com"><img src="https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white" target="_blank"></a>
<a href="https://www.linkedin.com/in/diego-neri/" target="_blank"><img src="https://img.shields.io/badge/-LinkedIn-%230077B5?style=for-the-badge&logo=linkedin&logoColor=white" target="_blank"></a>   
</div>
