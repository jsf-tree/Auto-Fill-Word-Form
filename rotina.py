# -*- coding: utf-8 -*-
# O encoding padrão do Python 3+ já é UTF-8, mas essa rotina foi desenvolvido em Python 2
# e talvez algum dia precise ser rodado em Python 2 de novo. Quem sabe?
"""
OBJETIVO:  Automatizar o preenchimento de FT-23 (Relatório de amostragem)
  Autoria: Juliano Santos Finck
  E-Mail: juliano.finck@gmail.com
  Data: 11/03/2021

RESUMO:
  1. IMPORTAR BIBLIOTECAS
  2. FUNÇÕES
  3. LER AS INFOS
     3.1 AMOSTRAGEM.xlsx
     3.2 CLIENTE_e_PROJETO.txt
     3.3 LER FT14
  4. PROCESSAMENTO
     4.1 PREPARANDO DIRETÓRIO
     4.2 DESCOMPACTANDO (UNZIPPING) O TEMPLATE DO 'FT-23' PARA PASTA tmp
     4.3 ABRIR E ADICIONAR AS INFORMAÇÕES DE CLIENTE e PROJETO
     4.4 ADICIONAR INFORMAÇÕES DE CADA AMOSTRAGEM EM LOOP
        4.4.1 SUBSTITUIR
        4.4.2. ZIPPAR, SALVAR, RENOMEAR PARA ".docx"
     4.5 LIMPAR DIRETÓRIO
  5. MENSAGEM FINAL

# =================================================================================== #
# OBSERVAÇÕES:
# Expandir o relatório de discrepâncias.
"""
# =================================================================================== #
# 1. IMPORTAR BIBLIOTECAS
# Função para tentar impotar; e, se falhar (não houver o pacote), instalar.

def install_if_nonexistent(module_name, install_name = 'same_name'):
    if install_name == 'same_name':
        install_name = module_name
    from importlib import util
    if util.find_spec(module_name) is not None:
        pass
    else:
        import os
        print('Instalando novo modulo... Aguarde\n\n')
        os.system('python -m pip install ' + install_name)

install_if_nonexistent('pandas')
install_if_nonexistent('xlrd')
install_if_nonexistent('openpyxl')
import pandas as pd
import numpy as np
import os
import shutil
import codecs
import math
from zipfile import ZipFile


# =================================================================================== #
# 2. FUNÇÕES
def mensagem_final(msg, assinatura, n):
    print('\033[92m| ' + n * ' ' + ' |' + '\033[0m')
    for i in msg.split('\n'):
        print('\033[92m| ' + '\033[0m' + i + (n - len(i)) * ' ' + '\033[92m |')
    print('\033[92m| ' + n * ' ' + ' |' + '\033[0m')
    for i in assinatura.split('\n'):
        line = '{:>' + str(n) + '}'
        line = line.format(i)
        print('\033[92m| ' + line + ' |')
    print('\033[92m| ' + n * ' ' + ' |' + '\033[0m')


def divisao(n):
    print('\033[92m' + '# ' + n * '=' + ' #')


# =================================================================================== #
# 3. LER AS INFOS
## 3.1 AMOSTRAGEM.xlsx
# Lê o excel
AMOSTRAGENS = pd.read_excel('INPUT\\1. AMOSTRAGENS.xlsx', index_col=0)
os.system("cls"); divisao(84)
# Armazena os ID da Amostra
i_amostragem = list(AMOSTRAGENS['ID da Amostra'])
# Retirar o '0' porque é a linha escondida de parâmetros
i_amostragem = [i for i, x in enumerate(i_amostragem) if 'nan' != str(x)][1:]
# i_amostragem = linhas com amostras em AMOSTRAGEM.xlsx

amostragens = []
for i in i_amostragem:
    linha = list(AMOSTRAGENS.iloc[i])
    amostragens.append(linha[1:])
parametros = list(AMOSTRAGENS.iloc[1])[1:]
# amostragens = lista cujos elementos são as linha de cada amostragem em AMOSTRAGENS.xlsx
# parametros = lista cujos elementos são o nome das colunas de AMOSTRAGENS.xlsx

## 3.2 CLIENTE_e_PROJETO.txt
# pessimamente escrito!
# ao invés de montar um dicionário, esta sendo lida cada linha que tenha ":" e salvando toda informação
# no resto dessa linha como um elemento da lista "cliente_projeto"
INFOS = r'INPUT\2. CLIENTE_e_PROJETO.txt'
print('\033[92m' + '1. CLIENTE e PROJETO.txt:' + '\033[0m')
dados = codecs.open(INFOS, mode='r', encoding='utf-8')
cliente_projeto = []
for row in dados:
    if ':' in row:
        print(row, end='')
        cliente_projeto.append(row[row.index(':') + 1:row.index(';')])
print('\n')

## 3.3 LER FT14
# Abre o FT14 para buscar quais os ID_amostra que tem e anotar os volumes esperados
LISTDIR = os.listdir('INPUT\\')
print(LISTDIR)
i = len(LISTDIR) - 1
# Busca o arquivo, cujo nome contenha "FT-14"
# todo Melhor que haja uma pasta "Input" onde estejam AMOSTRAGENS e FT-14
# todo também outra pasta "Output" onde estejam os relatórios preenchidos e o relatório de discrepância (se houver)
while i >= 0:
    if 'FT-14' in LISTDIR[i]:
        filename = LISTDIR[i]
        i = - 1
    i -= 1
FT14 = pd.read_excel('INPUT\\'+filename, index_col=0)

# Linhas de Parâmetros e ID_Amostras
i0_par = FT14.index.get_loc('Parâmetro') + 1
i1_par = FT14.index.get_loc('INSTRUÇÕES DE PREENCHIMENTO DA CADEIA DE CUSTÓDIA GÖRTLER')
i0_a = [i for i, x in enumerate(FT14.index.get_loc('ID da Amostra')) if x][0] + 2
i1_a = FT14.index.get_loc('INSTRUÇÕES DE PREENCHIMENTO DA CADEIA DE CUSTÓDIA (prencher com o nome do laboratório)')
# i0_par = 93; i1_par = 97; i0_a = 100; i1_a 109
# No Excel começa a contar de 1, no Python de 0. Ele identificou linha 93-97 tendo parâmetros (no Excel, 94-98)
# Identificou do 100 ao 109 tendo amostras. No Excel, 101-110.

# Obtem os "ID_Amostras" do FT-14 e Remover BC (List comprehensions)
ID_Amostras = list(FT14.index)[i0_a:i1_a]
ID_Amostras = [x for x in ID_Amostras if "BC-" not in x]

# Identificar parâmetros a marcar e volumes esperados
pars = []; vol = []
for amostra in ID_Amostras:
    print(amostra) #deletar
    # Localiza no DataFrame "FT14" onde está a Amostra e copia os "Parâmetros a marcar"
    pars.append(list(FT14.loc[amostra])[6])

    # Se houver VOC e TPH, remover TPH porque é o mesmo frasco
    # JUDITE DISSE QUE É SVOC AO INVÉS DE VOC
    if 'VOC,' in pars[-1] and 'TPH' in pars[-1]:
        PARS = pars[-1].replace('TPH', '').replace(' ', '').split(',')
    else:
        PARS = pars[-1].replace(' ', '').split(',')

    # Conta o volume esperado
    VOL = 0
    # PARS = [ ['VOC', 'SVOC', 'Metais', 'FP'],    ['VOC', 'SVOC', 'Metais', 'FP', 'PCB'],   ...]
    for PAR in PARS:
        for _ in range(i0_par, i1_par):
            if PAR != 'VOC' and PAR in list(FT14.index)[_]:
                temp = list(FT14.iloc[_])[3].replace('mL', '').replace('x', '*')
                print(temp) #deletar
                exec('VOL +=' + temp)
            elif PAR == 'VOC' and PAR in list(FT14.index)[_] and 'SVOC' not in list(FT14.index)[_]:
                temp = list(FT14.iloc[_])[3].replace('mL', '').replace('x', '*')
                print(temp) #deletar
                exec('VOL +=' + temp)
    print('\n')
    vol.append(VOL)

# =================================================================================== #
# 4. PROCESSAMENTO
## 4.1 PREPARANDO DIRETÓRIO
# Reinicializa os diretórios "\OUTPUT" e "\tmp" <- tmp usado para descompactar Template e preencher os FT pelo código-fonte
print('RELATÓRIOS FT-23 PREENCHIDOS:\033[0m')
try:
    shutil.rmtree('tmp')
except:
    os.mkdir('tmp')
try:
    shutil.rmtree('OUTPUT')
except:
    os.mkdir('OUTPUT')

## 4.2 DESCOMPACTANDO (UNZIPPING) O TEMPLATE DO 'FT-23' PARA PASTA tmp
with ZipFile('PADRAO\FT-23_rev 10 TEMPLATE.docx', 'r') as zip_ref:
    zip_ref.extractall('tmp')

## 4.3 ABRIR E ADICIONAR AS INFORMAÇÕES DE CLIENTE e PROJETO
doc_filename = 'tmp\word\document.xml'
fileR = codecs.open(doc_filename, mode='r+', encoding='utf-8')
doc_xml = []; count = 0
# abre document.xml e copia tudo o que está escrito, substituindo os VARX (que tem que estar na mesma ordem que em "cliente_projeto" com dicionário seria mais seguro
for row in fileR:
    if count != 0:
        for _ in range(0, len(cliente_projeto)):
            row = row.replace('VAR' + str(_ + 2), cliente_projeto[_])
    count += 1
    doc_xml.append(row)
fileR.close()

# abre header e copia tudo o que está escrito
header_filename = 'tmp\word\header1.xml'
fileR = codecs.open(header_filename, mode='r+', encoding='utf-8')
header_xml = []; count = 0
for row in fileR:
    header_xml.append(row)
fileR.close()

## 4.4 ADICIONAR INFORMAÇÕES DE CADA AMOSTRAGEM EM LOOP
### 4.4.1 SUBSTITUIR
j = 0  # (amarrar todos Drop-Down list a um índice único; facilita alterações)
for amostragem in amostragens:
    fileW = codecs.open(doc_filename, mode='w', encoding='utf-8')
    for I in doc_xml:
        i = len(amostragem) - 1
        while i >= 0:
            var = amostragem[i]
            if i == j + 1:
                I = I.replace('>Vagner Lopes Leivas', '>' + str(var))
            elif i == j + 6:
                I = I.replace('>Instável', '>' + str(var))
            elif i == j + 8:
                I = I.replace('>Sem observações', '>' + str(var))
            elif i == j + 7:
                I = I.replace('>Sem características relevantes', '>' + str(var))
            elif i == j + 9:
                I = I.replace('>Baixa vazão - ABNT NBR 15.847:2010', '>' + str(var))
            elif i == j + 10:
                I = I.replace('>Multiparâmetro: Mu_05 | YSI | Professional Plus', '>' + str(var))
            elif i == j + 11:
                I = I.replace('>Medidor de interface: I-16 | Solinst | 122', '>' + str(var))
            elif i == j + 12:
                I = I.replace('>Termômetro: T-02 | Simpla | DT160', '>' + str(var))
            elif i == j + 13:
                I = I.replace('>Bailer | Sauber System', '>' + str(var))
            else:
                I = I.replace('var' + str(i), str(var))
            i -= 1
        fileW.write(I)
    fileW.close()

    fileW = codecs.open(header_filename, mode='w', encoding='utf-8')
    for I in header_xml:
        I = I.replace('var0', amostragem[0])
        fileW.write(I)
    fileW.close()

    ### 4.4.2. ZIPPAR, SALVAR, RENOMEAR PARA ".docx"
    FILE = 'OUTPUT/FT-23_rev 10 - Relatório de Amostragem ' + amostragem[0].replace('/', '_')
    FILEdocx = FILE[:7] + FILE[22:] + '.docx'
    shutil.make_archive(FILE, 'zip', 'tmp')
    os.rename(FILE + '.zip', FILEdocx)

## 4.5 RELATÓRIO DE DISCREPÂNCIA
fileW = codecs.open('OUTPUT/Relatório_discrepâncias.csv', mode='w', encoding='utf-8')
fileW.write('FT-25;ID da Amostra;Volume amostragem (FT-03);Volume plano de amostragem (FT-14)\n')
print('\033[92m' + '2. FT-25, ID da Amostra, Volume amostragem (FT-03), e Volume plano de amostragem (FT-14)' + '\033[0m')
i = 0
for amostragem in amostragens:
    print(amostragem[0] + '  ' + amostragem[2] + '  ' + amostragem[26] + '  ' + str(vol[i] / 1000).replace('.', ','))
    fileW.write(amostragem[0] + ';' +
                amostragem[2] + ';' +
                amostragem[26] + ';' +
                str(vol[i] / 1000).replace('.', ',') + '\n')
    i += 1
    j = len(amostragem) - 1
    while j >= 0:
        if str(amostragem[j]) == 'nan':
            fileW.write(amostragem[0] + ';' +
                        amostragem[2] + ';' +
                        parametros[j] + ';' +
                        u'Faltou preenchimento' + '\n')  # alterar para o nome da variável
        j -= 1
fileW.close()
print('')
divisao(84)
# =================================================================================== #
## 4.5 LIMPAR DIRETÓRIO
shutil.rmtree('tmp')

# =================================================================================== #
# 5. MENSAGEM FINAL
print('')
divisao(84)
msg = 'Pronto! ' + str(len(amostragens)) + ' FT-23s preenchidas na pasta Relatórios!\n' \
                                           'Peça a algum colega que revise as FT-23s com FT-03, FT-14 e FT-25 em mãos!'

assinatura = 'Auto-preencher FT-23 v.0\n' \
             'Sapotec Sul 2021\n' \
             '----------------\n' \
             'Juliano Santos Finck\n' \
             '12/03/2021\n'
mensagem_final(msg, assinatura, 84)
divisao(84)
