# -*- coding: utf-8 -*-
"""
Created on Wed Dec 21 09:04:59 2022

@author: Projeto6
Autor: Giovanni Francisco Weiss
Data: 20/12/2022

Ferramente para renomear desenhos automaticamente.

Feito para obter informações de listas de materiais conforme o padrão
da Equimatec em 2022.

Para funcionar, é necessário que a lista contenha uma coluna contendo
os números dos desenhos na sexta posição e outra contendo os códigos
SAP na quarta posição.

Não funciona para muitos arquivos nomeados também, somente numéricos.
"""

import os  
import glob
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
# import pathlib

# Criar endereço da BOM e pasta desejadas.
Tk().withdraw()
filename = askopenfilename(title='Select BOM') 
path = askdirectory(title='Select Folder')
# path=pathlib.Path(__file__).parent.resolve() #Criar o endereço da pasta local.

#Ler o arquivo excel da lista de materiais.
df = pd.read_excel(filename,header = None)
df = df.astype(str) #Essencial para funcionar, torna toda a planilha string.
#df2 = df.copy()

#Listar arquivos com extensões determinadas.
types = ('*.pdf', '*.dxf', '*.step')
files_grabbed = []
for files in types:
    files_grabbed.extend(glob.glob(files))

###############################################
#Loop para processar todos os arquivos obtidos.
for ind, file in enumerate(files_grabbed):
    txt = os.path.join(path, file)
    
    #Extrair o nome do arquivo, remover extensão e isolar o número de série.
    x = txt.rfind(" ") #o rfind sempre acha a última ocorrência do caractere.
    if x == -1: #verificar casos onde o nome do arquivo não inclui o SAP.
        x = txt.rfind("\\")
    ponto = txt.rfind(".") #encontra o ponto antes da extensão.
    
    extensao = txt[ponto:] #isola a extensão para renomear o arquivo depois.
    serial = txt[:ponto] #remove a extensão.
    serial = serial[x+1:]
    
    #Cria uma duplicata do serial para renomear depois.
    serial2 = serial
    
    #Verificar casos tipo "123456_1", onde há múltiplas páginas.
    under = serial.rfind("_") 
    if under != -1:
        serial = serial[:under]
    
    #Extrair o código SAP do número de série em questão a partir da BOM.
    SAP = df.loc[df[5]==serial]
    if SAP.empty is True:
        SAP1 = "excluir" #Adiciona excluir ao nome de arquivos que não constam na lista.
    else:
        SAP1 = SAP.iloc[0,3]

    tupla = (SAP1,' - ',serial2, extensao)
    novo = ''.join(tupla) #Formação do novo nome do arquivo.
    
    os.rename(txt, novo) #Renomeia o arquivo.