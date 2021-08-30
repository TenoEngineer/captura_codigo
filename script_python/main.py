# -*- coding: utf-8 -*-
"""
Created on Tue Aug 24 12:16:05 2021

@author: heitor
"""
import os
from functions import LoopPast

path = 'N:\OBRAS_FINALIZADAS'
pastas = os.listdir(path)

#E SE EU TENTAR FAZER UM WHILE DE IR ENTRNAOD NAS PASTAS ATÉ ACHAR O CÓDIGO
camada_1=[]
for past in pastas:
    camada_1.append(LoopPast(path, past).find_past())   #Seleciona todas as pastas
for past in camada_1:
    camada_2=[]
    for file in past[1]:    #Dentro da pasta de uma obra, seleciona todas as pastas
        if LoopPast(past[0], file).find_past() is not None:
            camada_2.append(LoopPast(past[0], file).find_past())
    camada_3=[]
    for file in camada_2:
        LoopPast(file[0], file[1]).find_rm()
        camada_3.append(LoopPast(file[0], file[1]).find_rm())
    camada_4=[]       
    for file in camada_3:
        if file is not tuple:
            pass
        else:
            if LoopPast(file[0], file[1]).find_rm() is not None:
                camada_4.append(LoopPast(file[0], file[1]).find_rm())
    camada_5=[]       
    for file in camada_4:
        print(file)
        if file == True:
            next
        else:
            if LoopPast(file[0], file[1]).find_excel() is not None:
                camada_4.append(LoopPast(file[0], file[1]).find_excel())
                