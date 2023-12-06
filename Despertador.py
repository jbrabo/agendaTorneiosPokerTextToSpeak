# -*- coding: utf-8 -*-
"""
Created on Tue Oct 31 14:52:29 2023

@author: jvBra
"""
#BIBLIOTECAS PARA MANIPULAÇÃO DE DADOS
import pandas as pd

#BIBLIOTECAS MANIPULAÇÃO DE TEMPO
from time import localtime, sleep

#BIBLIOTECAS TextToSpeak(TTS)
import pyttsx3

#BIBLIOTECA PARA TRABALHAR COM DIRETÓRIOS
import os                                                
def ler_Agenda(arquivo,sheetName):
#   CRIANDO VARIAVEL DO TIPO LEITURA DE EXCEL NO PANDAS, ISSO RETORNA UMA TABELA
    silver_plan = pd.read_excel(arquivo,sheet_name=sheetName)

    #ORDENANDO TABELA ATAVÉS DA HORA DE REGISTRO
    silver_plan.sort_values('Hora de Registro',inplace=True)


    #RETORNAR VALOR ATUALIZADO PARA SILVER PLAN
    return silver_plan    

def agendar(tabela,treeV):
    #ABRINDO A PASTA E CARREGANDO A PLAN DE TORNEIOS    
    h_torneios=tabela["Hora de Registro"]
    d_torneios=tabela["Descrição"]
    s_torneios=tabela["Site"]
    p_torneios=tabela["Prioridade"]
    b_torneios=tabela["Buy In"]

    #REOORDENANDO A PLAN POR ORDEM DA HORA DE REGISTRO, 
    #PARA NÃO OCORRER PERIGO DE NÃO AVISAR UM TORNEIO NA HORA 
    #CASO NÃO FOR ALINHADA ANTES

    #COMPACTANTO OS DADOS PARA ITERAR
    torneios = zip(h_torneios,d_torneios,s_torneios, p_torneios, b_torneios)

    #CRIANDO UM PONTEIRO PARA CÁLCULAR TEMPO DE GRIND
    hi = localtime().tm_hour
    mi = localtime().tm_min

    #INICIALIZANDO O MEU OBJETO TTS
    r_tts = pyttsx3.init()    
    r_tts.say(f"Bom dia João Brabo, iniciando o Grind as {hi} e {mi}")
    #INICIANDO LAÇO DE REPETIÇÃO PARA PERCORRER TODOS OS EVENTOS DA PLAN
    #PODERIAMOS TRABALHAR QUALQUER AGENDA DESSA MANEIRA
    pontTree = 0
    for h, d, s, p, b in torneios:

        #percorrendo treeView com o Focus       
        id_ = treeV.get_children()[pontTree]
        treeV.focus(id_)
        treeV.selection_set(id_)
        pontTree +=1
        #SELECIONAR LINHA

        #ANUNCIA QUAL O TORNEIO ESTÁ NA VEZ
        msg_next = f"Próximo Torneio : {b} {d}, no Site: {s}, às {h.hour} horas e {h.minute} minutos"
        #print(msg_next)
        r_tts.say(msg_next)
        r_tts.runAndWait()

        #INICIA TESTE COM O TORNEIO DA VEZ
        while True:
            if localtime().tm_hour > h.hour:
                break
                
            elif localtime().tm_hour == h.hour:        

                    if localtime().tm_min > h.minute:
                        break
                        
                    elif localtime().tm_min < h.minute:
                        
                        #CASO NÃO PASSOU A HORA, NEM PASSOU OS MINUTOS
                        #ELE CONTINUA DENTRO DO WHILE E RETORNA PRO INICIO DO TESTE
                        sleep(5)
                        
                    else:
                        
                        #SE OS MINUTOS FOREM IGUAIS
                        msg_r = f"Hora de se Registrar no Torneio: {b} {d}, no Site: {s}. A prioridade desse Torneio é: {p}"
                        print(msg_r)
                        print(f'{h.hour}:{h.minute}')
                        r_tts.say(msg_r)
                        r_tts.runAndWait()
                        sleep(5)
                        break
                        
            else:
                continue
            #r_dptd.init()
    #FINALIZANDO O PONTEIRO PARA CÁLCULAR TEMPO DE GRIND       
    hf = localtime().tm_hour
    mf = localtime().tm_min

    #MSG DE FIM DE GRIND
    r_tts.say(f"Finalizando Registros às {hf} e {mf}")
    r_tts.runAndWait()       

#arquivo = "C:\\Users\\jvBra\\OneDrive\\Documents\\2023\\Python Scripts\\SilverPlan\\Data\\Torneios.xlsx"
#sheet = "SilverPlan_2023_Domingo"
#agendar(ler_Agenda(arquivo,sheet))


