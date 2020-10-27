# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 09:58:24 2020

@author: glpyz
"""
# --- MODULOS E IMPORTS ---------------------------------

import datetime
import os.path

import easygui
from easygui import *

import openpyxl
from openpyxl import load_workbook
import pandas as pd

import shutil
from shutil import copyfile

from datetime import date
import sys
#EMAIL
import win32com.client

# -------------------------------------------------------------------
#----------------------------------------------------------------------


# -- Base ----------------------------------------------------------
data = datetime.datetime.now()

arq = easygui.fileopenbox()

base = pd.read_excel(arq)
 
head = base.columns[0]
ncolunas = len(base.columns)
nlinhas = len(base.index)

resp = []

config = pd.read_excel(arq,1)
pesquisa = config.loc[0][0]
# nan_value = config.loc[0][1]

base['Envio email'] = "nan"
for i in range(nlinhas):
    
    
    # Pegando valores da base dados
    campo_chamado = base.columns[0]
    numero_chamado = base.loc[i][0]

    cliente = base.loc[i][13]
    
    responsavel = base.loc[i][15]
    
    email = base.loc[i][19]
    
    if str(email) == "nan":
        base.iat[i,20] = "Falta email"
        continue
    else:
        base.iat[i,20] = "Em preparação"
    
    # filtrando os reponsaveis pelos chamados
    if not resp:
        resp.append(responsavel)
    elif responsavel not in resp:
        resp.append(responsavel)
        
       
#-------------------------------------------------------------------
#----------------------------------------------------------------------

    orientacao_dia = int(data.strftime("%H"))
    saudacao=["Bom dia ","Boa Tarde ","Boa noite "]    
    
    if orientacao_dia <=11:
        saudacao_email = saudacao[0]
    elif orientacao_dia <=17:
        saudacao_email = saudacao[1]
    else:
        saudacao_email = saudacao[2]
    
    
    
    text = """<p>{}!</p>
    <p>Estamos enviando este e-mail para avaliar como foi o tratamento do chamado encerrado.</p>
    <p>&nbsp;Chamado: <span style="background-color: #ffff00; font-size: 28px;">{}</span> --&gt; utilizar esse numero no primeiro item da pesquisa</p>
    <p>&Eacute; de extrema importancia que voce complete a pesquisa, para melhorar o processo.</p>
    <b><a title="Pesquisa" href="{}">Pesquisa</a><\b>
    <p>Atenciosamente,</p>
    <p>Equipe Customer.</p>
    <p>&nbsp;</p>""".format(saudacao_email,numero_chamado,pesquisa)

base.to_excel("Final.xlsx",index=False)
# base.save("Final2.xlsx")
# --- Criando os EMAILS -------------------------------------


    
# orientacao_dia = int(data.strftime("%H"))
# saudacao=["Bom dia, ","Boa Tarde, ","Boa noite, "]    

# if orientacao_dia <=11:
#     saudacao_email = saudacao[0]
# elif orientacao_dia <=17:
#     saudacao_email = saudacao[1]
# else:
#     saudacao_email = saudacao[2]
   
   
# email_path = os.listdir(item)

# for x in meses:
#     if meses[x] in item.lower():
#         mes = meses[x]
#         print(mes)
#     else:
#         pass

# mes = mes.capitalize()  

# trasp_email = os.path.split(item)
# print(trasp_email)
# trasp_email = trasp_email[-1]
# print(trasp_email)

# outlook = win32com.client.Dispatch('Outlook.Application')
# email = outlook.CreateItem(0)
# # email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
# email.To= ''
# email.BodyFormat= 2
# email.Subject="Avalição dos Fornecedores - Base de Dados (%s) - %s"%(trasp_email,mes)
# # email.Subject= email.Subject.replace('[compName]','test')
# email.HTMLBody= (
# """{} 
#     esperamos que todos estejam bem!<p>
    
#     Nossa reunião de avaliação está próxima!<br>
#     Estamos disponibilizando a base de dados referente ao mês de <b> {}  </b> para que vocês possam
#     analisar e nos informar o que ocorreu.<p> 
#     Para qualquer tipo de dúvida, estamos à disposição!<br>"""%(saudacao_email,mes)
    
    
#     )

# #email - anexos
# # email.HTMLBody= email.HTMLBody.replace('fname', 'test')
# for x in email_path:
#     email.Attachments.Add(Source= os.path.join(item,x))

   


# #email - exibição
#     email.Display(False)


# email.SaveAs(item+ '\\EMAIL- %s_%s.msg'%(trasp_email,mes), 3)

            
# -----------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------