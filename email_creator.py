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
c = 0
Data = datetime.datetime.now()
linkp = "https://forms.office.com/Pages/DesignPage.aspx#FormId=e7Oy_KBda0abgwAUtnp8eBa1jXSyJMNKlWZqCEOcTjFURTlSQjNKMjBXM0dCTEdMUEg1NE45NVIzTi4u&Preview=%7B%22PreviousTopView%22%3A%22None%22%7D&TopView=Preview"

# -- Base ----------------------------------------------------------
data = datetime.datetime.now()

arq = easygui.fileopenbox()

base = pd.read_excel(arq)
 
head = base.columns[0]
ncolunas = len(base.columns)
nlinhas = len(base.index)

resp = []

# config = pd.read_excel(arq,1)
# pesquisa = config.loc[0][0]
# nan_value = config.loc[0][1]

envio = []

base['Envio email'] = "nan"
for i in range(nlinhas):
    
    
    # Pegando valores da base dados
    campo_chamado = base.columns[0]
    numero_chamado = base.loc[i][0]

    cliente = base.loc[i][13]
    
    responsavel = base.loc[i][15]
    
    email_destino = base.loc[i][19]
    
    if str(email_destino) == "nan":
        base.iat[i,20] = "Falta email"
        continue
    else:
        base.iat[i,20] = "Em preparação"
    
    # filtrando os reponsaveis pelos chamados
    if not resp:
        resp.append(responsavel)
    elif responsavel not in resp:
        resp.append(responsavel)
        
       
    envio.append("<br>{}: {} {}<\br>".format(c,numero_chamado,email_destino))
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
    
    
    
    text = """<p style="margin: 0cm; margin-bottom: .0001pt;">Prezado {}, tudo bem?</p>
    <p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
    <p style="margin: 0cm; margin-bottom: .0001pt;">Meu contato &eacute; referente uma pesquisa de satisfa&ccedil;&atilde;o relacionada ao chamado <span style="background-color: #ffcc00;">{}</span> que foi finalizado, ela serve para nos ajudar na melhoria dos nossos atendimentos, leva menos que 05 minutos, pode nos ajudar?</p>
    <p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
    <p style="margin: 0cm; margin-bottom: .0001pt;">Link: <a href="{}">{}</a></p>
    <p>Atenciosamente / Best regards,</p>
    <p>&nbsp;<br /><span style="font-size: 20px;"><strong><span style="color: #de0043;"><em>Pesquisa - Teste</em></span></strong></span> <br /><span style="color: #10384f;"><strong>Distribution CP </strong></span></p>
    <p><br /><span style="color: #3adeff; letter-spacing: -1px; font-size: 16px;"><strong>////////////////////</strong></span></p>
    <p><br />Bayer Brazil &ndash; Crop Science</p>
    <p>Rua Domingos Jorge, 1100 |</p>
    <p>Web: http://www.bayer.com</p>
    
    
    
    """.format(saudacao_email,numero_chamado,linkp,linkp)
    # """.format(saudacao_email,numero_chamado,pesquisa)


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

    assunto = "Pesquisa - teste 1"
    
    
    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    # email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
    email.To= email_destino
    email.BodyFormat= 2
    email.Subject= assunto
    # email.Subject= email.Subject.replace('[compName]','test')
    email.HTMLBody= (text)



#email - anexos
# email.HTMLBody= email.HTMLBody.replace('fname', 'test')
# for x in email_path:
#     email.Attachments.Add(Source= os.path.join(item,x))
# signature = easygui.fileopenbox()
# sign = win32com.client.Dispatch('Word.Application')
# doc = app.Documents.Open(r'D:\winGUI\test\1.doc')
# doc = sign.Documents.Open(signature)
# doc.Content.Copy()
# doc.Close()
   
# email.GetInspector.WordEditor.Range(Start=0, End=0).Paste()

#email - exibição
    # email.Display(False)


    email.SaveAs("Teste_pesquisa - {}.msg".format(c),3)
    base.iat[i,20] = "OK"
    c+=1
    email.Send()
     

nenvios = len(envio)    
insert = ("").join(envio)

data_rg = (str(Data.strftime("%Y.%m.%d_%H.%M.%S")))

texto_verificacao = ("""

<p>=================== Verifica&ccedil;&atilde;o de Envios ================</p>
{}
<p>{}, Total de envios: {}</p>

""").format(insert,data_rg,nenvios)

assunto_verificacao = "Teste - relatorio_email"

emailv = outlook.CreateItem(0)
# email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
emailv.To= 'gustavo.dossantos@bayer.com;beatriz.goncalves@bayer.com'
emailv.BodyFormat= 2
emailv.Subject= assunto_verificacao
# email.Subject= email.Subject.replace('[compName]','test')
emailv.HTMLBody= (texto_verificacao)   


emailv.SaveAs("{}.msg".format(assunto_verificacao),3)
emailv.Send()   

base.to_excel("Final.xlsx",index=False)
# -----------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------

"""<p>{}!</p>
    <p>Estamos enviando este e-mail para avaliar como foi o tratamento do chamado encerrado.</p>
    <p>&nbsp;Chamado: <span style="background-color: #ffff00; font-size: 28px;">{}</span> --&gt; utilizar esse numero no primeiro item da pesquisa</p>
    <p>&Eacute; de extrema importancia que voce complete a pesquisa, para melhorar o processo.</p>
    <b><a title="Pesquisa" href="{}">Pesquisa</a><\b>
    <p>Atenciosamente,</p>
    <p>Equipe Customer.</p>
    <p>&nbsp;</p>
    <p>Atenciosamente / Best regards,</p>
    <p>&nbsp;<br /><span style="font-size: 20px;"><strong><span style="color: #de0043;"><em>Pesquisa - Teste</em></span></strong></span> <br /><span style="color: #10384f;"><strong>Distribution CP </strong></span></p>
    <p><br /><span style="color: #3adeff; letter-spacing: -1px; font-size: 16px;"><strong>////////////////////</strong></span></p>
    <p><br />Bayer Brazil &ndash; Crop Science</p>
    <p>Rua Domingos Jorge, 1100 |</p>
    <p>Web: http://www.bayer.com</p>"""