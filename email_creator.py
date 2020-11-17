# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 09:58:24 2020

@author: glpyz
"""
# --- MODULOS E IMPORTS ---------------------------------

import datetime
#import os.path

import easygui
from easygui import *

import openpyxl
from openpyxl import load_workbook
import pandas as pd

# import shutil
# from shutil import copyfile

from datetime import date
import sys
#EMAIL
import win32com.client

# -------------------------------------------------------------------
#----------------------------------------------------------------------
c = 0
Data = datetime.datetime.now()

print(Data)
print ("Automação para enviar emails (Pesquisa de Satisfação)")
print("v 1.03\n")


linkd = "https://forms.office.com/Pages/ResponsePage.aspx?id=e7Oy_KBda0abgwAUtnp8eBa1jXSyJMNKlWZqCEOcTjFURTlSQjNKMjBXM0dCTEdMUEg1NE45NVIzTi4u"
# linkp = easygui.enterbox(msg='Entre com o link do Forms', title='Link Forms ', default=linkd, strip=True)
# email_confirmacao = easygui.enterbox(msg='Entre com os e-mails para recebe a confirmação\n Separar os e-mails com ; (ponto e virgula)', title='E-mail para controle ', default='', strip=True)
# assunto = easygui.enterbox(msg='Qual assunto/titulo do e-mail ?', title='Assunto - E-mail ', default='Pesquisa de Satisfação', strip=True)
# easygui.enterbox

campos = ["Link Forms","Assunto","E-mail Origem\Remetente","E-mail Relatorio"]
default_valores = [linkd,"Pesquisa de Satisfação","",""]




#---------------------------------------------------------------------------
# -- Base ----------------------------------------------------------




valores = easygui.multenterbox("oi","poe ae",campos,default_valores)

linkp = valores[0]
assunto = valores[1]
email_sender = valores[2]
email_confirmacao = valores[3]



data = datetime.datetime.now()

arq = easygui.fileopenbox("Selecionar arquivo",title="Arquivo Excel com RTV\e-mails",filetypes=["*.xlsx"],multiple=False)

base = pd.read_excel(arq)
 
head = base.columns[0]
ncolunas = len(base.columns)
nlinhas = len(base.index)

resp = []

# config = pd.read_excel(arq,1)
# pesquisa = config.loc[0][0]
# nan_value = config.loc[0][1]

envio = []
recusa = []

base['Envio email'] = "nan"


if ccbox("Iniciar Processamento?", "Inicar processo para enviar e-mails"):
    pass
else:
    sys.exit(0)

for i in range(nlinhas):
    
    
    # Pegando valores da base dados
    campo_chamado = base.columns[0]
    try:
        numero_chamado = int(base.loc[i][0])
    except:
        base.iat[i,ncolunas] = "Falta n° chamado"
        continue
    
    nome_chamado = base.loc[i][20]

    cliente = base.loc[i][13]
    
    responsavel = base.loc[i][15]
    
    email_destino = base.loc[i][19]
    
    if str(nome_chamado) == 'nan':
        base.iat[i,ncolunas] = "Falta Assunto do chamado"
        recusa.append(" {};".format(numero_chamado))
        continue
    
    if str(email_destino) == "nan":
        base.iat[i,ncolunas] = "Falta email"
        recusa.append(" {};".format(numero_chamado))
        continue
    else:
        base.iat[i,ncolunas] = "Em preparação"
    
    # filtrando os reponsaveis pelos chamados
    if not resp:
        resp.append(responsavel)
    elif responsavel not in resp:
        resp.append(responsavel)
        
    print ("{}: {} {}\n".format(c,numero_chamado,email_destino)) 
    print ("==============================================\n")
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
    
    
    
    text = """<p style="margin: 0cm; margin-bottom: .0001pt;">Prezado, {}!</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">Tudo bem?</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">Recentemente voc&ecirc; entrou em contato com a Central de Intera&ccedil;&atilde;o referente ao chamado <strong>{} - {}</strong>.</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">Poderia, por gentileza, nos ajudar na melhoria dos nossos atendimentos respondendo essa pesquisa que leva menos de 2min?</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">Link: <a href="{}">{}</a></p>
<p style="margin: 0cm; margin-bottom: .0001pt;">&nbsp;</p>
<p style="margin: 0cm; margin-bottom: .0001pt;">Agradecemos a colabora&ccedil;&atilde;o!</p>
<p>Atenciosamente / Best regards,</p>
<p>&nbsp;<br /><span style="font-size: 20px; ont-family: Arial;"><strong><span style="color: #de0043;"><em>Central de Intera&ccedil;&atilde;o</em></span></strong></span>&nbsp;<br /><span style="color: #10384f;"><strong>Customer Interaction</strong></span></p>
<p><span style="color: #3adeff; letter-spacing: -1px; font-size: 16px;"><strong>////////////////////</strong></span></p>
<p>Bayer Brazil &ndash; Crop Science</p>
<p>Contatos:<br />Tel: 0800 940 6000 (Op&ccedil;&atilde;o 1)<br />E-mail: <a href="mailto:cal.monsanto.brasil@monsanto.com">cal.monsanto.brasil@monsanto.com</a>&nbsp;<br />Web: <a href="http://www.bayer.com">http://www.bayer.com&lt;\a&gt;</a></p>
<p>&nbsp;</p>
    
    
    
    """.format(saudacao_email,numero_chamado,nome_chamado,linkp,linkp)
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

    # assunto = "Pesquisa - teste 1"
    
    
    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    # email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
    email.To= email_destino
    if email_sender == "":
        pass
    else:
         email.SentOnBehalfOfName= email_sender
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
#HINT
#email - exibição
    email.Display(False)


    # email.SaveAs("{} - {} - {}.msg".format(assunto,c,numero_chamado),3)
    base.iat[i,ncolunas] = "OK"
    c+=1
    #HINT
    # email.Send()
     

nenvios = len(envio)    
insert = ("").join(envio)
insert_recusa = ("").join(recusa)
data_rg = (str(Data.strftime("%Y.%m.%d_%H.%M.%S")))

texto_verificacao = ("""

<p>=================== Verifica&ccedil;&atilde;o dos Envios ================</p>
{}<p>Chamados não enviados: {}<\p>
<p>{}, Total de itens enviados: {}</p>

""").format(insert,insert_recusa,data_rg,nenvios)

assunto_verificacao = "Relatorio_email"

emailv = outlook.CreateItem(0)
# email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
emailv.To= email_confirmacao
emailv.BodyFormat= 2
emailv.Subject= assunto_verificacao
# email.Subject= email.Subject.replace('[compName]','test')
emailv.HTMLBody= (texto_verificacao)   


# emailv.SaveAs("{} - {}.msg".format(assunto_verificacao,data_rg),3)


#HINT
emailv.Display(False)

if not email_confirmacao == "":
    emailv.Send()   

base.to_excel("Verificação - {}.xlsx".format(data_rg),index=False)

easygui.msgbox("Finalizado")
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