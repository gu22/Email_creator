# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 19:15:42 2020

@author: gusan
"""
#--------------------------------------------------------------------------
#                           IMPORTs
#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
import tkinter as tk
import tkinter.ttk as ttk

import pygubu
from pygubu.widgets.pathchooserinput import PathChooserInput
from pygubu.widgets.scrollbarhelper import ScrollbarHelper


import win32com.client
import configparser

import time

import pandas as pd



#[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
#                                   ROTINAS
#[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

default = configparser.ConfigParser()


try:
     default.read('Configuracao.ini')
except:
    pass
    






#!!!!!!!!!!!!!!!!!!!!!!
#############################################################################
#                               VARIAVEIS
#############################################################################
padrao = default['DEFAULT']
padrao_email = default['Email']
padrao_planilha = default['Planilha']

# Set informações do arquivo Configuração.ini

link_forms = padrao['link_forms']

email_controle = padrao_email['controle']
email_remetente = padrao_email['remetente']

assunto_email = padrao_email['assunto']
corpo_email = padrao_email['corpo']
assinatura_email = padrao_email['assinatura']

coluna_chamado = padrao_planilha['colunaChamados']
coluna_email = padrao_planilha['colunaEmails']
coluna_assunto = padrao_planilha['colunaAssuntos']




text_output = ""
planilha_chamados = ""
email_chamado = ""
smart_paste = ""
excel = ""
# config_excel =""

X = 0
obs = 0
# excel = pd.read_excel('teste.xlsx')

#============================================================================
#                           TRATAMENTO DADOS
#===========================================================================

# colunas = list(excel.columns)






#========================================================================
#                       GUI
#------------------------------------------------------------------------
class Automail:
    def __init__(self, master=None):
        # build ui
        self.base_ds = ttk.Frame(master)
        self.forms_link = ttk.Entry(self.base_ds)
        self.forms_link.place(anchor='nw', relx='0.01', rely='0.05', width='400', x='0', y='0')
        self.mail_control = ttk.Entry(self.base_ds)
        self.mail_control.config(exportselection='true')
        self.mail_control.place(anchor='nw', relx='0.01', rely='0.21', width='400', x='0', y='0')
        self.label_forms = ttk.Label(self.base_ds)
        self.label_forms.config(text='Link Forms')
        self.label_forms.place(anchor='nw', relx='0.01', x='0', y='0')
        self.label_mail = ttk.Label(self.base_ds)
        self.label_mail.config(text='E-mails para relatórios')
        self.label_mail.place(anchor='nw', relx='0.01', rely='0.16', x='0', y='0')
        self.separator_2 = ttk.Separator(self.base_ds)
        self.separator_2.config(orient='horizontal')
        self.separator_2.place(anchor='nw', relx='0.0', rely='0.47', width='600', y='0')
        self.arquivo_ch = PathChooserInput(self.base_ds)
        self.arquivo_ch.config(type='file')
        self.arquivo_ch.place(anchor='nw', relx='0.01', rely='0.38', width='465', x='0', y='0')
        self.label_arquivo = ttk.Label(self.base_ds)
        self.label_arquivo.config(text='Excel com chamados')
        self.label_arquivo.place(anchor='nw', relx='0.01', rely='0.33', x='0', y='0')
        self.progressbar_1 = ttk.Progressbar(self.base_ds)
        self.progressbar_1.config(maximum='100', orient='vertical', value='0')
        self.progressbar_1.place(anchor='nw', height='190', relx='0.61', rely='0.51', width='15', x='0', y='0')
        self.button_ok = ttk.Button(self.base_ds)
        self.button_ok.config(text='OK')
        self.button_ok.place(anchor='nw', relx='0.86', rely='0.92', x='0', y='0')
        self.button_enviar = ttk.Button(self.base_ds)
        self.button_enviar.config(text='Enviar')
        self.button_enviar.place(anchor='nw', height='80', relx='0.75', rely='0.56', width='80', y='0')
        self.progressbar_2 = ttk.Progressbar(self.base_ds)
        self.progressbar_2.config(maximum='100', orient='horizontal')
        self.progressbar_2.place(anchor='nw', height='6', relx='0.75', rely='0.77', width='80', x='0', y='0')
        self.button_config = ttk.Button(self.base_ds)
        self.button_config.config(text='Configurações')
        self.button_config.place(anchor='nw', relx='0.66', rely='0.92', x='0', y='0')
        self.button_config.bind('<1>', self.callback, add='')
        self.forms_colar = ttk.Button(self.base_ds)
        self.forms_colar.config(text='Colar Inteligente')
        self.forms_colar.place(anchor='nw', relx='0.82', rely='0.38', x='0', y='0')
        self.scrollbarhelper_2 = ScrollbarHelper(self.base_ds, scrolltype='both')
        self.output = tk.Text(self.scrollbarhelper_2.container)
        self.output.config(height='10', relief='flat', state='normal', undo='false')
        self.output.config(width='50')
        
        self.output.place(anchor='nw', width='120', x='0', y='0')
        self.scrollbarhelper_2.add_child(self.output)
        self.scrollbarhelper_2.place(anchor='nw', height='190', relx='0.01', rely='0.51', width='355', x='0', y='0')
        self.base_ds.config(height='400', relief='flat', width='600')
        self.base_ds.pack(anchor='center', expand='false', side='top')
        self.base_ds.pack_propagate(0)

        # Main widget
        self.mainwindow = self.base_ds
        

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#                                   INICIALIZAÇÃO
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        self.forms_link.insert(0,link_forms)
        self.mail_control.insert(0, email_controle)
        
        self.button_enviar.config(state='disabled')





#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#                                       CHAMANDO FUNÇÕES
#///////////////////////////////////////////////////////////////////////////////////////////            

        self.button_ok.config(command=self.sair)
        self.button_config.config(command=self.setting)
        
        self.button_enviar.config(command=self.enviar)

        self.arquivo_ch.bind('<<PathChooserPathChanged>>', self.arquivo_excel)

#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#////////////////////////////////////////////////////////////////////////////////////////////






    def callback(self, event=None):
            pass

    def run(self):
            self.mainwindow.mainloop()






#--------------------------------------------------------------------------------------------
#                                           FUNÇÕES
#---------------------------------------------------------------------------------------------

    def sair(self):
        root.destroy()
        if obs ==1:
            root_config.destroy()




    def enviar(self):
        global corpo_email
        corpo_emailhtml = corpo_email.replace('\n', '<br>')
        outlook = win32com.client.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)


        email.BodyFormat= 1
        email.Subject= "assunto"
       
        email.HTMLBody= corpo_emailhtml
        email.Display(False)


        self.output.insert("0.0", corpo_email)


    def arquivo_excel(self,event=None):
        global excel,colunas,X,excel_config
        
        excel = self.arquivo_ch.cget('path')
        try:
            excel_config = pd.read_excel(excel)
            excel = pd.read_excel(excel)
            colunas = list(excel_config)
            X = 1
            print ('Arquivo carregado')
            self.button_enviar.config(state='enabled')
        except:
         print ("Arquivo não pode ser carregado")
         X = 0
         self.arquivo_ch.configure(path="")
         self.button_enviar.config(state='disabled')

#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    class AutomailConfig:
        def __init__(self, master=None):
            # build ui

            self.frame_5 = ttk.Frame(master)
            self.label_6 = ttk.Label(self.frame_5)
            self.label_6.config(anchor='center', font='{Arial} 20 {bold}', text='Configurações')
            self.label_6.place(anchor='nw', relx='0.29', rely='0.04', x='0', y='0')
            self.notebook_2 = ttk.Notebook(self.frame_5)
            self.frame_6 = ttk.Frame(self.notebook_2)
            self.n_chamado = ttk.Entry(self.frame_6)
            self.n_chamado.config(justify='center')
            self.n_chamado.place(anchor='nw', relx='0.23', rely='0.05', width='30', x='0', y='0')
            self.n_assunto = ttk.Entry(self.frame_6)
            self.n_assunto.config(justify='center')
            self.n_assunto.place(anchor='nw', relx='0.55', rely='0.05', width='30', x='0', y='0')
            self.n_email = ttk.Entry(self.frame_6)
            self.n_email.config(justify='center')
            self.n_email.place(anchor='nw', relx='0.89', rely='0.05', width='30', x='0', y='0')
            self.label_7 = ttk.Label(self.frame_6)
            self.label_7.config(text='Coluna Chamado')
            self.label_7.place(anchor='nw', relx='0.01', rely='0.05', x='0', y='0')
            self.label_8 = ttk.Label(self.frame_6)
            self.label_8.config(text='Coluna E-mail')
            self.label_8.place(anchor='nw', relx='0.7', rely='0.05', x='0', y='0')
            self.label_9 = ttk.Label(self.frame_6)
            self.label_9.config(text='Coluna Assunto')
            self.label_9.place(anchor='nw', relx='0.35', rely='0.05', x='0', y='0')
            self.scrollbarhelper_3 = ScrollbarHelper(self.frame_6, scrolltype='both')
            self.treeview_1 = ttk.Treeview(self.scrollbarhelper_3.container)
            self.treeview_1.pack(side='top')
            self.scrollbarhelper_3.add_child(self.treeview_1)
            self.scrollbarhelper_3.place(anchor='nw', relx='0.01', rely='0.15', width='480', x='0', y='0')
            self.frame_6.config(height='150', width='150')
            self.frame_6.place(anchor='nw', relx='0.0', x='0', y='0')
            self.notebook_2.add(self.frame_6, text='Planilha')
            self.frame_7 = ttk.Frame(self.notebook_2)
            self.notebook_1_2 = ttk.Notebook(self.frame_7)
            self.frame_1 = ttk.Frame(self.notebook_1_2)
            self.campo_de = ttk.Entry(self.frame_1)
            self.campo_de.place(anchor='nw', relx='0.21', rely='0.06', width='250', x='0', y='0')
            self.label_1 = ttk.Label(self.frame_1)
            self.label_1.config(text='De\Remetente:')
            self.label_1.place(anchor='nw', relx='0.01', rely='0.06', x='0', y='0')
            self.label_2 = ttk.Label(self.frame_1)
            self.label_2.config(text='Assunto:')
            self.label_2.place(anchor='nw', relx='0.08', rely='0.18', x='0', y='0')
            self.campo_assunto = ttk.Entry(self.frame_1)
            self.campo_assunto.place(anchor='nw', relx='0.21', rely='0.18', width='250', x='0', y='0')
            self.frame_1.config(height='200', width='200')
            self.frame_1.pack(side='top')
            self.notebook_1_2.add(self.frame_1, text='Envio')
            self.frame_2 = ttk.Frame(self.notebook_1_2)
            self.text_msg = tk.Text(self.frame_2)
            self.text_msg.config(blockcursor='false', cursor='arrow', font='TkDefaultFont', height='10')
            self.text_msg.config(insertunfocussed='none', setgrid='false', takefocus=False, width='50')
            self.text_msg.place(anchor='nw', height='235', relx='0.01', rely='0.01', width='460', x='0', y='0')
            self.frame_2.config(height='200', width='200')
            self.frame_2.pack(side='top')
            self.notebook_1_2.add(self.frame_2, text='Texto')
            self.frame_3 = ttk.Frame(self.notebook_1_2)
            self.text_ass = tk.Text(self.frame_3)
            self.text_ass.config(height='10', width='50')
            self.text_ass.place(anchor='nw', height='180', relx='0.01', rely='0.01', width='460', x='0', y='0')
            self.editor_online = ttk.Button(self.frame_3)
            self.editor_online.config(text='Editor Online')
            self.editor_online.place(anchor='nw', relx='0.81', rely='0.82', x='0', y='0')
            self.button_config_copy = ttk.Button(self.frame_3)
            self.button_config_copy.config(compound='top', text='Copiar')
            self.button_config_copy.place(anchor='nw', relx='0.64', rely='0.82', x='0', y='0')
            self.frame_3.config(height='200', width='200')
            self.frame_3.pack(side='top')
            self.notebook_1_2.add(self.frame_3, text='Assinatura')
            self.notebook_1_2.config(height='260', width='475')
            self.notebook_1_2.place(anchor='nw', relx='0.01', rely='0.05', x='0', y='0')
            self.frame_7.config(height='200', width='200')
            self.frame_7.pack(side='top')
            self.notebook_2.add(self.frame_7, state='normal', text='E-mail')
            self.notebook_2.config(height='280', width='490')
            self.notebook_2.place(anchor='center', relx='0.50', rely='0.52', x='0', y='0')
            self.button_config_ok = ttk.Button(self.frame_5)
            self.button_config_ok.config(text='OK')
            self.button_config_ok.place(anchor='nw', relx='0.81', rely='0.93', x='0', y='0')
            self.button_preview = ttk.Button(self.frame_5)
            self.button_preview.config(text='Preview')
            self.button_preview.place(anchor='nw', relx='0.02', rely='0.93', x='0', y='0')
            self.button_save = ttk.Button(self.frame_5)
            self.button_save.config(text='Save')
            self.button_save.place(anchor='nw', relx='0.64', rely='0.93', x='0', y='0')
            self.frame_5.config(height='400', width='500')
            self.frame_5.pack(side='top')

            # Main widget
            self.mainwindow = self.frame_5
#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#                                       CHAMANDO FUNÇÕES
#///////////////////////////////////////////////////////////////////////////////////////////

            self.button_save.config(command=self.save)
            
            self.button_config_ok.config(command=self.ok_config)
            
            

            # self.mainwindow.protocol("WM_DELETE_WINDOW",self.sair_config)

    #++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    #                                INICIALIZAÇÃO
    #++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            

            self.n_chamado.insert(0,coluna_chamado )
            self.n_assunto.insert(0,coluna_assunto )
            self.n_email.insert(0, coluna_email)

            self.campo_assunto.insert(0,assunto_email)
            self.campo_de.insert(0,email_remetente)

            self.text_msg.insert("0.0",corpo_email)
            self.text_ass.insert("0.0",assinatura_email)



            global excel_config,colunas,X

            if not X == 0:
                index_c = list(range(1,(1+(len(excel_config.columns)))))
                colunas_excel = list(index_c)
                
                self.treeview_1.config(selectmode="none")
                self.treeview_1['columns'] = colunas_excel
           
                for i in colunas_excel:
                    self.treeview_1.column(i, width=100,anchor='n')
                    self.treeview_1.heading(i, text=i)
                
                self.treeview_1.insert("",'end',text=0,values=colunas)
                    
                for index, row in excel_config.iterrows():
                    self.treeview_1.insert("",'end',text=(index+1),values=list(row))
        
                # definindo largura da coluna de INDEX
                self.treeview_1.column('#0',width=30)





        def run(self):
            self.mainwindow.mainloop()







#---------------------------------------------------------------------------------
#                               FUNÇÕES CONFIGURAÇÕES
#--------------------------------------------------------------------------------

        def sair_config():
            global obs
            obs = 0
            print('Fechou')

        def save(self):

            # Janela Principal
            padrao['link_forms'] = link_forms
            padrao_email['controle'] = email_controle
            # # Janela Configuraçoes
            # padrao_email['remetente']
            # padrao_email['assunto']
            # padrao_email['corpo']
            # padrao_email['assinatura']
            
            # padrao_planilha['colunaChamados']
            # padrao_planilha['colunaEmails']
            # padrao_planilha['colunaAssuntos']

            with open ('Configuracao.ini','w') as stg:
                default.write(stg)

        def ok_config(self):
            global obs
            root_config.destroy()
            obs = 0

    def setting(self):
        global root_config,link_forms,obs,email_controle
        print('ok')

        if obs ==0:
            root_config = tk.Tk()
            
            w2=550
            h2=300
            ws=root.winfo_screenwidth()
            hs=root.winfo_screenheight()
            x2=(ws/2)-(w2/2)
            y2=(hs/2)-(h2/2)
            root_config.geometry('+%d+%d'%(x2,y2))
            
            self.AutomailConfig(root_config)
            obs = 1

        ## HINT Get informações
            link_forms = self.forms_link.get()
            print (link_forms)
            email_controle = self.mail_control.get()




            def sair_config():
                global obs
                obs = 0
                print('Fechou')
                root_config.destroy()

            # try:
            root_config.protocol("WM_DELETE_WINDOW", sair_config)
            # except tk.TclError:
            #     print("saindo")
            #     pass








#============================================================================================
#                                           EXECUÇÃO
#============================================================================================

if __name__ == '__main__':

    root = tk.Tk()
    
    w=700
    h=400
    ws=root.winfo_screenwidth()
    hs=root.winfo_screenheight()
    x=(ws/2)-(w/2)
    y=(hs/2)-(h/2)
    root.geometry('+%d+%d'%(x,y))
    


    app = Automail(root)
    # appconfig = AutomailConfig(root)
    app.run()

# ========================================================================================
#HINT Helpers

    # Print todas as infos *pandas*
# with pd.option_context('display.max_rows', None, 'display.max_columns', None):  
    # print(df)
    
    # Adicionar linha inicio *pandas*
# excel.loc[-1] =
# excel.index = excel.index+1
# excel = excel.sort_index()
