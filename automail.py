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
from tkinter import Toplevel
from pygubu.widgets.pathchooserinput import PathChooserInput
from pygubu.widgets.scrollbarhelper import ScrollbarHelper
from pygubu.widgets.dialog import Dialog
from pygubu.widgets.scrollbarhelper import ScrollbarHelper


#############################################################################
#                               VARIAVEIS
#############################################################################

link_forms = ""
email_controle = ""
email_remetente = ""

corpo_email = ""
assinatura_email = ""











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
        # TODO - self.arquivo_ch: code for custom option 'textvariable' not implemented.
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
        _text_ = '''Output'''
        self.output.insert('0.0', _text_)
        self.output.place(anchor='nw', width='120', x='0', y='0')
        self.scrollbarhelper_2.add_child(self.output)
        # TODO - self.scrollbarhelper_2: code for custom option 'usemousewheel' not implemented.
        self.scrollbarhelper_2.place(anchor='nw', height='190', relx='0.01', rely='0.51', width='355', x='0', y='0')
        self.base_ds.config(height='400', relief='flat', width='600')
        self.base_ds.pack(anchor='center', expand='false', side='top')
        self.base_ds.pack_propagate(0)

        # Main widget
        self.mainwindow = self.base_ds

#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#                                       CHAMANDO FUNÇÕES
#///////////////////////////////////////////////////////////////////////////////////////////            

        self.button_ok.config(command=self.sair)
        self.button_config.config(command=self.setting)




#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#////////////////////////////////////////////////////////////////////////////////////////////






    def callback(self, event=None):
            pass

    def run(self):
            self.mainwindow.mainloop()
            




#--------------------------------------------------------------------------------------------
#                                           FUNÇÕES
#---------------------------------------------------------------------------------------------
    # def sair(self):

    #     Automail().destroy()
    #     print("sair...")

    def sair(self):
        root.destroy()

    # def configuracao(self):
    #     top = Toplevel(AutomailConfig())



#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!111
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    class AutomailConfig:
        def __init__(self, master=None):
            # build ui
            # self.configuracao = Dialog(master)
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
            # TODO - self.scrollbarhelper_3: code for custom option 'usemousewheel' not implemented.
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
            # self.configuracao.config(height='100', modal='false', width='200')
    
            # Main widget
            self.mainwindow = self.frame_5
    
    
        def run(self):
            self.mainwindow.mainloop()



    def setting(self):
        print('ok')
        root_config = tk.Tk()
        root_config.geometry('+%d+%d'%(x,y))
        opcoes = self.AutomailConfig(root_config)
        





    






#============================================================================================
#                                           EXECUÇÃO
#============================================================================================

if __name__ == '__main__':
    import tkinter as tk
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


