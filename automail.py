# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 19:15:42 2020

@author: gusan
"""


import tkinter as tk
import tkinter.ttk as ttk
from pygubu.widgets.pathchooserinput import PathChooserInput
from pygubu.widgets.scrollbarhelper import ScrollbarHelper


class GuiMail02App:
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
        
    def callback(self, event=None):
            pass

    def run(self):
            self.mainwindow.mainloop()
            
#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#                                       CHAMANDO FUNÇÕES
#///////////////////////////////////////////////////////////////////////////////////////////            

#--------------------------------------------------------------------------------------------
#                                           FUNÇÕES
#---------------------------------------------------------------------------------------------

        # def 


#============================================================================================
#                                           EXECUÇÃO
#============================================================================================

if __name__ == '__main__':
    import tkinter as tk
    root = tk.Tk()
    app = GuiMail02App(root)
    app.run()


