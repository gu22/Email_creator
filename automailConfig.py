# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 23:00:41 2020

@author: gusan
"""

import tkinter as tk
import tkinter.ttk as ttk
from pygubu.widgets.dialog import Dialog
from pygubu.widgets.scrollbarhelper import ScrollbarHelper


class GuiMail02App:
    def __init__(self, master=None):
        # build ui
        self.dialog_3 = Dialog(master)
        self.frame_5 = ttk.Frame(self.dialog_3.toplevel)
        self.label_6 = ttk.Label(self.frame_5)
        self.label_6.config(anchor='center', font='{Arial} 20 {bold}', text='Configurações')
        self.label_6.place(anchor='nw', relx='0.29', rely='0.04', x='0', y='0')
        self.notebook_2 = ttk.Notebook(self.frame_5)
        self.frame_6 = ttk.Frame(self.notebook_2)
        self.n_chamado = ttk.Entry(self.frame_6)
        self.n_chamado.config(justify='center')
        _text_ = '''
'''
        self.n_chamado.delete('0', 'end')
        self.n_chamado.insert('0', _text_)
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
        self.notebook_2.bind('', self.callback, add='')
        self.button_config_ok = ttk.Button(self.frame_5)
        self.button_config_ok.config(text='OK')
        self.button_config_ok.place(anchor='nw', relx='0.81', rely='0.93', x='0', y='0')
        self.button_preview = ttk.Button(self.frame_5)
        self.button_preview.config(text='Preview')
        self.button_preview.place(anchor='nw', relx='0.02', rely='0.93', x='0', y='0')
        self.button_save = ttk.Button(self.frame_5)
        self.button_save.config(text='Save')
        self.button_save.place(anchor='nw', relx='0.64', rely='0.93', x='0', y='0')
        self.frame_5.config(height='400', text='E-mail', width='500')
        self.frame_5.pack(side='top')
        self.dialog_3.config(height='100', modal='false', width='200')

        # Main widget
        self.mainwindow = self.dialog_3

    def callback(self, event=None):
        pass

    def run(self):
        self.mainwindow.mainloop()

if __name__ == '__main__':
    import tkinter as tk
    root = tk.Tk()
    app = GuiMail02App(root)
    app.run()

