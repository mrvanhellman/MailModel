from tkinter import *
from tkinter import ttk
from hardcore import StoreList, MailBody, SaveMail

unit_picker = ''


class Interface:

    def __init__(self, master, shop_list):
        myframe = Frame(master)
        myframe.grid(column=0, row=0, sticky="NSEW")

        self.shop_list = shop_list

        self.welcome_message = Label(text='e-Mail Models v0.1', width=35, height=2)
        self.welcome_message.grid(row=0, column=0, columnspan=2, sticky="NSEW")

        self.msg1 = Label(text='Escolha a unidade:')
        self.msg1.grid(row=1, column=0)

        self.msg2 = Label(text='O que você está orçando?')
        self.msg2.grid(row=2, column=0)

        self.msg3 = Label(text='Chamado ou Referência:')
        self.msg3.grid(row=3, column=0)

        self.store_pick = ttk.Combobox(master, values=shop_list)
        self.store_pick.grid(row=1, column=1)

        self.subject_input = Entry()
        self.subject_input.grid(row=2, column=1)

        self.ref_input = Entry()
        self.ref_input.grid(row=3, column=1)

        self.button_save = Button(text="Salvar", command=self.save_model)
        self.button_save.grid(row=4, column=0, columnspan=2)

        self.msg5 = Label(text='')
        self.msg5.grid(row=5, column=0, columnspan=2)

        self.proxy_file = StringVar
        self.unit_pick = StringVar

        self.unid = ""
        self.city = StringVar
        self.state = StringVar
        self.cname = ""
        self.cnpj = StringVar
        self.insce = StringVar
        self.ende = StringVar

    def save_model(self):
        store = self.store_pick.get()
        subject = self.subject_input.get()
        reference = self.ref_input.get()
        self.unit_pick = store[0:6]
        self.msg5['text'] = 'Modelo salvo em Rascunhos'
        text = StorePicker(self.unit_pick)
        text.unit_selected()
        full_list = StoreList()
        self.list_picker(full_list.store_list())
        self.cname = self.company_name(self.unid)
        mail_body = MailBody(self.unid, self.city, self.state, self.cname, self.cnpj, self.insce, self.ende)
        new_mail = SaveMail(mail_body.text_block(), subject, reference)
        new_mail.new_mail()

    def list_picker(self, store_list):
        for sub_list in store_list:
            if unit_picker in sub_list:
                store_index = store_list.index(sub_list)
                self.unid, self.city, self.state, self.cnpj, self.insce, self.ende = store_list[store_index]
                return self.unid, self.city, self.state, self.cnpj, self.insce, self.ende

    def company_name(self, unid):
        if unid.startswith('40'):
            return "Rech Importadora e Distribuidora S.A."
        elif unid.startswith('60'):
            return "Rech Agricola S/A"
        elif unid.startswith('TE'):
            return "Telmac Comércio Importação e Exportação Ltda"


class StorePicker:

    def __init__(self, unit):
        self.unit_picker = unit

    def unit_selected(self):
        global unit_picker
        unit_picker = self.unit_picker
        return unit_picker
