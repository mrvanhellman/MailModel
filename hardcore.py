import win32com.client as win32
import csv
outlook = win32.Dispatch("Outlook.Application")

# Global variable
unit_picker = ''


class SaveMail:
    def __init__(self, mail_body, subject, reference):
        self.mail = outlook.CreateItem(0)
        self.body = mail_body
        self.subject = f'Solicitação de Orçamento - {subject} - {reference}'

    def new_mail(self):
        self.mail.Subject = self.subject
        self.mail.HTMLBody = self.body
        self.mail.Save()


class MailBody:

    def __init__(self, unid, city, state, cnpj, insce, ende):
        self.unid = unid
        self.city = city
        self.state = state
        self.cnpj = cnpj
        self.insce = insce
        self.ende = ende

    def text_block(self):
        unid = self.unid
        city = self.city
        state = self.state
        cnpj = self.cnpj
        insce = self.insce
        ende = self.ende
        self.text = """
        <div>
        <p>
        Segue os dados para faturamento e entrega: </P
        <div>
        Unidade: {} - {}-{}</div>
        <p>
        Cnpj: {}
        </p>
        <p>
        IE: {}
        </p>
        <p>
        Endereço: {}
        </p>
        <div>
        
        <p>Faturamento para 60 dias.</p>
        <p>Datas fixas de pagamentos: 05, 10, 15, 20 e 25.</p>
        <p>O envio das XML : nfe@rech.com</p>
        <p>Entregar para: </p>

        """

        # print(text.format(unid, city, state, cnpj, insce, ende))
        return self.text.format(unid, city, state, cnpj, insce, ende)


class StoreList:

    def __init__(self):
        self.unid = ''
        self.city = ''
        self.state = ''
        self.cnpj = ''
        self.insE = ''
        self.ende = ''
        self.full_store_list = []
        self.head_store_list = []

    def store_list(self):
        with open("stores.csv", "r") as csvfile:
            dirt_file = csv.reader(csvfile, delimiter=' ', quotechar='|')
            simple_list = []
            for row in dirt_file:
                list_rows = list(row)
                simple_list.append(list_rows)
            self.full_store_list = simple_list
            return self.full_store_list

    def store_head(self):
        with open("stores.csv", "r") as csvfile:
            dirt_file = csv.reader(csvfile, delimiter=' ', quotechar='|')
            simple_list = []
            for row in dirt_file:
                list_rows = list(row)
                simple_list.append(list_rows[0:3])
                for row in simple_list:
                    cod, cit, stt = row
                    formated_text = f'{cod} - {cit}/{stt}'
                self.head_store_list.append(formated_text)

            return self.head_store_list






