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
        self.text = ""

    def text_block(self):
        unid = self.unid
        city = self.city
        state = self.state
        cnpj = self.cnpj
        insce = self.insce
        ende = self.ende
        self.text = """
            </head>
             
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Boa tarde,<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>&nbsp;<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Solicito o orçamento do item abaixo:<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>&nbsp;<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>&nbsp;<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Favor considerar 60 dias para faturamento.<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Datas fixas de pagamentos: 05, 10, 15, 20 e 25.<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>E-mail para envio de XML: nfe@rech.com<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Endereço de Entrega e Faturamento:<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>&nbsp;<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><b><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Unidade: {} - {}/{}</span></b><span style='mso-ascii-font-family:
            Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
            mso-bidi-font-family:Calibri;color:black;mso-fareast-language:PT-BR'><o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>CNPJ: {}<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>IE: {}<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>End.: {}<o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><a
            name="_MailAutoSig"><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
            "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
            color:black;mso-fareast-language:PT-BR'>&nbsp;</span></a><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'><o:p></o:p></span></p>
            
            <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black;
            mso-fareast-language:PT-BR'>Atenciosamente,<o:p></o:p></span></p>
            
            <p class=MsoNormal><o:p>&nbsp;</o:p></p>
            
            </div>

        """

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
                for terms in simple_list:
                    cod, cit, stt = terms
                    format_text = f'{cod} - {cit}/{stt}'
                self.head_store_list.append(format_text)

            return self.head_store_list
