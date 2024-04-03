import PyPDF2
import pandas as pd
import re
import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

diretorio = 'Z:\Scripts\Enviar emails'
diretorio_enviados = 'Z:\Enviados'
tabela = pd.read_csv('Cadastros.csv')
print(tabela.columns)

def extrair_cnpj_pdf(arquivo_pdf):
    try:
        with open(arquivo_pdf, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            page = reader.pages[0]
            text = page.extract_text()
            cnpj = re.search(r'[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}', text)
            if cnpj:
                return cnpj.group()
            return ''
    except Exception as e:
        print('Deu ruim logo no começo')
            
smtp_server = 'mail.temporario.net'
smtp_port = 587
smtp_username = 'teste@temporario.net'
smtp_password = 'SenhaTop'

lista_envio = []

for arquivo_pdf in os.listdir(diretorio):
    if arquivo_pdf.endswith('.pdf'):
        try:
            cnpj = extrair_cnpj_pdf(os.path.join(diretorio, arquivo_pdf))
            cnpj_values = tabela.get('CNPJ', [])
            if cnpj in cnpj_values:
                email_destinatario = tabela[tabela['CNPJ']==cnpj]['Email'].values[0]
                arquivo = {
                    'nome': arquivo_pdf,
                    'caminho': os.path.join(diretorio, arquivo_pdf),
                    'cnpj': cnpj,
                    'destinatario': email_destinatario
                }
                lista_envio.append(arquivo)
                
        except Exception as e:
            print('Erro ao processar arquivo', arquivo_pdf, ':', str(e))

for arquivo in lista_envio:
    try:
        email_destinatario = arquivo['destinatario']
        message = MIMEMultipart()
        message['From'] = smtp_username
        message['To'] = email_destinatario
        message['Subject'] = 'Arquivo PDF'
        corpo = 'Testando o envio de arquivos em PDF.'
        body = MIMEText(corpo, 'html')
        message.attach(body)

        with open(arquivo['caminho'], 'rb') as f:
            attach = MIMEApplication(f.read(),_subtype = 'pdf')
            attach.add_header('Content-Disposition','attachment',filename=str(arquivo['nome']))
            message.attach(attach)
            
        smtp = smtplib.SMTP(smtp_server, smtp_port)
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)
        smtp.sendmail(smtp_username, email_destinatario, message.as_string())
        print(f"Arquivo {arquivo['nome']} enviado para {email_destinatario}")
        smtp.quit()                
        
        shutil.move(arquivo['caminho'], os.path.join(diretorio_enviados, arquivo['nome']))

    except Exception as e:
        print('Erro ao enviar email para', email_destinatario, ':', str(e))

print('Todos os emails foram enviados com sucesso!')
            
