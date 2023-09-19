import os
import json
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Carregue as configurações do arquivo JSON
config_file = 'config.json'
if os.path.exists(config_file):
    with open(config_file, 'r') as f:
        config = json.load(f)
else:
    raise Exception("O arquivo de configurações não foi encontrado.")

# Configurações do servidor SMTP do Outlook (Carregadas do arquivo .json)
smtp_server = config.get('smtp_server')
smtp_port = config.get('smtp_port')
smtp_username = config.get('smtp_username')
smtp_password = config.get('smtp_password')

# Carregue o arquivo CSV com ";" como separador
caminho_do_arquivo = r'C:\beneficiarios\benef.csv'
dataframe = pd.read_csv(caminho_do_arquivo, sep=';')

# Salve o DataFrame de volta como um arquivo CSV com "," como separador
novo_caminho_do_arquivo = r'C:\beneficiarios\novo_arquivo.csv'
dataframe.to_csv(novo_caminho_do_arquivo, sep=',', index=False)

# Ler os dados do CSV usando Pandas
df = pd.read_csv(r"C:\beneficiarios\novo_arquivo.csv", dtype={'VENDAS': str})

df.columns = df.columns.str.strip()

for _, row in df.iterrows():
    destinatario_email = row["EMAIL"]
    vendas = str(row["VENDAS"]).zfill(10)  # Nome do arquivo da coluna "VENDAS" com 10 dígitos

    # Carregue o corpo do e-mail a partir de um arquivo
    email_body_file = 'email_body.html'
    if os.path.exists(email_body_file):
        with open(email_body_file, 'r') as f:
            email_body = f.read()
    else:
        raise Exception("O arquivo do corpo do e-mail não foi encontrado.")

    # Configurar o e-mail
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = destinatario_email
    msg['Subject'] = "Boleto Mensal - UnimedRV"

    # Anexar o corpo do e-mail em formato HTML
    msg.attach(MIMEText(email_body, 'html'))

    # Construir o caminho completo para o arquivo PDF
    pdf_path = rf'C:\boletos\{vendas}.pdf'

    # Verificar se o arquivo PDF existe
    if os.path.exists(pdf_path):
        # Anexar o arquivo PDF correspondente
        with open(pdf_path, "rb") as pdf_file:
            pdf_data = pdf_file.read()
            pdf_attachment = MIMEApplication(pdf_data, Name=os.path.basename(pdf_path))
            msg.attach(pdf_attachment)
    else:
        print(f"Arquivo PDF não encontrado para {destinatario_email}")
        continue

    # Configurar o servidor SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_username, smtp_password)

    # Enviar o e-mail
    server.sendmail(smtp_username, destinatario_email, msg.as_string())
    server.quit()

    data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mensagem = f"{data_hora} - E-mail enviado para {destinatario_email}"
    print(mensagem)

    # Salvar a mensagem em um arquivo de log
    log_path = r'C:\beneficiarios\log.txt'
    with open(log_path, 'a') as log_file:
        log_file.write(mensagem + '\n')

print("Todos os e-mails foram enviados!")