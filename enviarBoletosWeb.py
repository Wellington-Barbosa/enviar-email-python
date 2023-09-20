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

# Variável para rastrear se ocorreu algum erro
erro_ocorreu = False

# Criar um arquivo CSV para rastrear boletos enviados
registro_boletos_enviados = 'boletos_enviados.csv'

# Verificar se o arquivo de registro já existe
if not os.path.exists(registro_boletos_enviados):
    # Se não existir, crie o arquivo com cabeçalhos
    with open(registro_boletos_enviados, 'w') as f:
        f.write('EMAIL,VENDAS\n')

for _, row in df.iterrows():
    destinatario_email = row["EMAIL"].strip()  # Remove espaços em branco
    vendas = str(row["VENDAS"]).zfill(10).strip()  # Remove espaços em branco e adiciona zeros à esquerda

    # Verificar se o boleto já foi enviado para este destinatário
    with open(registro_boletos_enviados, 'r') as f:
        registros_enviados = f.readlines()
        if f'{destinatario_email},{vendas}\n' in registros_enviados:
            data_hora_1 = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            ja_enviado = f"{data_hora_1} - Boleto já enviado para {destinatario_email}, VENDAS: {vendas}"
            print(ja_enviado)

            # Salvar a mensagem de erro em um arquivo de log
            log_path = r'C:\beneficiarios\log.txt'
            with open(log_path, 'a') as log_file:
                log_file.write(ja_enviado + '\n')

            continue  # Pule este boleto, pois já foi enviado

    # Carregue o corpo do e-mail a partir de um arquivo
    email_body_file = 'email_body.html'
    if os.path.exists(email_body_file):
        with open(email_body_file, 'r', encoding='utf-8') as f:  # Especifique a codificação UTF-8
            email_body = f.read()
    else:
        raise Exception("O arquivo do corpo do e-mail não foi encontrado.")

    # Construir o caminho completo para o arquivo PDF
    pdf_path = rf'C:\boletos\{vendas}.pdf'

    # Verificar se o arquivo PDF existe
    if os.path.exists(pdf_path):
        # Configurar o e-mail
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = destinatario_email
        msg['Subject'] = "Boleto Mensal - UnimedRV"

        # Anexar o corpo do e-mail em formato HTML com codificação UTF-8
        msg.attach(MIMEText(email_body, 'html', 'utf-8'))

        # Anexar o arquivo PDF correspondente
        with open(pdf_path, "rb") as pdf_file:
            pdf_data = pdf_file.read()
            pdf_attachment = MIMEApplication(pdf_data, Name=os.path.basename(pdf_path))
            msg.attach(pdf_attachment)

        # Configurar o servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)

        # Enviar o e-mail
        server.sendmail(smtp_username, destinatario_email, msg.as_string())
        server.quit()

        # Após o envio bem-sucedido do boleto, registre o envio
        with open(registro_boletos_enviados, 'a') as f:
            f.write(f'{destinatario_email},{vendas}\n')

        data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        mensagem = f"{data_hora} - E-mail enviado para {destinatario_email}"
        print(mensagem)

        # Salvar a mensagem em um arquivo de log
        log_path = r'C:\beneficiarios\log.txt'
        with open(log_path, 'a') as log_file:
            log_file.write(mensagem + '\n')
    else:
        erro_msg = f"Arquivo PDF nao encontrado para {destinatario_email}"
        print(erro_msg)

        # Salvar a mensagem de erro em um arquivo de log
        log_path = r'C:\beneficiarios\log.txt'
        with open(log_path, 'a') as log_file:
            log_file.write(erro_msg + '\n')

        # Marcar que ocorreu um erro
        erro_ocorreu = True

# Imprimir a mensagem final apenas se nenhum erro ocorreu
if not erro_ocorreu:
    print("Todos os e-mails foram enviados!")
