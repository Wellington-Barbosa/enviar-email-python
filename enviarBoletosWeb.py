import os
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Configurações do servidor SMTP do Outlook
smtp_server = 'smtp.office365.com'
smtp_port = 587
smtp_username = 'naoresponda.unimed@unimedrv.com.br'
smtp_password = 'Nr@01020304'

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

    # Criar o corpo do e-mail
    mensagem = """
        <p>Olá,</p> 
        <p>Espero encontrá-lo(a)</p>
        <p>Segue em anexo o boleto simples do contrato de plano de saúde da Unimed Rio Verde.</p>
        <p><b>Esse e-mail não recebe retorno, por gentileza, não responder.</b></p>
        <p>Atenciosamente,</p>
        <p>Departamento Financeiro.</p>
    """

    # Configurar o e-mail
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = destinatario_email
    msg['Subject'] = "Boleto Mensal - UnimedRV"

    # Anexar o corpo do e-mail em formato HTML
    msg.attach(MIMEText(mensagem, 'html'))

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