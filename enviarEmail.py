import win32com.client as win32
import os
import pandas as pd

# Criando a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Ler os dados do CSV usando Pandas
df = pd.read_csv("benef.csv", dtype={'VENDAS': str}, sep=';')

for _, row in df.iterrows():
    destinatario = {
        "email": row["EMAIL"],
        "vendas": row["VENDAS"]  # Nome do arquivo da coluna "VENDAS"
    }

    # Criar e-mail
    email = outlook.CreateItem(0)

    # Configurar informações do e-mail
    email.To = destinatario["email"]
    email.Subject = "Teste de envio de e-mail automático"
    email.HTMLBody = """
        <p>Olá Wellington, tudo bem?</p>

        <p>Este aqui é um e-mail de teste.</p>

        <p>Por favor desconsiderar.</p>

        <p>Atenciosamente,</p>

        <p>Equipe de Desenvolvimento.</p>
    """

    # Construir o caminho completo para o arquivo PDF
    pdf_path = rf'C:\boletos\{destinatario["vendas"]}.pdf'

    # Verificar se o arquivo PDF existe
    if os.path.exists(pdf_path):
        # Anexar o arquivo PDF correspondente
        email.Attachments.Add(pdf_path)
    else:
        print(f"Arquivo PDF não encontrado para {destinatario['email']}")

    # Enviar o e-mail
    email.Send()
    print(f"E-mail enviado para {destinatario['email']}")

print("Todos os e-mails foram enviados!")
