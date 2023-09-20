# Documentação do Script Python - Envio de Boletos por E-mail

Este é um script Python que automatiza o processo de envio de boletos por e-mail para beneficiários. O script faz uso de bibliotecas como `os`, `json`, `pandas`, `datetime`, `smtplib` e módulos relacionados a e-mails para realizar o envio. Abaixo, você encontrará uma documentação detalhada explicando o funcionamento do script e como configurá-lo.

## Funcionalidade

O script realiza as seguintes tarefas:

1. Carrega as configurações do servidor SMTP do Outlook a partir de um arquivo JSON.
2. Lê um arquivo CSV contendo informações dos beneficiários.
3. Salva o conteúdo do DataFrame em um novo arquivo CSV com um separador diferente.
4. Verifica se boletos já foram enviados para cada beneficiário.
5. Envia boletos por e-mail para destinatários que ainda não receberam.
6. Registra os envios bem-sucedidos em um arquivo CSV e mantém um registro de logs.

## Pré-requisitos

Antes de executar o script, é importante garantir que os seguintes pré-requisitos sejam atendidos:

1. Um servidor SMTP do Outlook com informações de configuração (endereço, porta, nome de usuário e senha) deve ser configurado e as informações devem ser fornecidas em um arquivo JSON chamado `config.json`.
2. O arquivo CSV contendo informações dos beneficiários deve estar disponível no caminho especificado em `caminho_do_arquivo`.
3. Um arquivo HTML chamado `email_body.html` deve ser criado para o corpo do e-mail.
4. Os boletos a serem enviados devem estar no formato PDF e localizados em um diretório específico (`C:\boletos`).

## Configuração

O arquivo `config.json` deve ter o seguinte formato:

```json
{
    "smtp_server": "seu_servidor_smtp",
    "smtp_port": porta_do_servidor_smtp,
    "smtp_username": "seu_nome_de_usuario",
    "smtp_password": "sua_senha"
}
```

## Uso

Para usar o script, siga estas etapas:

1. Garanta que todos os pré-requisitos sejam atendidos.
2. Configure o arquivo `config.json` com as informações corretas do servidor SMTP.
3. Crie o arquivo `email_body.html` com o conteúdo desejado para o corpo do e-mail.
4. Certifique-se de que os boletos a serem enviados estejam no diretório `C:\boletos`.
5. Execute o script Python.

## Resultados

O script produz os seguintes resultados:

- Envia boletos por e-mail para destinatários que ainda não receberam.
- Mantém um registro dos boletos enviados em um arquivo CSV chamado `boletos_enviados.csv` no diretório `Histórico de Envios`.
- Registra todas as atividades e erros em um arquivo de log chamado `log.txt` no diretório `Log's de Envio`.

## Considerações Finais

Este script foi desenvolvido para automatizar o processo de envio de boletos por e-mail para beneficiários, proporcionando uma maneira eficiente e organizada de gerenciar o envio e manter um registro das operações realizadas. Certifique-se de configurar corretamente as informações de configuração do servidor SMTP e os caminhos dos arquivos antes de executar o script.