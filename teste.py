import imaplib
import email
import openpyxl

# Configurações do servidor IMAP e informações da conta de e-mail
IMAP_SERVER = 'imap-mail.outlook.com'
EMAIL = 'marcos.ribeiro_@hotmail.com'
PASSWORD = 'M@r131917'

# Conectar-se ao servidor IMAP e selecionar a caixa de entrada
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)
mail.select('inbox')

# Pesquisar por e-mails
result, data = mail.search(None, 'ALL')
print(result)
# Lista para armazenar os corpos dos e-mails
email_bodies = []

# Iterar pelos IDs dos e-mails
for num in data[0].split():
    result, msg_data = mail.fetch(num, '(RFC822)')
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Extrair o corpo do e-mail
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                email_bodies.append(part.get_payload(decode=True).decode())
    else:
        email_bodies.append(msg.get_payload(decode=True).decode())

# Criar um arquivo do Excel
wb = openpyxl.Workbook()
ws = wb.active

# Preencher o Excel com os corpos dos e-mails
for i, body in enumerate(email_bodies, start=1):
    ws.cell(row=i, column=1, value=body)

# Salvar o arquivo do Excel
wb.save('emails.xlsx')

# Fechar a conexão com o servidor IMAP
mail.logout()
