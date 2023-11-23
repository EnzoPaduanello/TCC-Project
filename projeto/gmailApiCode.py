import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

#email = input("Digite o endereço de email para ser enviado")

# Autenticação
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
flow = InstalledAppFlow.from_client_secrets_file('c:\\Users\\ENZOVITORIANOPUTTINI\\Desktop\\Projeto_TCC\\key.json', SCOPES)
creds = flow.run_local_server(port=0)


# Construindo o serviço
service = build('gmail', 'v1', credentials=creds)


def enviar_email(subject, to, body):
    # Codificar a mensagem
    message = base64.urlsafe_b64encode(f"Subject: {subject}\nTo: {to}\n\n{body}".encode('utf-8')).decode('utf-8')


    # Enviar o e-mail
    send_message = service.users().messages().send(userId="me", body={'raw': message}).execute()
    print(f"Message sent to {to}, Message Id: {send_message['id']}")


# Testando
#enviar_email("Teste API Gmail", "enzo.paduanello@gmail.com", "Olá. Bom Dia!")