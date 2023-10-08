from imap_tools import MailBox, AND
import os
from dotenv import load_dotenv

load_dotenv()
senha = os.getenv("SENHA")
usuario = os.getenv("USUARIO")

# https://www.systoolsgroup.com/imap/
meu_email = MailBox('outlook.office365.com').login(usuario, senha)

# pegar email envaidos por um remetente especifico
#https://github.com/ikvk/imap_tools#seach-criteria

lista_emails = meu_email.fetch(AND(from_="enviado_por_tal@email.com", to = "outro_usuario@email.com"))
for email in lista_emails:
    print(email.subject)
    print(email.text)


# pegar anexo de um email
lista_emails = meu_email.fetch(AND(from_="enviado_por_tal@email.com"))
for email in lista_emails:
    if len(email.attachments) > 0:
        for anexo in email.attachments:
            if "nome_arquivo" in anexo.filename:  # nao precisa da extensao final
                informacoes_anexo = anexo.payload
                with open("nome_arquivo.txt","wb") as arquivo_txt: # consegue usar extensao .xlsx, apenas substitua o que esta em .txt
                    arquivo_txt.write(informacoes_anexo)