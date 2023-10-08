from pathlib import Path
import win32com.client  #pip install pywin32
import datetime
import re

# Cria a pasta
destino = Path.cwd() / "Arquivos do email"
destino.mkdir(parents=True, exist_ok=True)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

root_folders = outlook.Folders.Item(1)
inbox = root_folders.Folders['pasta_outlook'] # acessa a pasta criada no outlook para pegar os arquivos 
messages = inbox.Items

for message in messages:
    subject = message.Subject
    body = message.body
    attachments = message.Attachments
    received_date = message.ReceivedTime
    received_date = received_date.strftime("%Y-%m-%d")

    # Criar pastas separadas para cada mensagem, excluindo caracteres especiais e demais
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    topico = received_date + subject
    target_folder = destino / re.sub('[^0-9a-zA-Z]+', '', subject) / topico 
    target_folder.mkdir(parents=True, exist_ok=True)

    # Escreve a mensagem contida no email
    Path(target_folder / f"Arquivo do dia {received_date}.txt").write_text(str(body)+str(received_date))

    # Salva os anexos e trata caract. especiais
    for attachment in attachments:
        filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
        attachment.SaveAsFile(target_folder / filename)