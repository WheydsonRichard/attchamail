from pathlib import Path
import datetime
import re

import win32com.client  #pip install pywin32


# Criar um pasta chamada Output
output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Conectar ao outlook 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Conectar à pasta
#inbox = outlook.Folders("youremail@provider.com").Folders("Inbox")
inbox = outlook.GetDefaultFolder(6)
# https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# DeletedItems=3, Outbox=4, SentMail=5, Inbox=6, Drafts=16, FolderJunk=23

# Receber mensagens
messages = inbox.Items

for message in messages:
    subject = message.Subject
    body = message.body
    attachments = message.Attachments

    # Crie uma pasta separada para cada mensagem, exclua caracteres especiais e carimbo de data/hora
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    target_folder = output_dir / re.sub('[^0-9a-zA-Z]+', '', subject) / current_time
    target_folder.mkdir(parents=True, exist_ok=True)

    # Gravar corpo em arquivo de texto
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    # Salvar anexos e excluir especiais
    for attachment in attachments:
        filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
        attachment.SaveAsFile(target_folder / filename)