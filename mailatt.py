import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# Acessando caixa de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 corresponde à pasta de entrada 
# Inicialize uma variável de contagem
count = 1
# Itere pelos itens da caixa de entrada
for item in inbox.Items:
    if item.Subject == "Conheça nossos novos colegas": # titulo buscado
        for attachment in item.Attachments:
            # Gere um nome de arquivo sequencial
            new_filename = f"C:/Users/lgonzales/Desktop/Script Python/mail/image{count}"
            # Adicione uma extensão ao nome de arquivo com base no tipo do anexo
            file_extension = attachment.FileName.split('.')[-1]
            new_filename = f"{new_filename}.{file_extension}"
            # Salve o anexo com o novo nome de arquivo
            attachment.SaveAsFile(new_filename)
            print(f"Anexo '{attachment.FileName}' baixado como '{new_filename}'")           
            count += 1
