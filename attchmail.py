import flet as ft
import win32com.client as win32
from openpyxl import Workbook
from pathlib import Path
import os 
import re
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
    
    


def baixar_anexos(email, destino):
    attachments = email.Attachments
    num_arquivos_baixados = 0
    arquivos_baixados = []  # Lista para armazenar os nomes dos arquivos baixados
    
    total_anexos = len(attachments)
    
    # Verifica se há mais de um arquivo anexado
    if len(attachments) > 1:
        # Cria uma pasta dentro do diretório de destino para armazenar os arquivos
        # Renomeando a pasta com base no assunto do e-mail
        subject_folder_name = re.sub(r'[;,:@/]', '_', email.Subject)  # Remover caracteres especiais
        sub_destino = destino / subject_folder_name.strip()  # Pasta com letras e números apenas
        sub_destino.mkdir(parents=True, exist_ok=True)  # Criar diretório de destino
        
        # Salva todos os arquivos anexados dentro da pasta
        for index, attachment in enumerate(attachments):
            try:
                save_path = os.path.join(sub_destino, attachment.FileName)
                print("Salvando anexo em:", save_path)
                attachment.SaveAsFile(save_path)
                num_arquivos_baixados += 1
                arquivos_baixados.append(attachment.FileName)
            except Exception as e:
                print(f"Erro ao baixar anexo para '{save_path}': {e}")
                try:
                    # Se o erro ocorrer, tenta salvar diretamente na pasta de destino principal
                    save_path = os.path.join(destino, attachment.FileName)
                    print("Tentando salvar anexo diretamente na pasta de destino:", save_path)
                    attachment.SaveAsFile(save_path)
                    num_arquivos_baixados += 1
                    arquivos_baixados.append(attachment.FileName)
                except Exception as e:
                    print(f"Erro ao baixar anexo diretamente na pasta de destino: {e}")
                    arquivos_baixados.append(f"Erro: {e}")
                    
            
    else:
        # Caso haja apenas um arquivo anexado, baixa normalmente
        for attachment in attachments:
            if attachment.FileName.endswith(('.zip', '.rar', '.doc', '.docx')):
                try:
                    save_path = os.path.join(destino, attachment.FileName)
                    print("Salvando anexo em:", save_path)
                    attachment.SaveAsFile(save_path)
                    num_arquivos_baixados += 1
                    arquivos_baixados.append(attachment.FileName)
                except Exception as e:
                    print(f"Erro ao baixar anexo: {e}")
                    arquivos_baixados.append(f"Erro: {e}")
    return num_arquivos_baixados, arquivos_baixados


def main(page: ft.Page):
    page.title = "Baixar Anexos de E-mails"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.fonts = {
        "Kanit": "https://raw.githubusercontent.com/google/fonts/master/ofl/kanit/Kanit-Bold.ttf"
    }
    
    
    
   

    
    page.add(ft.Container(
            content= ft.Column([
            ft.Row([ft.Icon(ft.icons.ATTACH_EMAIL, size=50),
        ft.Text("AttachSavePro", size= 60 , font_family= "Kanit")
        ],
        ft.MainAxisAlignment.CENTER),]), padding=10))

    def selecionar_destino(e):
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela principal

        selected_directory = filedialog.askdirectory()
        if selected_directory:
            destino = Path(selected_directory)
            txt_destino.value = selected_directory
            



    # Campos de entrada
    txt_data_inicial = ft.TextField(value="", label="Data Inicial (DD-MM-YYYY)", width=250)
    txt_data_final = ft.TextField(value="", label="Data Final (DD-MM-YYYY)", width=250)
    txt_destino = ft.TextField(value="", disabled=True)  # Desativado por padrão
    txt_assunto = ft.TextField(value="", label="Assunto do E-mail", width=250)
    
    # Botão para selecionar o diretório de destino    
    btn_selecionar_destino = ft.ElevatedButton(text="Selecionar Pasta para Salvar", on_click=selecionar_destino, width=250)
    
    def exibir_alerta_conclusao():
        # Cria uma janela de pop-up
        alerta = tk.Tk()
        alerta.title("Concluído")

        # Adiciona um rótulo com a mensagem de conclusão
        lbl_alerta = tk.Label(alerta, text="Execução concluída com sucesso!")
        lbl_alerta.pack(padx=20, pady=10)

        # Mantém a janela de pop-up aberta até ser fechada pelo usuário
        alerta.mainloop()

    def executar_click(e):
        # Coletar os valores dos campos de entrada
        data_inicial = datetime.strptime(txt_data_inicial.value, "%d-%m-%Y")
        data_final = datetime.strptime(txt_data_final.value, "%d-%m-%Y")
        destino = Path(txt_destino.value)
        assunto = txt_assunto.value
        
        

        # Cria a pasta de destino se não existir
        destino.mkdir(parents=True, exist_ok=True)
        
        

        # Conectar-se ao Outlook
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 corresponde à pasta "Caixa de Entrada"

        # Lista para armazenar os detalhes dos e-mails
        emails_details = []

        # Iterar sobre os e-mails na caixa de entrada
        for message in inbox.Items:
            if assunto in message.Subject and data_inicial.date() <= message.ReceivedTime.date() <= data_final.date():
                subject = message.Subject
                received_date = message.ReceivedTime.date()
                received_time = message.ReceivedTime.time()
                
                # Baixar os anexos e contar a quantidade baixada
                num_arquivos_baixados, arquivos_baixados = baixar_anexos(message, destino)
                
                if num_arquivos_baixados > 0:
                    status_download = "Baixado"
                else:
                    status_download = "Nenhum anexo baixado"
                
                # Adiciona informações sobre os arquivos baixados à lista
                email_detail = [subject, received_date, received_time, status_download, num_arquivos_baixados, arquivos_baixados]
                emails_details.append(email_detail)

        # Criar planilha Excel para armazenar os detalhes dos e-mails
        wb = Workbook()
        ws = wb.active
        ws.append(["Assunto", "Data de Recebimento", "Hora de Recebimento", "Status", "Quantidade de Arquivos Baixados", "Arquivos Baixados"])

        # Adicionar detalhes dos e-mails à planilha
        for email_detail in emails_details:
            # Concatenar os nomes dos arquivos em uma única string
            arquivos_baixados_str = ', '.join(email_detail[5])
            # Substituir o valor da lista de nomes de arquivo pela string concatenada
            email_detail[5] = arquivos_baixados_str
            # Adicionar detalhes do e-mail à planilha
            ws.append(email_detail)

        # Salvar a planilha Excel
        excel_file_path = os.path.join(destino, "emails.xlsx")
        wb.save(excel_file_path)
        


       

        print(f"Planilha Excel criada com sucesso em '{excel_file_path}'.")
        
         # Exibir alerta de conclusão
        exibir_alerta_conclusao()
        


    # Botões

    btn_executar = ft.ElevatedButton(text="Executar", on_click=executar_click, width=250)
    
    
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    # Adicionar os componentes à página
   
    
    
    page.add(
        ft.Column(
            [
                txt_data_inicial,
                txt_data_final,
                btn_selecionar_destino,  # Adicionando apenas o botão
                txt_assunto,
                btn_executar,
                
            
            ],
            alignment=ft.MainAxisAlignment.CENTER,
        )
    )

ft.app(main)
