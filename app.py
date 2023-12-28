import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

listOfClientsInXml = openpyxl.load_workbook('py-test.xlsx')
main_page_xml = 'clientes'
page_client = listOfClientsInXml[main_page_xml]

for client in page_client.iter_rows(min_row=2):
    client_name = client[0].value
    client_phone = client[1].value
    client_expiration = client[2].value
    
    message_client = f'Olá {client_name}, seu boleto vencerá no dia {client_expiration.strftime('%d/%m/%Y')}'
    
    
    try:
        redirect_to_whatsapp = f'https://web.whatsapp.com/send?phone={client_phone}&text={quote(message_client)}'
        
        webbrowser.open(redirect_to_whatsapp)
        
        sleep(10)
    
        arrow_send = pyautogui.locateCenterOnScreen('arrow-whatsapp.png')
        
        sleep(5)
        
        pyautogui.click(arrow_send[0], arrow_send[1])
        
        sleep(5)
        
        pyautogui.hotkey('ctrl', 'w')
        
        sleep(5)
    except:
        print(f'Não foi possível mandar mensagem para {client_name}')
        with open('errors.csv', 'a', newline='', encoding='utf-8') as file:
            file.write(f'{client_name},{client_phone}')