import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(5)

workbook = openpyxl.load_workbook('colaboradores.xlsx')
pagina_clienetes = workbook ['Plan1']

for linha in pagina_clienetes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value

        mensagem = f'Teste'

        try:
                link_mensagem_whatapp = F'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                webbrowser.open(link_mensagem_whatapp)
                sleep(10)
                pyautogui.hotkey('Enter')
                sleep(3)
                pyautogui.hotkey('ctrl','w')
                sleep(5)
        except:
                print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}')