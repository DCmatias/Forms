import pyautogui
import pandas as pd
import os.path
import gspread
import time
import schedule 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, NoSuchWindowException
from pyautogui import write
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow



SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDS_FILE = "baixar_base.json"

def autenticar():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def importar_dados_google_sheets(creds, id_planilha, nome_guia):
    cliente = gspread.authorize(creds)
    planilha = cliente.open_by_key(id_planilha)
    dados_guia = planilha.worksheet(nome_guia).get_all_values()
    dados_df = pd.DataFrame(dados_guia[1:], columns=dados_guia[0])
   
    return dados_df

# Substitua 'ID_DA_PLANILHA' pelo ID da sua planilha
id_planilha = '1unLoVoV4YUJl7zjNHkIL6jTUIFm2qc4-LwkKJGi4uhI'
# Substitua 'NomeDaSuaGuia' pelo nome da sua guia
nome_guia = 'PACOTES'
#nome_guia_2 = 'PCTS_LANCADOS'

# Autentique-se
creds = autenticar()

# Importe os dados da planilha
dados_df = importar_dados_google_sheets(creds, id_planilha, nome_guia)

# Filtro com duas condições usando o operador lógico AND (&)
dados_filtrados = dados_df[dados_df['status'] != 'LANÇADO'].copy()

 # Convertendo a coluna 'protocolo' para números inteiros usando loc
dados_filtrados.loc[:, 'protocolo'] = dados_filtrados['protocolo'].astype(int)


navegador = webdriver.Chrome()
navegador.get('https://docs.google.com/forms/d/e/1FAIpQLScUDXW1BenOvMX3vlzdzfR35WYeanuq9Ej1kGeTWAsOlrtI2g/viewform')

dados_atualizados_lista = []


def main():


    for i, protocolo in enumerate(dados_filtrados['protocolo']):
        
        try:
      
            protocolo = dados_filtrados['protocolo'].iloc[i]
            facility = dados_filtrados['facility'].iloc[i]
            pacotes = dados_filtrados['pacotes'].iloc[i]
            posicoes = dados_filtrados['posicoes'].iloc[i]
            pallets_madeira = dados_filtrados['pallets_madeira'].iloc[i]
            gayloards = dados_filtrados['gayloards'].iloc[i]
            manga_ = dados_filtrados['manga_'].iloc[i]
            data = dados_filtrados['data'].iloc[i]
            hora = dados_filtrados['hora'].iloc[i]
            min = dados_filtrados['min'].iloc[i]
            status = dados_filtrados['status'].iloc[i]

        

            # protocolo
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(int(protocolo))

            # facility

            click = navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[1]/div[1]').click()

            # Substitua pelo nome da opção desejada
            opcao = int(facility)

            pyautogui.sleep(2)
            # Coloque isto onde o programa deveria selecionar uma opção
            for _ in range(opcao):
                write(['down'])
            write(['enter'])

            # tipo
            pyautogui.sleep(2)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div/span/div/div[2]/label').click()

            # proximo
            pyautogui.sleep(2)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()

            #pacotes
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(pacotes)

            # posicoes pallets
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(posicoes)

            # pallets madeira
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(pallets_madeira)

            # gayloards
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(gayloards)

            # manga pallets
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(manga_)

            # proximo
            pyautogui.sleep(2)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div[2]/span').click()

            # data
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(data)

            #hora
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(hora)
            
            #minuto
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/div/div[3]/div/div[1]/div/div[1]/input').send_keys(min)

            # enviar
            pyautogui.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div[2]/span').click()


            # Após preencher o formulário
            dados_filtrados['status'].iloc[i] = 'LANÇADO'
            
            # Salvar dados na aba "PCTS_LANCADOS"
            aba_destino = 'PCTS_LANCADOS'
            cliente = gspread.authorize(creds)
            planilha = cliente.open_by_key(id_planilha)
            guia_destino = planilha.worksheet(aba_destino)
            guia_destino.append_row(dados_filtrados.iloc[i].tolist(), value_input_option='RAW')

            # Clicar no botão "Enviar outra resposta"
            navegador.find_element(By.LINK_TEXT, 'Enviar outra resposta').click()

            # Aguardar alguns segundos para a página recarregar antes de iniciar a próxima iteração
            pyautogui.sleep(1)  # Ajuste o tempo conforme necessário


        except Exception as e:
            navegador = webdriver.Chrome()
            navegador.get('https://docs.google.com/forms/d/e/1FAIpQLScUDXW1BenOvMX3vlzdzfR35WYeanuq9Ej1kGeTWAsOlrtI2g/viewform')

            continue  # Reinicia o loop para a próxima iteração


# Agendar a execução da função main a cada 5 minutos
schedule.every(15).minutes.do(main)

# Agendar a execução do código diariamente às 18:00
#schedule.every().day.at("18:00").do(main)

while True:
    schedule.run_pending()
    time.sleep(1)