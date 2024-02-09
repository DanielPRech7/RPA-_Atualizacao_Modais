import pyautogui
import webbrowser
import pandas as pd
import time
import pyperclip
from pynput.mouse import Button, Controller
from selenium import webdriver
import re



url = 'https://chat.openai.com/'
webbrowser.open(url)
time.sleep(10)


caminho_arquivo = 'C:/Users/User/Downloads/meta_tags/meta_tags/meta_planilha.xlsx'

sript = "Seguindo a estrutura padrao de: nome do produto com palavras-chave para SEO, texto persuasivo, caracteristicas e call to action. - Exemplo: Conheca o guarda-roupa Paris, a escolha perfeita para o seu quarto, feito em MDF na cor marrom cinamomo, possui 6 portas e 4 gavetas, venha aproveitar!- . Crie uma breve descricao meta-tag para SEO que contenha ate 160 caracteres, seguindo esse padrao que foi passado, utilizando palavras com grande relevancia de busca para melhorar o SEO e ranqueamento do produto, sem utilizar informacoes de peso e medidas, usando como base o seguinte texto: "

df = pd.read_excel(caminho_arquivo)
max_caract = 2500

df['DescricaoProduto'] = df['DescricaoProduto'].apply(lambda x: x[:max_caract] if isinstance(x, str) and len(x) > max_caract else x)

try:
    contador = 0

    for index, row in df.iterrows():

        while True:
            if contador % 20 == 0 and contador != 0 and contador != 1: #espaço de respiro
                time.sleep(600) 
                break
            else:
                break

        valor_celula = row['DescricaoProduto']
        contador += 1
        copiar_celula = df.at[contador, 'DescricaoProduto']

        
        pyautogui.click(x=96, y=85)
        time.sleep(3)

        pyautogui.click(x=515, y=641)
        time.sleep(3)

        pyautogui.typewrite(sript + copiar_celula)
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(60)

        pyautogui.moveTo(x=733, y=450) #selecionar
        time.sleep(3)
        mouse = Controller()
        mouse.click(Button.left, 3)
        time.sleep(3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(3)

        texto = pyperclip.paste()
        tamanho_texto = len(re.findall(r'\w', texto))
        #print(tamanho_texto)

        while tamanho_texto>= 195 or tamanho_texto == None: #loop de cont caract

            pyautogui.hotkey('F5')
            time.sleep(7)

            pyautogui.click(x=52, y=176) #excluir chat
            time.sleep(2)

            pyautogui.moveTo() 
            pyautogui.click(x=236, y=182) #excluir chat
            time.sleep(3)

            pyautogui.click(x=210, y=181) #excluir chat
            time.sleep(7)

            pyautogui.click(x=96, y=85) #abrir chat
            time.sleep(3)

            pyautogui.click(x=515, y=641) #selecionar caixa do chat
            time.sleep(3)

            pyautogui.typewrite(sript + copiar_celula)
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(60)

            pyautogui.moveTo(x=733, y=450) #selecionar
            time.sleep(3)
            mouse = Controller()
            mouse.click(Button.left, 3)
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(3)


            texto = pyperclip.paste()
            tamanho_texto = len(re.findall(r'\w', texto))

        time.sleep(2)
        pyautogui.click(x=52, y=176) #excluir chat
        time.sleep(2)

        pyautogui.click(x=236, y=182) #excluir chat
        time.sleep(2)
               
        pyautogui.click(x=210, y=181) #excluir chat
        time.sleep(10)

        df.at[contador, 'meta_tags'] = texto


    df.to_excel(caminho_arquivo, index=False)
    pyautogui.click(x=1345, y=7) #fechar janela
    print("Código acabou")
    

except Exception as e:

    df.to_excel(caminho_arquivo, index=False)
    print(f"O Código parou na linha: {contador}")
    pyautogui.click(x=1345, y=7) #fechar janela
    print("Código acabou")
    

except KeyboardInterrupt:
    df.to_excel(caminho_arquivo, index=False)
    print(f"O Código parou na linha: {contador}")
    pyautogui.click(x=1345, y=7) #fechar janela
    print("Código acabou")
    
    