from playwright.sync_api import sync_playwright

#módulo de geração de log
import logging

#módulo para arredondar numero
import math

import pyautogui

#módulo de tempo de espera
import time

import cred

#**************************** Inicio ****************************

#iniciando o navegador
with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False)  #Headless = não fazer no modo invisvel.
    pagina = navegador.new_page()

    #configs finais do navegador
    pagina.goto("https://empresa.holmesdoc.com/#home")

    pagina.locator('//*[@id="tiUser"]').fill(cred.email)
    pagina.locator('//*[@id="tiPass"]').fill(cred.senha)
    pagina.locator('//*[@id="login"]/div[4]/button').click()

    #  log config básico. W é de Write, A de append. W ele apaga o ultimo log, A, ele vai somando
    logging.basicConfig(filename='Excluir.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')

    #  entrar na página de erros
    pagina.goto("https://empresa.holmesdoc.com/#search/+_nature%253A%2522Incompleto%2522%2520+_uploaddate%253A%255B20161128%2520TO%2520*%255D/p1/sortedby/_uploaddate/DESC/GRID")


    #  Calcular quantidade de páginas
    qtd = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
    qtd0 = qtd.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
    qtd1 = int(qtd0)
    qtd2 = math.ceil(qtd1/54)  #arredondar para cima
    logging.warning('************ - Total de páginas a serem excluidas: ' + str(qtd2) + " - ************")

    i = qtd2

    #iniciar a repetição, de acordo com o tamanho da tabela (usando enumerate)
    while i >= 1:

        procurar = "sim"
        procurar1 = "sim"
        #  trocar entre pegina 1 e 2 pra agilizar a exclusão. Esta é uma forma de ultrapassar a limitação do próprio sistema
        try:
            pagina.locator('//*[@id="1"]').click()
            time.sleep(4)

            #identificando onde ta a informação no excel, e o i, indica a linha. Conversão em String para evitar erro
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[7]/span[1]').click()
            time.sleep(1)
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[6]/span[1]').click()
            time.sleep(2)
            
            # inicio do reconhecimento de imagem
            while procurar == "sim":
                try:
                    img = pyautogui.locateCenterOnScreen('Extras\Sim.png', confidence=0.3)
                    pyautogui.click(img.x, img.y)
                    procurar = "não"
                    pyautogui.moveRel(50, 50)
                except:
                  # continuar tentando até dar certo
                    time.sleep(1)
                    print("Não localizado")

            #salvando no log em texto, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Total de ' + str(i) + " paginas restantes!")

            time.sleep(1)

            i = i-1

        except:
            #salvando erro no log, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Erro ao excluir a Página ' + str(i) + " !")



        #  trocar entre página 1 e 2, pra agilizar a exclusão
        try:
            pagina.locator('//*[@id="2"]').click()
            time.sleep(4)

            #identificando onde ta a informação no excel, e o i, indica a linha. Conversão em String para evitar erro
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[7]/span[1]').click()
            time.sleep(1)
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[6]/span[1]').click()
            time.sleep(2)
            while procurar1 == "sim":
                try:
                    img = pyautogui.locateCenterOnScreen('Extras\sim.png', confidence=0.9)
                    pyautogui.click(img.x, img.y)
                    procurar1 = "não"
                    pyautogui.moveRel(50, 50)
                except:
                    time.sleep(1)
                    print("Não localizado")

            #salvando no log em texto, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Total de ' + str(i) + " paginas restantes!")

            time.sleep(1)

            i = i - 1
        except:
            #salvando erro no log, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Erro ao excluir a Página ' + str(i) + " !")

    #salvando no log,informando que finalizou o bot
    logging.warning('Bot finalizado')
