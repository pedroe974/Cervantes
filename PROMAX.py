
import subprocess
import time
from datetime import datetime
import schedule

import pyautogui as auto
import pygetwindow as gw

nome_arquivo = r"C:\Users\pedro.alves\Desktop\DADOS.txt"
with open(nome_arquivo, 'r') as file:
    linhas = file.readlines()  # Lê todas as linhas do arquivo
    for linha in linhas:
        if "Login:" in linha:  # Verifica se a linha contém "Login:"
            usuario = linha.split("Login:")[1].strip()  # Extrai o usuário
        if "Senha:" in linha:  # Verifica se a linha contém "Senha:"
            senha = linha.split("Senha:")[1].strip()  # Extrai a senha

#CHAMAR O EDGE
def abri_promax():
    # URL do site que você deseja abrir
    url = "https://cervantes.promaxcloud.com.br/pw/"
    navegador = "C:\\Program Files\\Internet Explorer\\iexplore.exe"

    while True:
        # Abre o navegador
        processo = subprocess.Popen([navegador, url])
        print("Abrindo o navegador...")
        time.sleep(10)

        # Maximiza a janela
        auto.hotkey("alt", "space")
        time.sleep(2)
        auto.press("x")
        time.sleep(10)

        try:
            # Tenta localizar a imagem na tela
            localizacao = auto.locateOnScreen('PROMAX.png', confidence=0.6)
            if localizacao:
                print("Imagem encontrada! Prosseguindo com o processo.")
                break  # Sai do loop se a imagem for encontrada
            else:
                raise Exception("Imagem não encontrada.")
        except auto.ImageNotFoundException:
            print("Imagem não encontrada. Fechando o navegador e tentando novamente.")
        except Exception as e:
            print(f"Erro ao localizar imagem: {e}")

        # Encerra o navegador e reinicia o processo
        processo.terminate()
        time.sleep(5)

    print("Processo concluído com sucesso!")


def inserir_dados(time_padrao):
    # Selecionar o local
    auto.sleep(time_padrao)
    auto.click(1118, 378)
    # Fechar o Edge
    auto.sleep(time_padrao)
    auto.click(1888, 91)
    # Campo de Login
    auto.sleep(time_padrao*2)
    auto.tripleClick(1888, 89)
    auto.sleep(time_padrao)
    auto.tripleClick(965, 256)
    auto.sleep(time_padrao)
    auto.press('delete')
    auto.sleep(time_padrao)
    auto.write(usuario)
    auto.sleep(time_padrao)
    # Campo de Senha
    auto.tripleClick(963, 282)
    auto.sleep(time_padrao)
    auto.press('delete')
    auto.sleep(time_padrao)
    auto.write(senha)
    auto.sleep(time_padrao)
    # Botão confirmar
    auto.tripleClick(960, 318)
    auto.sleep(time_padrao)
def definir_unidade(time_padrao,sede):
    # Selecionar o campo da unidade
    auto.rightClick(980, 282)

    if sede==1:
        # Matriz igual a i
        auto.sleep(time_padrao)
        auto.press('up')
        auto.sleep(time_padrao)
        auto.press('enter')
    else:
        # Filial igual a 2
        auto.sleep(1)
        auto.press('down')
        auto.sleep(1)
        auto.press('enter')
    # Botão confirma
    auto.tripleClick(960, 318)

def seguir(time_padrao):
    # Passar dos pop_up
    auto.sleep(time_padrao*2)
    auto.press('enter')
    auto.sleep(time_padrao*2)
    auto.press('enter')
    auto.sleep(time_padrao*2)
    auto.press('enter')


def janela_certa(titulo_janela):
    # Lista todas as janelas abertas com seus nomes (títulos)
    windows = gw.getAllTitles()

    # Busca a janela pelo título
    janela = gw.getWindowsWithTitle(titulo_janela)

    # Verifica se a janela foi encontrada
    if janela:
        # Ativa a janela, colocando-a em primeiro plano
        janela[0].activate()
        print(f"Janela '{titulo_janela}' selecionada e trazida para o primeiro plano.")
    else:
        print(f"Janela com título '{titulo_janela}' não encontrada.")




windows = gw.getAllTitles()
def rotina():
    for window in windows:
        print(window)


"""while True:

    # Pegar o horário atual
    agora = datetime.now()
    # Verificar se o horário é 7:00 ou mais tarde
    if agora.hour > 17 or (agora.hour == 16 and agora.minute >= 53):
        rotina()
        break
    else:
        # Esperar por 1 minuto antes de verificar novamente
        print("Aguardando até 7:00 da manhã...")
        time.sleep(60)"""

# Caminho da imagem que você deseja validar




