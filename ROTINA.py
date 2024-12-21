#ROTINA
import pyautogui as auto
import DATA as dt
import PROMAX as pr
import pyperclip
import win32gui
import win32api
import win32con
import os
import time

def salvar_csv(caminho,tempo_base,nome):
    #Novo Padrão
    tempo_base=3
    # Selecionar CSV
    auto.sleep(tempo_base)
    auto.click(185, 148)
    # Clicar no botão de Salvar
    auto.sleep(tempo_base)
    auto.leftClick(1307, 1010)
    # Salvar como botão direito
    auto.sleep(tempo_base)
    auto.click(1380, 985)
    # Selecionar pasta
    auto.sleep(tempo_base)
    auto.click(160, 55)
    auto.press('blackspace')
    # Colocar o caminho
    auto.sleep(tempo_base)
    pyperclip.copy(caminho)
    auto.hotkey('ctrl', 'v')
    auto.sleep(tempo_base)
    auto.press('enter')
    # selecionar o nome
    auto.sleep(tempo_base)
    auto.doubleClick(138, 921)
    # Colocar o nome
    auto.sleep(tempo_base)
    auto.write(nome)
    # Botão Salvar
    auto.sleep(tempo_base)
    auto.click(1742, 995)
    # Salvar
    auto.sleep(tempo_base)
    auto.press('left')
    auto.sleep(tempo_base)
    auto.press('enter')
    auto.sleep(tempo_base)
    auto.click()
    auto.sleep(15)
    auto.click(1894,10)


def fechar_promax():
    auto.sleep(15)
    pr.janela_certa("PromaxWEB - Trabalho — Microsoft​ Edge")
    auto.click(1894, 10)

def obs():
    # Ir na tela de rotina
    auto.sleep(5)
    pr.janela_certa("OBS 30.2.3 - Perfil: Sem nome - Cenas: Sem nome")
    auto.sleep(5)
    #Finalizar o aplicativo(1715,827)
    auto.click(1715,827)
    auto.sleep(5)
    auto.click(1898,0)
    auto.sleep(3)

def abri_rotina(rotina,tempo_base):
    # Ir na tela de rotina
    auto.sleep(5)
    pr.janela_certa("PromaxWEB - Trabalho — Microsoft​ Edge")
    auto.tripleClick(1743, 174)
    auto.sleep(tempo_base)
    # Escrever a rotina
    auto.write(rotina)
    # selecionar OK
    auto.click(1874,180)
    auto.sleep(3)


def salvar_csv_01_25_16_01(caminho,tempo_base,nome):
    # Selecionar CSV
    auto.sleep(tempo_base)
    auto.click(185, 148)
    # Clicar no botão de Salvar
    auto.sleep(tempo_base)
    auto.leftClick(1329, 1007)
    # Salvar como botão direito
    auto.sleep(tempo_base)
    auto.click(1380, 985)
    # Selecionar pasta
    auto.sleep(tempo_base)
    auto.click(167, 59)
    auto.press('blackspace')
    # Colocar o caminho
    auto.sleep(tempo_base)
    pyperclip.copy(caminho)
    auto.hotkey('ctrl', 'v')
    auto.press('enter')
    # selecionar o nome
    auto.sleep(tempo_base)
    auto.doubleClick(164, 894)
    # Colocar o nome
    auto.sleep(tempo_base)
    auto.write(nome)
    # Botão Salvar
    auto.sleep(tempo_base)
    auto.click(1742, 995)
    # Salvar
    auto.sleep(tempo_base)
    auto.press('left')
    auto.sleep(tempo_base)
    auto.press('enter')
    auto.sleep(tempo_base)
    auto.click()
    auto.sleep(15)
    auto.click(1894,10)

def salvar_csv_OCP(caminho,tempo_base,nome):
    tempo_base = 3
    # Selecionar CSV
    auto.sleep(tempo_base)
    auto.click(185, 148)
    # Clicar no botão de Salvar
    auto.sleep(tempo_base)
    auto.leftClick(1329,1007)
    # Salvar como botão direito
    auto.sleep(tempo_base)
    auto.click(1390, 985)
    # Selecionar pasta
    auto.sleep(tempo_base)
    auto.click(167, 59)
    auto.press('blackspace')
    # Colocar o caminho
    auto.sleep(tempo_base)
    pyperclip.copy(caminho)
    auto.hotkey('ctrl', 'v')
    auto.sleep(tempo_base)
    auto.press('enter')
    # selecionar o nome
    auto.sleep(tempo_base)
    auto.doubleClick(138, 921)
    # Colocar o nome
    auto.sleep(tempo_base)
    auto.write(nome)
    # Botão Salvar
    auto.sleep(tempo_base)
    auto.click(1742, 995)
    # Salvar
    auto.sleep(tempo_base)
    auto.press('left')
    auto.sleep(tempo_base)
    auto.press('enter')
    auto.sleep(tempo_base)
    auto.click()
    auto.sleep(15)
    auto.click(1894, 10)

def salvar(caminho,tempo_base,nome):
    # Selecionar CSV
    auto.sleep(tempo_base)
    auto.click(147, 146)
    # Clicar no botão de Salvar
    auto.sleep(tempo_base)
    auto.leftClick(1307, 1010)
    # Salvar como botão direito
    auto.sleep(tempo_base)
    auto.click(1380, 985)
    # Selecionar pasta
    auto.sleep(tempo_base)
    auto.click(167, 59)
    auto.press('blackspace')
    # Colocar o caminho
    auto.sleep(tempo_base)
    pyperclip.copy(caminho)
    auto.hotkey('ctrl', 'v')
    auto.sleep(tempo_base)
    auto.press('enter')
    # selecionar o nome
    auto.sleep(tempo_base)
    auto.doubleClick(164, 894)
    # Colocar o nome
    auto.sleep(tempo_base)
    auto.write(nome)
    # Botão Salvar
    auto.sleep(tempo_base)
    auto.click(1742, 995)
    # Salvar
    auto.sleep(tempo_base)
    auto.press('left')
    auto.sleep(tempo_base)
    auto.press('enter')
    auto.sleep(tempo_base)
    auto.click()
    auto.sleep(tempo_base)
    auto.hotkey('alt', 'f4')
#ROTINAS

def rotina_01_11 (tempo_base,local):
    abri_rotina("01.11",1)
    tempo_base = 1
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(1884, 547)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local,1,"01.11.csv")


def rotina_01_20_01_47 (tempo_base,sede,local):
    abri_rotina("01.20.01.47",1)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão classificação
    auto.sleep(tempo_base)
    auto.doubleClick(601, 206)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=2)
    # Apertar o botão pressionar
    auto.sleep(tempo_base)
    auto.click(1856, 417)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local, 5, f"{sede}.csv")

def rotina_03_01_47_01 (tempo_base,cdd,sede,local,revenda):
    abri_rotina("03.01.47.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar quebra 1
    auto.sleep(tempo_base)
    auto.doubleClick(488, 172)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1450, 279)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1538, 279)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    if cdd == 1:
        # Selecionar todos os CDDS da GEO
        auto.sleep(tempo_base)
        auto.click(328, 392)
        # Selecionar unidade
        auto.sleep(tempo_base)
        auto.click(328, 430)
         # Unir todas as unidades
        auto.sleep(tempo_base)
        auto.click(390, 316)
        # Selecionar o botao
        auto.sleep(tempo_base)
        auto.click(743, 514)
    else:
        pass
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1859, 554)
    # Esperar para carregar
    auto.sleep(80)
    # Esperar para carregar
    if revenda==1:
        salvar_csv(local, 2, f"Revendas.{dt.mes}.{dt.ano}.csv")
    else:
        salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_03_02_37 (tempo_base,sede,que1,que2,que3,item,local,entrada,data=1,operacao=2,salvar=1):
    abri_rotina("03.02.37",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar quebra 1
    auto.sleep(tempo_base)
    auto.doubleClick(478, 174)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=que1)
    # Selecionar quebra 2
    auto.sleep(tempo_base)
    auto.doubleClick(470, 195)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=que2)
    # Selecionar quebra 3
    auto.sleep(tempo_base)
    auto.doubleClick(470, 220)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=que3)
    # Selecionar Itens
    if item == 1:
        auto.sleep(tempo_base)
        auto.click(338, 378)
    else:
        pass
    if entrada == 1:
        auto.sleep(tempo_base)
        auto.click(339,430)
        #Lista NFs de Compra
        auto.sleep(tempo_base)
        auto.click(496,388)
        #Selecionar PIS/CONFINS
        auto.sleep(tempo_base)
        auto.click(496, 360)
        #Data Fiscal
        auto.sleep(tempo_base)
        auto.click(339, 486)
    else:
        pass
    # Selecionar Itens
    if data == 1:
        # Selecionar data inical
        auto.sleep(tempo_base)
        auto.tripleClick(1414, 257)
        # Inserir data inicial
        auto.sleep(tempo_base)
        auto.write(dt.data_inicial)
        # Selecionar data final
        auto.sleep(tempo_base)
        auto.tripleClick(1500, 257)
        # Inserir data final
        auto.sleep(tempo_base)
        auto.write(dt.dia_anterior)
    elif data == 2:
        # Selecionar data inical
        auto.sleep(tempo_base)
        auto.tripleClick(1414, 257)
        # Inserir data inicial
        auto.sleep(tempo_base)
        auto.write(dt.primeiro_dia_mes_ano_anterior)
        # Selecionar data final
        auto.sleep(tempo_base)
        auto.tripleClick(1500, 257)
        # Inserir data final
        auto.sleep(tempo_base)
        auto.write(dt.ultimo_dia_mes_ano_anterior)
    else:
        # Selecionar data inical
        auto.sleep(tempo_base)
        auto.tripleClick(1414, 257)
        # Inserir data inicial
        auto.sleep(tempo_base)
        auto.write(dt.primeiro_dia_mes_anterior)
        # Selecionar data final
        auto.sleep(tempo_base)
        auto.tripleClick(1500, 257)
        # Inserir data final
        auto.sleep(tempo_base)
        auto.write(dt.ultimo_dia_mes_anterior)

    #Operacação
    if operacao == 1:
        # Selecionar operacao inicial
        auto.sleep(tempo_base)
        auto.tripleClick(1395, 189)
        # Inserir operacao inicial
        auto.sleep(tempo_base)
        auto.write('251')
        # Selecionar operacao final
        auto.sleep(tempo_base)
        auto.tripleClick(1483, 189)
        # Inserir operacao final
        auto.sleep(tempo_base)
        auto.write('314')
    else:
         pass
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1863, 581)
    # Esperar para carregar
    auto.sleep(120)

    # Esperar para carregar
    if salvar==1:
        salvar_csv(local, 5, f"{sede}.{dt.mes}.{dt.ano}.csv")
    else:
        salvar_csv(local, 5, f"{dt.ano}_{dt.mes}_{sede}.inf")


def rotina_01_09 (tempo_base,local):
    abri_rotina("01.09",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(1883, 451)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local,2,"01.09.csv")

def rotina_01_12 (tempo_base,local):
    abri_rotina("01.12",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(1888, 475)
    # Esperar para carregar
    auto.sleep(10)

    salvar_csv(local,1,"01.12.csv")

def rotina_02_05_02 (tempo_base,sede,politica,local):
    abri_rotina("02.05.02",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Desmarcar opção Vasilhame/Garrafeira
    auto.sleep(tempo_base)
    auto.click(77, 408)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.doubleClick(1397, 293)
    # Inserir incial
    auto.sleep(tempo_base)
    auto.write("1")
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1475, 293)
    # Inserir final
    auto.sleep(tempo_base)
    auto.write("1")
    # Clicar no botao visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 483)
    # Esperar para carregar
    auto.sleep(15)
    if politica==1:
        salvar_csv(local, 1,f"{sede}.{dt.dia}.{dt.mes}.{dt.ano}.csv")
    else:
        salvar_csv(local,1,f"{sede}.csv")

def rotina_03_01_11 (tempo_base,sede,local):
    abri_rotina("03.01.11",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botao visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 536)
    # Esperar para carregar
    auto.sleep(10)
    salvar_csv(local,1,f"{sede}.csv")

def rotina_03_01_36_04 (tempo_base,sede,local,filtro):
    abri_rotina("03.01.36.04",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    if filtro == 1:
        # Deselecionar opção de Rejeitados
        auto.sleep(tempo_base)
        auto.click(214, 339)
        # Selecionar Quant. em Hectolitro
        auto.sleep(tempo_base)
        auto.click(657, 362)
        # Selecionar Não Lista Prod.em Falta
        auto.sleep(tempo_base)
        auto.click(657, 385)
    else:
        pass
    # Clicar no botao visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 556)
    # Esperar para carregar
    auto.sleep(10)
    salvar_csv(local,1,f"{sede}.csv")

def rotina_03_02_24(tempo_base,sede,filtro,local,salvar=1):
    abri_rotina("03.02.24",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classicar Vendedor
    auto.sleep(tempo_base)
    auto.doubleClick(540, 199)
    # Selecionar Vendedor
    auto.sleep(tempo_base)
    auto.press('down', presses=filtro)
    # Selecionar AS
    auto.sleep(tempo_base)
    auto.doubleClick(578, 326)
    # Selecionar Responsab.Processo
    auto.sleep(tempo_base)
    auto.click(527, 359)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1472, 201)
    # Inserir incial
    auto.sleep(tempo_base)
    auto.write(dt.primeiro_dia_tres_meses_antes)
    # Clicar no botao visualizar
    auto.sleep(tempo_base)
    auto.click(1854, 565)
    # Esperar para carregar
    auto.sleep(90)
    if salvar==1:
        salvar_csv(local,5,f"{sede}.{dt.mes}.{dt.ano}.csv")
    else:
        salvar_csv(local, 5, f"{sede}.csv")

def rotina_03_11_49_02 (tempo_base,sede,local):
    abri_rotina("03.11.49.02",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(940, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Todos os Mapas
    auto.sleep(tempo_base)
    auto.click(610, 325)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.doubleClick(1371, 233)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(str(dt.data_inicialf))
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.doubleClick(1611, 233)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(str(dt.dia_anteriorf))
    # Selecionar unidade
    auto.sleep(tempo_base)
    auto.click(328, 430)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1740, 396)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")
def rotina_03_11_20 (tempo_base,sede,local):
    abri_rotina("03.11.20",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar classificação
    auto.sleep(tempo_base)
    auto.doubleClick(495, 206)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1635, 200)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1765, 200)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar unidade
    auto.sleep(tempo_base)
    auto.click(328, 430)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 485)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_01_05_07_04_02 (tempo_base,sede,todos,local,salvar=1):
    abri_rotina("01.05.07.04.02",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(940, 17)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    if todos==1:
        # Selecionar o botão todos
        auto.sleep(tempo_base)
        auto.click(91, 198)
    elif todos==2:
        # Selecionar o botão bloqueados
        auto.sleep(tempo_base)
        auto.click(111,233)
    else:
        # Selecionar o botão bloqueados
        auto.sleep(tempo_base)
        auto.click(111, 233)
        # Selecionar em cadastramento
        auto.sleep(tempo_base)
        auto.click(108, 285)
    # Selecionar o AS
    auto.sleep(tempo_base)
    auto.click(1499, 228)
    # Botão Gerar CSV
    auto.sleep(tempo_base)
    auto.click(1836, 428)
    # Botão Esperar
    auto.sleep(180)
    # Enter
    auto.hotkey('enter')
    if salvar==1:
        salvar_csv(local, 1, f"{sede}.csv")
    else:
        salvar_csv(local, 1, f"{sede}.inf")

def rotina_03_11_40 (tempo_base,sede,local):
    abri_rotina("03.11.40",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1416, 198)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1492, 199)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 424)
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_03_11_35_04 (tempo_base,sede,local):
    abri_rotina("03.11.35.04",1)
    # Tempo_espera
    auto.sleep(5)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1409, 200)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.primeiro_dia_do_ano)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 430)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    salvar_csv(local, 3, f"{sede}.{dt.ano}.csv")

def rotina_01_20_01_24(tempo_base, sede,local):
    abri_rotina("01.20.01.24", 1)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão classificação
    auto.sleep(tempo_base)
    auto.doubleClick(601, 206)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=2)
    # Apertar o botão pressionar
    auto.sleep(tempo_base)
    auto.click(1856, 417)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local, 1,f"{sede}.csv")

def rotina_01_20_01_48 (tempo_base,sede):
    abri_rotina("01.20.01.48",1)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão classificação
    auto.sleep(tempo_base)
    auto.doubleClick(507, 174)
    # Pressionar o botão para baixo
    auto.sleep(tempo_base)
    auto.press('down', presses=2)
    # Apertar o botão pressionar
    auto.sleep(tempo_base)
    auto.click(1853, 321)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(r'C:\Users\pedro.alves\PycharmProjects\pythonProject\AUTOMA_REVEBI\RATING\01200148', 1, f"{sede}.inf")


def rotina_23_01_01_02 (tempo_base,sede,local):
    abri_rotina("23.01.01.02",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Desmarca Toda Geografia
    auto.sleep(tempo_base)
    auto.click(1282,161)
    # Marca Exebir Itens
    auto.sleep(tempo_base)
    auto.click(1466, 184)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(475,213)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.primeira_data)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(557,215)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.ultima_data)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1856, 468)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    #Clicar na opção salvar como
    auto.sleep(tempo_base)
    auto.click(819,582)
    auto.sleep(tempo_base)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_05_10 (tempo_base,sede,local):
    abri_rotina("05.10",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(552, 206)
    #Selecionar opção
    auto.sleep(tempo_base)
    auto.press('down')
    #Click na opção de Converte para Hectolitro
    auto.sleep(tempo_base)
    auto.click(540,456)
    #Click na opção de Totaliza Linha Marca
    auto.sleep(tempo_base)
    auto.click(402,481)
    #Selecionar mês
    auto.sleep(tempo_base)
    auto.doubleClick(528,278)
    #Inserir mes e ano
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}/{dt.ano}")
    #Marcar converte em Hectolitro
    auto.sleep(tempo_base)
    auto.click(540, 415)
    # Inserir mes e ano
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}/{dt.ano}")
    auto.sleep(tempo_base)
    #Inserir dias uteis
    auto.doubleClick(1482,246)
    auto.sleep(tempo_base)
    auto.write(dt.dias_uteis)
    auto.sleep(tempo_base)
    # Inserir dias acumulados
    auto.doubleClick(1480,268)
    auto.sleep(tempo_base)
    auto.write(dt.dias_uteis)
    auto.sleep(tempo_base)
    #Selecionar visualizar
    auto.sleep(tempo_base)
    auto.click(1853,529)
    auto.sleep(15)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_03_05_26 (tempo_base,sede,local):
    abri_rotina("03.05.26",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(600,204)
    # Pressionar colocar na opção Comercial
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1513,202)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1590, 202)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1862, 460)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_03_07_13 (tempo_base,sede,local):
    abri_rotina("03.07.13",1)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar visualizar
    auto.sleep(tempo_base)
    auto.click(1858,413)
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    salvar_csv(local, 1, f"{sede}.csv")

def rotina_02_02_20 (tempo_base,sede,local,salvar=1):
    abri_rotina("02.02.20",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(552,163)
    auto.sleep(tempo_base)
    auto.press('down')
    # Selecionar Todos
    auto.sleep(tempo_base)
    auto.click(335,377)
    #Selecionar Exibir Status do Documento
    auto.sleep(tempo_base)
    auto.click(1126,386)
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1852, 609)
    #Tempo de esperar do relatorio
    auto.sleep(150)
    if salvar == 1:
        salvar_csv(local, 2, f"{sede}.csv")
    else:
        salvar_csv(local, 2, f"{sede}.inf")

def rotina_02_02_29 (tempo_base,sede,local):
    abri_rotina("02.02.29",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(550,207)
    auto.sleep(tempo_base)
    auto.press('down')
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1852, 473)
    #Tempo de esperar do relatorio
    auto.sleep(30)
    salvar_csv(local, 2, f"{sede}.csv")

def rotina_02_02_53 (tempo_base,sede,local):
    abri_rotina("02.02.53",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(504,182)
    auto.sleep(tempo_base)
    auto.press('down',presses=3)
    #Selecionar Periodo
    auto.sleep(tempo_base)
    auto.tripleClick(482,249)
    auto.sleep(tempo_base)
    auto.write(f'{dt.mes}/{dt.ano}')
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1858,501)
    #Tempo de esperar do relatorio
    auto.sleep(30)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_15_05_01(tempo_base,sede,local):
    abri_rotina("15.05.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Ordem
    auto.sleep(tempo_base)
    auto.click(344, 217)
    # Ordem
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Competencia
    auto.sleep(tempo_base)
    auto.click(362,233)
    # Competencia
    auto.sleep(tempo_base)
    auto.press('down',presses=1)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(1171,184)
    # Inserir data
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}{dt.ano_atual}")
    # Selecionar NBZ
    auto.sleep(tempo_base)
    auto.click(1354,204)
    # Selecionar Depto
    auto.sleep(tempo_base)
    auto.click(1356,266)
    # Selecionar Pacote
    auto.sleep(tempo_base)
    auto.click(1356, 326)
    # Selecionar VBZ
    auto.sleep(tempo_base)
    auto.click(1356, 383)
    # Selecionar Conta
    auto.sleep(tempo_base)
    auto.click(1356, 450)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 537)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2,f"{sede}.{dt.ano}.{dt.mes}.csv")


def rotina_15_05_04(tempo_base,sede,local):
    abri_rotina("15.05.04",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Comprometido + Realizada
    auto.sleep(tempo_base)
    auto.click(401,421)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(1137,202)
    # Inserir data
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}{dt.ano_atual}")
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.doubleClick(1334, 203)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}{dt.ano_atual}")
    # Selecionar NBZ
    auto.sleep(tempo_base)
    auto.click(1321,226)
    # Selecionar Depto
    auto.sleep(tempo_base)
    auto.click(1322,284)
    # Selecionar Depto
    auto.sleep(tempo_base)
    auto.click(1322, 348)
    # Selecionar Pacote
    auto.sleep(tempo_base)
    auto.click(1322, 348)
    # Selecionar VBZ
    auto.sleep(tempo_base)
    auto.click(1322, 407)
    # Selecionar Conta
    auto.sleep(tempo_base)
    auto.click(1322, 465)
    # Selecionar Conta
    auto.sleep(tempo_base)
    auto.click(1852, 566)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2,f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_01_20_01_27 (tempo_base,sede,local,salvar=1):
    abri_rotina("01.20.01.27",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(544,196)
    auto.sleep(tempo_base)
    auto.press('down',presses=3)
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1851,361)
    #Tempo de esperar do relatorio
    auto.sleep(30)
    if salvar ==1:
        salvar_csv(local, 2, f"{sede}.csv")
    else:
        salvar_csv(local, 2, f"01200127_{sede}.inf")
def rotina_12_06_07(tempo_base,sede,local):
    abri_rotina("12.06.07",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(527,197)
    auto.sleep(tempo_base)
    auto.press('down',presses=5)
    #Selecionar data
    auto.sleep(tempo_base)
    auto.tripleClick(510, 217)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    #Selecionar especie 2
    auto.sleep(tempo_base)
    auto.tripleClick(1539, 260)
    auto.sleep(tempo_base)
    auto.write('2')
    # Selecionar especie 21
    auto.sleep(tempo_base)
    auto.tripleClick(1605, 260)
    auto.sleep(tempo_base)
    auto.write('21')
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1850,556)
    #Tempo de esperar do relatorio
    auto.sleep(240)
    auto.click(802,284)
    auto.sleep(240)
    auto.click(802, 284)
    auto.sleep(240)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_03_02_36_08 (tempo_base,sede,local):
    abri_rotina("03.02.36.08",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(937, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Inserir data inicial
    auto.sleep(tempo_base)
    auto.tripleClick(419,171)
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    #Inserir data final
    auto.sleep(tempo_base)
    auto.tripleClick(594, 171)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    #Selecionar visualizar
    auto.sleep(tempo_base)
    auto.click(1831,258)
    auto.sleep(15)
    salvar_csv_OCP(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_03_11_34_05(tempo_base,sede,local):
    abri_rotina("03.11.34.05",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar Classificação Geral
    auto.sleep(tempo_base)
    auto.doubleClick(504,207)
    auto.sleep(tempo_base)
    auto.press('down',presses=2)
    #Selecionar data inicial
    auto.sleep(tempo_base)
    auto.tripleClick(1420,200)
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1497,200)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1850,450)
    #Tempo de espera do relatorio
    auto.sleep(60)
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_01_20_11 (tempo_base,sede,variavel,local):
    abri_rotina("01.20.11",2)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar grupo de vendas
    auto.sleep(tempo_base)
    auto.doubleClick(570,226)
    # Pressionar opção segmentado
    auto.sleep(tempo_base)
    auto.press('down', presses=variavel)#Selecionar o opção
    #Selecionar ativos
    auto.sleep(tempo_base)
    auto.click(444,383)
    # Selecionar Bloqueados
    auto.sleep(tempo_base)
    auto.click(445,407)
    # Selecionar Temporario
    #auto.sleep(tempo_base)
    #auto.click(576,381)
    # Apertar o botão pressionar
    auto.sleep(tempo_base)
    auto.click(1855, 562)    # Esperar para carregar
    auto.sleep(20)
    # Esperar para carregar
    salvar_csv(local, 1, f"{sede}.{dt.mes}.{dt.ano}.csv")

def rotina_03_18_05(tempo_base,sede,local,salvar=1):
    abri_rotina("03.18.05",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(500, 206)
    auto.sleep(tempo_base)
    auto.press('down', presses=2)  # Selecionar Externa
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1850,560)
    #Tempo de espera do relatorio
    auto.sleep(180)
    if salvar ==1:
        salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")
    else:
        salvar_csv(local, 2, f"03.18.05_{sede}.inf")


def rotina_03_11_34_03 (tempo_base,sede,local):
    abri_rotina("03.11.34.03",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Data incial
    auto.sleep(tempo_base)
    auto.tripleClick(1406,200)
    auto.sleep(tempo_base)
    auto.write(dt.primeiro_dia_do_ano)
    # Data Final
    auto.sleep(tempo_base)
    auto.click(1483, 200)
    auto.tripleClick(tempo_base)
    auto.write(dt.dia_anteriorf)
    auto.sleep(tempo_base)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 430)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local,1,f"{sede}.csv")


def rotina_03_11_35_06 (tempo_base,sede,local):
    abri_rotina("03.11.35.06",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Classificação
    auto.sleep(tempo_base)
    auto.click(502, 205)
    auto.sleep(tempo_base)
    auto.press('down',1)
    # Data incial
    auto.sleep(tempo_base)
    auto.tripleClick(1419,200)
    auto.sleep(tempo_base)
    auto.write(dt.primeiro_dia_do_ano)
    # Data Final
    auto.sleep(tempo_base)
    auto.tripleClick(1495, 200)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anteriorf)
    auto.sleep(tempo_base)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 450)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local,1,f"{sede}.csv")


def rotina_01_02_24 (tempo_base,sede,local):
    abri_rotina("01.20.04",1)
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificação
    auto.sleep(tempo_base)
    auto.click(554, 207)
    auto.sleep(tempo_base)
    auto.press('down', 1)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 480)
    # Esperar para carregar
    auto.sleep(15)
    # Aperta o botão
    salvar_csv(local,1,f"010224_{sede}.inf")

def rotina_01_02_03 (tempo_base,local):
    abri_rotina("01.02.03",1)
    tempo_base = 1
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(940, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(168, 142)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local,1,"010203.inf")

def rotina_01_02_04 (tempo_base,local):
    abri_rotina("01.02.04",1)
    tempo_base = 1
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(940, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(168, 142)
    # Esperar para carregar
    auto.sleep(10)
    # Esperar para carregar
    salvar_csv(local,1,"010204.inf")

def rotina_03_05_09 (tempo_base,sede,data,local):
    abri_rotina("03.05.09",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(523,208)
    # Pressionar colocar na opção Cliente
    auto.sleep(tempo_base)
    auto.press('down', presses=7)
    #Opção Hectolitro
    auto.sleep(tempo_base)
    auto.click(583,333)
    if data == 1:
        # Selecionar data inical
        auto.sleep(tempo_base)
        auto.tripleClick(1455,202)
        # Inserir data inicial
        auto.sleep(tempo_base)
        auto.write(f"{dt.mes}{dt.ano}")
        # Selecionar data final
        auto.sleep(tempo_base)
        auto.tripleClick(1530, 202)
        # Inserir data final
        auto.sleep(tempo_base)
        auto.write(f"{dt.mes}{dt.ano}")
    else:
        # Selecionar data inical
        auto.sleep(tempo_base)
        auto.tripleClick(1455, 202)
        # Inserir data inicial
        auto.sleep(tempo_base)
        auto.write(f"{dt.tres_mes_atras}{dt.ano}")
        # Selecionar data final
        auto.sleep(tempo_base)
        auto.tripleClick(1530, 202)
        # Inserir data final
        auto.sleep(tempo_base)
        auto.write(f"{dt.mes}{dt.ano}")
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1853, 495)
    # Esperar para carregar
    auto.sleep(120)
    # Esperar para carregar
    salvar_csv(local, 2, f"{dt.ano}-{dt.mes}-{sede}.CSV")

def rotina_03_11_29 (tempo_base,sede,local):
    abri_rotina("03.11.29",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(500,200)
    # Pressionar colocar na opção Comercial
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1425,200)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1498,200)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 431)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    salvar_csv(local, 2, f"{dt.ano}-{dt.mes}-{sede}.inf")


def rotina_03_17_02 (tempo_base,sede,local):
    abri_rotina("03.17.02",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Todos
    auto.sleep(tempo_base)
    auto.click(432,378)
    # Selecionar todos os arquivos
    # Selecionar Tipo de Documentos Todos
    auto.sleep(tempo_base)
    auto.click(1440, 223)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1845, 566)
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    salvar_csv(local, 2, f"{dt.ano}_{dt.mes}_{sede}.inf")


def rotina_05_12 (tempo_base,sede,local):
    abri_rotina("05.12",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(557,198)
    # Selecionar Setor/Cliente
    auto.sleep(tempo_base)
    auto.press('down',presses=6)
    # Selecionar Totalizar Marca
    auto.sleep(tempo_base)
    auto.click(495, 446)
    # Selecionar Totalizar Embalagem
    auto.sleep(tempo_base)
    auto.click(495, 467)
    # Selecionar Totalizar Embalagem
    auto.sleep(tempo_base)
    auto.click(495, 487)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1851, 550)
    # Esperar para carregar
    auto.sleep(150)
    # Esperar para carregar
    salvar_csv(local, 2, f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_12_06_01(tempo_base,sede,local,especie_final,filtro,salvar=1):
    abri_rotina("12.06.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(940, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classicar Classificação
    auto.sleep(tempo_base)
    auto.doubleClick(495, 180)
    # Selecionar Especie
    auto.sleep(tempo_base)
    auto.press('down', presses=filtro)
    #Notas não atualizadas
    auto.sleep(tempo_base)
    auto.click(490, 301)
    # Dia anterior
    auto.sleep(tempo_base)
    auto.doubleClick(1504, 245)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    #Especie
    auto.sleep(tempo_base)
    auto.doubleClick(1490,350)
    auto.sleep(tempo_base)
    auto.write(str(especie_final))
    # Visualizar
    auto.sleep(tempo_base)
    auto.click(1852, 690)
    # Esperar para carregar
    auto.sleep(60)
    if salvar == 1:
        salvar_csv(local,2,f"{sede}.csv")
    else:
        salvar_csv(local, 2, f"{sede}.inf")
def rotina_12_06_04 (tempo_base,sede,local):
    abri_rotina("12.06.04",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificiação
    auto.sleep(tempo_base)
    auto.doubleClick(390,191)
    #Selecionar todos
    auto.sleep(tempo_base)
    auto.press('down',presses=10)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(390, 211)
    # Selecionar todos
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1412, 208)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1499,208)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 562)
    # Esperar para carregar
    auto.sleep(60)
    salvar_csv(local, 2,f"{sede}.{dt.mes}.{dt.ano}.inf")


def rotina_12_06_06 (tempo_base,sede,local):
    abri_rotina("12.06.06",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificiação
    auto.sleep(tempo_base)
    auto.doubleClick(516,201)
    #Selecionar todos
    auto.sleep(tempo_base)
    auto.press('down',presses=1)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(518, 220)
    # Selecionar todos
    auto.sleep(tempo_base)
    auto.press('down', presses=1)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1361, 238)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1538,238)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1851, 529)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2,f"{sede}.{dt.mes}.{dt.ano}.inf")

def rotina_15_05_03 (tempo_base,sede):
    abri_rotina("15.05.03",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificiação
    auto.sleep(tempo_base)
    auto.doubleClick(468,221)
    #Selecionar todos
    auto.sleep(tempo_base)
    auto.press('down',presses=6)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(480,247)
    # Selecionar todos
    auto.sleep(tempo_base)
    vezes = int(dt.mes)-1
    auto.press('down', presses=vezes)
    # Selecionar NBZ
    auto.sleep(tempo_base)
    auto.click(1466, 184)
    # Selecionar Depto
    auto.sleep(tempo_base)
    auto.click(1466, 244)
    # Selecionar Pacote
    auto.sleep(tempo_base)
    auto.click(1466, 304)
    # Selecionar Pacote
    auto.sleep(tempo_base)
    auto.click(1466, 364)
    # Selecionar VBZ
    auto.sleep(tempo_base)
    auto.click(1466, 364)
    # Selecionar VBZ
    auto.sleep(tempo_base)
    auto.click(1466, 425)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1851, 529)
    # Esperar para carregar
    auto.sleep(60)
    salvar_csv(r'Y:\BANCO DE DADOS CERVANTES\15 - OBZ\150503', 2,f"{sede}.{dt.mes}.{dt.ano}.inf")


def rotina_15_05_05(tempo_base,sede,local):
    abri_rotina("15.05.05",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Comprometido + Realizada
    auto.sleep(tempo_base)
    auto.click(407,374)
    # Data
    auto.sleep(tempo_base)
    auto.doubleClick(1137,202)
    # Inserir data
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}{dt.ano_atual}")
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.doubleClick(1334, 203)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(f"{dt.mes}{dt.ano_atual}")
    # Selecionar NBZ
    auto.sleep(tempo_base)
    auto.click(1321,226)
    # Selecionar Depto
    auto.sleep(tempo_base)
    auto.click(1322,288)
    # Selecionar Pacote
    auto.sleep(tempo_base)
    auto.click(1322, 344)
    # Selecionar VBZ
    auto.sleep(tempo_base)
    auto.click(1322, 407)
    # Selecionar Conta
    auto.sleep(tempo_base)
    auto.click(1322, 463)
    # Selecionar Visualizar
    auto.sleep(tempo_base)
    auto.click(1855, 573)
    # Esperar para carregar
    auto.sleep(30)
    salvar_csv(local, 2,f"{sede}.{dt.mes}.{dt.ano}.csv")


def rotina_15_05_10 (tempo_base,sede):
    abri_rotina("150510",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Clicar no botão CSV
    auto.sleep(tempo_base)
    auto.click(1855, 424)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    salvar(r'Y:\BANCO DE DADOS CERVANTES\15 - OBZ\150510', 1, f"150510_{sede}.csv")


def rotina_16_03_01 (tempo_base,sede):
    abri_rotina("16.03.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar o Data e Fiscal
    auto.sleep(tempo_base)
    auto.click(541, 377)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1483, 199)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1570, 199)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1854, 466)
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    salvar_csv(r'Y:\BANCO DE DADOS CERVANTES\16 - Contabilidade\160301', 2,f"{dt.ano}_{dt.mes}_{sede}.inf")



def rotina_16_03_05 (tempo_base,sede,local):
    abri_rotina("16.03.05",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar classificação
    auto.sleep(tempo_base)
    auto.click(590, 198)
    # Selecionar down
    auto.sleep(tempo_base)
    auto.press('down',2)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1860, 410)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    salvar(local, 2,f"{dt.ano}_{dt.mes}_{sede}.inf")

def rotina_16_03_08 (tempo_base,sede,local):
    abri_rotina("16.03.08",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar data inical
    auto.sleep(tempo_base)
    auto.tripleClick(1435, 200)
    # Inserir data inicial
    auto.sleep(tempo_base)
    auto.write(dt.data_inicial)
    # Selecionar data final
    auto.sleep(tempo_base)
    auto.tripleClick(1522, 200)
    # Inserir data final
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1852,391)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.hotkey('enter')
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    salvar_csv(local, 2,f"{dt.ano}-{dt.mes}-{sede}.csv")

def rotina_21_04_07 (tempo_base,sede):
    abri_rotina("21.04.07",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(941, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar todos
    auto.sleep(tempo_base)
    auto.click(95, 213)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1830,650)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    auto.sleep(tempo_base)
    auto.hotkey('enter')
    salvar_csv(r'C:\Users\pedro.alves\PycharmProjects\pythonProject\BANCO_DE_DADOS\210407', 2,f"{dt.ano}-{dt.mes}-{sede}.inf")

def rotina_01_20_12 (tempo_base,sede):
    abri_rotina("01.20.12",1)
    # Tempo_espera
    auto.sleep(3)
    pr.janela_certa('Clientes (Gerador) [PEDRO - pedro] - Trabalho — Microsoft​ Edge')
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Inserir codigo do relatorio
    auto.sleep(tempo_base)
    auto.click(475,195)
    auto.write('11')
    # Colocar classificação numerica
    auto.sleep(tempo_base)
    auto.doubleClick(503,220)
    auto.sleep(tempo_base)
    auto.press('down',presses=1)
    #Desmarca Total e seleciona Ativo e Bloqueado
    auto.sleep(tempo_base)
    auto.click(407,364)
    auto.sleep(tempo_base)
    auto.click(407,388)
    auto.sleep(tempo_base)
    auto.click(407,413)
    #Inserir momento 0
    auto.sleep(tempo_base)
    auto.doubleClick(1400, 480)
    auto.sleep(tempo_base)
    auto.write('0')
    #Visualizar
    auto.sleep(tempo_base)
    auto.doubleClick(1850,560)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    auto.sleep(tempo_base)
    auto.click(148,147)
    auto.sleep(tempo_base)
    salvar_csv(r'\\192.168.1.1\arquivos\COMPARTILHADOS\BANCO DE DADOS\VENDAS\01.20.12', 2,f"{dt.ano}-{dt.mes}-{sede}.inf")



def rotina_01_25_16_01 (tempo_base,sede):
    abri_rotina("01.25.16.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(941, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1451,717)
    # Selecionar o exportação
    auto.sleep(tempo_base)
    auto.click(1214,607)
    # Esperar para carregar
    auto.sleep(30)
    # Esperar para carregar
    auto.sleep(tempo_base)
    salvar_csv_01_25_16_01(r'\\192.168.1.1\arquivos\COMPARTILHADOS\BANCO DE DADOS\VENDAS\01.25.16.01', 2,f"{dt.ano}-{dt.mes}-{sede}.inf")

def rotina_02_05_03 (tempo_base,sede):
    abri_rotina("02.05.03", 1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(731, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Selecionar o periodo anterior
    auto.sleep(tempo_base)
    auto.tripleClick(1415,228)
    auto.sleep(tempo_base)
    auto.write('01/01/2023')
    #Definir os depositos inicial
    auto.sleep(tempo_base)
    auto.tripleClick(1398,295)
    auto.sleep(tempo_base)
    auto.write('51')
    auto.sleep(tempo_base)
    auto.tripleClick(1474,293)
    auto.sleep(tempo_base)
    auto.write('51')
    # Selecionar o visualizar
    auto.sleep(tempo_base)
    auto.click(1851, 481)
    # Esperar para carregar
    auto.sleep(60)
    # Esperar para carregar
    auto.sleep(tempo_base)
    salvar_csv(r'\\192.168.1.1\arquivos\COMPARTILHADOS\BANCO DE DADOS\VENDAS\02.05.03 MATERIAIS LEVE', 2,f"{dt.ano}-{dt.mes}-{sede}.inf")

def rotina_SIV_03_03_04(tempo_base,sede):
    auto.sleep(5)
    pr.janela_certa("PromaxWEB - Trabalho — Microsoft​ Edge")
    # Acessar rotina
    auto.sleep(tempo_base)
    auto.click(93, 303)
    # Maximizar
    auto.sleep(tempo_base)
    auto.click(941, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    #Selecionar operacão
    auto.sleep(tempo_base)
    auto.click(61, 225)
    # Selecionar a 3.3
    auto.sleep(tempo_base)
    auto.click(70,280)
    # Selecionar a 3.3.4
    auto.sleep(tempo_base)
    auto.click(110,355)
    #Exportar
    auto.sleep(tempo_base)
    auto.click(741,437)
    #Espera um pouco e da enter
    auto.sleep(tempo_base)
    auto.press('enter',presses=1)
    salvar_csv(r'\\192.168.1.1\arquivos\COMPARTILHADOS\BANCO DE DADOS\VENDAS\3.3.4',1,f"{dt.ano}-{dt.mes}-{sede}.inf")


def rotina_01_20_04 (tempo_base,sede,local):
    abri_rotina("01.20.04",1)
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Classificação
    auto.sleep(tempo_base)
    auto.click(554, 207)
    auto.sleep(tempo_base)
    auto.press('down', 1)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 480)
    # Esperar para carregar
    auto.sleep(15)
    # Aperta o botão
    salvar_csv(local,1,f"012004_{sede}.inf")


def rotina_23_01_01_01(tempo_base,sede,local):
    abri_rotina("23.01.01.01",1)
    # Tempo_espera
    auto.sleep(3)
    # Maximizar
    auto.click(727, 15)
    # Fechar aba Explore
    auto.sleep(tempo_base)
    auto.click(1888, 75)
    # Data incial
    auto.sleep(tempo_base)
    auto.tripleClick(1422,200)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    # Data Final
    auto.sleep(tempo_base)
    auto.tripleClick(1492, 200)
    auto.sleep(tempo_base)
    auto.write(dt.dia_anterior)
    auto.sleep(tempo_base)
    #Visualizar
    auto.sleep(tempo_base)
    auto.click(1850, 563)
    #Clicar enter
    auto.sleep(30)
    # Clicar enter
    auto.sleep(tempo_base)
    auto.press('enter', presses=1)
    salvar_csv(local,1,f"23010101_{sede}.inf")



