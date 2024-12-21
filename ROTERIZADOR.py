from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import subprocess
import pyautogui as auto
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import time
import glob
import os
import shutil


def get_previous_weekday(date):
    weekday = date.weekday()

    # Domingo (6) e Segunda (0) -> retorna sexta-feira anterior
    if weekday == 6 or weekday == 0:
        previous_weekday = date - timedelta(days=(2 if weekday == 0 else 1))
    else:
        # Qualquer outro dia retorna o dia útil anterior
        previous_weekday = date - timedelta(days=1)

    return previous_weekday




# Exemplo de uso:
hoje = datetime.now()
dia_anterior = get_previous_weekday(hoje)
dia_anterior=dia_anterior.strftime("%d%m%Y")




def renomear_arquivo_por_nome(download_dir, parte_nome, novo_nome_completo):
    """
    Função que busca o arquivo mais recente que contém uma parte específica do nome e renomeia-o.

    :param download_dir: Diretório onde os arquivos estão localizados.
    :param parte_nome: Parte do nome que o arquivo deve conter para ser considerado.
    :param novo_nome_completo: O novo nome completo que o arquivo deve ter (com caminho e extensão).
    """
    # Lista todos os arquivos no diretório de download
    files = os.listdir(download_dir)

    # Filtra os arquivos que contêm a parte do nome desejado
    arquivos_filtrados = [f for f in files if parte_nome in f]

    # Verifica se há arquivos que correspondem ao filtro
    if arquivos_filtrados:
        # Cria os caminhos completos para os arquivos filtrados
        full_paths = [os.path.join(download_dir, f) for f in arquivos_filtrados]

        # Obtém o arquivo mais recente entre os que contêm a parte do nome
        arquivo_mais_recente = max(full_paths, key=os.path.getctime)

        # Verifica se o arquivo com o novo nome já existe e o exclui
        if os.path.exists(novo_nome_completo):
            # Certifica-se de que o arquivo a ser excluído não é o mesmo que será renomeado
            if os.path.abspath(arquivo_mais_recente) != os.path.abspath(novo_nome_completo):
                os.remove(novo_nome_completo)
                print(f"Arquivo antigo '{novo_nome_completo}' excluído.")

        # Renomeia o arquivo mais recente
        shutil.move(arquivo_mais_recente, novo_nome_completo)
        print(f"Arquivo '{arquivo_mais_recente}' renomeado para '{novo_nome_completo}'.")
    else:
        print(f"Nenhum arquivo encontrado contendo '{parte_nome}'.")

def substituir_arquivo(arquivo_novo, arquivo_destino):
    """
    Função para substituir um arquivo antigo por outro, sem exclusão manual.

    :param arquivo_novo: Caminho completo do novo arquivo que será copiado.
    :param arquivo_destino: Caminho completo do arquivo de destino que será sobrescrito.
    """
    # Copiar o arquivo novo para o local do arquivo de destino (substitui automaticamente)
    shutil.copy(arquivo_novo, arquivo_destino)

def substituir_arquivo_com_nome_especifico(directory, arquivo_destino, nome_parcial):
    """
    Função para substituir um arquivo de destino pelo arquivo mais recente de um diretório,
    contanto que o nome do arquivo mais recente contenha uma palavra-chave específica.

    :param directory: Caminho para o diretório onde os arquivos estão.
    :param arquivo_destino: Caminho completo do arquivo de destino que será substituído.
    :param nome_parcial: Parte do nome que o arquivo mais recente deve conter para ser considerado.
    """
    # Obter todos os arquivos no diretório
    full_paths = glob.glob(os.path.join(directory, "*"))

    # Pegar o arquivo mais recente
    latest_file = max(full_paths, key=os.path.getctime)

    # Verificar se o nome do arquivo mais recente contém o nome_parcial
    if nome_parcial in os.path.basename(latest_file):
        # Se o nome parcial estiver presente, copiar o arquivo para o destino (substitui automaticamente)
        shutil.copy(latest_file, arquivo_destino)
    else:
        pass


# Obter o ano atual
ano_atual = datetime.now().year

# Primeiro dia do ano
primeiro_dia_do_ano = datetime(ano_atual, 1, 1)
primeiro_dia_do_ano=primeiro_dia_do_ano.strftime("%d/%m/%Y")

# Obter a data atual
hoje = datetime.today()
# Verificar se o dia atual é o primeiro dia do mês
if hoje.day == 1:
    # Se for o primeiro dia, calcular o primeiro dia do mês anterior
    data_inicial = (hoje - relativedelta(months=1)).replace(day=1)
    # Último dia do mês anterior
    ultimo_dia = hoje.replace(day=1) - relativedelta(days=1)
    # Dia anterior (do último mês)
    dia_anterior = ultimo_dia
else:
    # Primeiro dia do mês atual
    data_inicial = hoje.replace(day=1)
    # Último dia do mês atual
    ultimo_dia = (hoje + relativedelta(months=1)).replace(day=1) - relativedelta(days=1)
    # Dia anterior (do dia atual)
    dia_anterior = hoje - relativedelta(days=1)

data_inicialf=data_inicial.strftime("%d%m%Y")
dia_anteriorf=dia_anterior.strftime("%d%m%Y")
data_inicial=data_inicial.strftime("%d/%m/%Y")
dia_anterior=dia_anterior.strftime("%d/%m/%Y")
mes=ultimo_dia.strftime("%m")
ano=ultimo_dia.strftime("%Y")

# Subtraindo 3 meses
tres_meses_antes = hoje - relativedelta(months=3)
# Obtendo o primeiro dia do mês
primeiro_dia_tres_meses_antes = tres_meses_antes.replace(day=1)
primeiro_dia_tres_meses_antes=primeiro_dia_tres_meses_antes.strftime("%d/%m/%Y")

def rodar_moc():

    # Configurações do ChromeDriver
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Inicia o Chrome maximizado
    # Substitua pelo caminho do seu ChromeDriver
    download_dir = r"X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ"
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,  # Diretório de download padrão
        "download.prompt_for_download": False,  # Desabilita a janela de confirmação de download
        "download.directory_upgrade": True,  # Atualiza o diretório se já estiver configurado
        "safebrowsing.enabled": True  # Habilita downloads seguros
    })


    # Inicializa o WebDriver
    driver = webdriver.Chrome(options=chrome_options)
    # Acessa o site
    driver.get("https://roteirizador.app/#/login")
    time.sleep(20)
    # Espera até que o campo de e-mail esteja presente
    wait = WebDriverWait(driver, 20)
    # Insere o e-mail
    email_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/div[2]/input")))
    email_field.send_keys("nicole.duarte@cervantes.com.br")
    # Insere a senha
    email_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/div[3]/input")))
    email_field.send_keys("Stella123@")
    # Clica no botão de login
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/button")))
    login_button.click()

    time.sleep(60)

    #Cliente
    # Agora que está logado, acessa outro site na mesma janela
    driver.get("https://roteirizador.app/#/relatorio/cliente")  # Substitua pelo site que deseja acessar
    time.sleep(20)
    # Clica no componente <ng-select> para abrir as opções
    ng_select = wait.until(EC.element_to_be_clickable((By.ID, "formly_5_ng-select_idDeposito_0")))
    ng_select.click()
    # Espera até que as opções estejam visíveis (encontre o seletor correto das opções)
    opcao_desejada = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/ng-dropdown-panel/div/div[2]/div[1]")))
    opcao_desejada.click()
    #Cliclar no botão filtra
    time.sleep(40)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    #Cliclar no botão Exportacao
    time.sleep(60)
    export_button =WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    time.sleep(60)
   
    #Indicador roterizador
    # Agora que está logado, acessa outro site
    driver.get("https://roteirizador.app/#/relatorio/indicadores-roteirizacao")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    # Clica no componente <ng-select> para abrir as opções
    ng_select = wait.until(EC.element_to_be_clickable(((By.XPATH, "//input[@type='text' and @aria-autocomplete='list']"))))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-0")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():
        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(60)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH, "//input[@placeholder='__/__/____']")    # Insira a data no formato desejado

    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    #Cliclar no botão filtra
    time.sleep(15)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    # Cliclar no botão Exportacao
    time.sleep(60)
    export_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    time.sleep(60)

    #Roterizador
    # Agora que está logado, acessa outro site
    driver.get("https://roteirizador.app/#/relatorio/roteirizacao")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    auto.click(0,1)

    # Clica no componente <ng-select> para abrir as opções

    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-autocomplete='list']")))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-0")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():

        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(30)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.CLASS_NAME, "form-control")

    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    #Cliclar no botão filtra
    time.sleep(15)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    #Cliclar no botão exportar
    time.sleep(60)
    export_button = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    time.sleep(30)
    auto.hotkey('esc')
    time.sleep(60)

    #Romaneio
    # Agora que está logado, acessa outro site na mesma janela
    driver.get("https://roteirizador.app/#/relatorio/romaneio")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    # Clica no componente quem tem as opções de DEPOSITO
    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-0")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():
        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(30)
    #
    #Filtra
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    time.sleep(60)
    #Botão Exportar
    export_button = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    time.sleep(30)

    #Pedidos
    
    # Agora que está logado, acessa outro site
    # Substitua pelo site que deseja acessar

    driver.get("https://roteirizador.app/#/pedido")
    # Localizar o campo de entrada
    time.sleep(30)
    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox']//input[@aria-autocomplete='list']")))
    ng_select.click()
    opcao_desejada = wait.until(
        EC.element_to_be_clickable((By.XPATH,"/html/body/ng-dropdown-panel/div/div[2]/div[1]")))
    opcao_desejada.click()
    time.sleep(30)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH,"//input[@placeholder='__/__/____']")  # Insira a data no formato desejado
    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    #Cliclar no botão filtraf
    time.sleep(60)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'btn-filter')))
    botonfilter.click()
    time.sleep(30)
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)
    # Cliclar no botão exportar
    botao = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//button[span[text()=' Exportar ']]"))
    )
    time.sleep(60)
    botao.click()
    time.sleep(150)
    # Botão OK que aparece na tela
    button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Continuar']")))
    # Clicar no botão
    button.click()
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)





    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ', 'romaneio', r'\\192.168.1.1\arquivos\COMPARTILHADOS\WHATSAPP AUTOMÁTICO\BASES\ROTEIRIZADOR\ROTEIRIZADOR_MOC.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ', 'cliente', r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\Clientes\MATRIZ\Matriz.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ','relatorio-roteirizacao','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Roteirização\MATRIZ\Matriz.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ','indicadores-roteirizacao','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Indicadores Roteirização\MATRIZ\Matriz.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\MATRIZ','pedidos','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Pedidos\MATRIZ\Matriz.csv')

    auto.click(0,1)
    auto.click(0,1)

def rodar_bm():
    # Definindo o diretório de download
    download_dir = r"X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL"  # Insira aqui o caminho completo da pasta desejada
    # Configurações do ChromeDriver
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Inicia o Chrome maximizado
    # Substitua pelo caminho do seu ChromeDriver

    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,  # Diretório de download padrão
        "download.prompt_for_download": False,  # Desabilita a janela de confirmação de download
        "download.directory_upgrade": True,  # Atualiza o diretório se já estiver configurado
        "safebrowsing.enabled": True  # Habilita downloads seguros
    })

    # Inicializa o WebDriver
    driver = webdriver.Chrome(options=chrome_options)
    # Acessa o site
    driver.get("https://roteirizador.app/#/login")
    time.sleep(30)
    # Espera até que o campo de e-mail esteja presente
    wait = WebDriverWait(driver, 30)
    # Insere o e-mail
    email_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/div[2]/input")))
    email_field.send_keys("nicole.duarte@cervantes.com.br")
    # Insere a senha
    email_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/div[3]/input")))
    email_field.send_keys("Stella123@")
    # Clica no botão de login
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/hb-app/hb-login-template/hb-login/div/div[2]/form/button")))
    login_button.click()

    time.sleep(30)


    auto.click(0,1)
    #Cliente
    # Agora que está logado, acessa outro site na mesma janela
    driver.get("https://roteirizador.app/#/relatorio/cliente")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    # Clica no componente <ng-select> para abrir as opções
    ng_select = wait.until(EC.element_to_be_clickable((By.ID, "formly_5_ng-select_idDeposito_0")))
    ng_select.click()
    # Selecionar a cidade
    opcao_desejada = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/ng-dropdown-panel/div/div[2]/div[2]")))
    opcao_desejada.click()
    # Filtra
    time.sleep(15)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    time.sleep(15)
    botonfilter.click()
    time.sleep(30)
    # Botão Exportar
    time.sleep(60)
    export_button = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    #Indicador roterizador
    # Agora que está logado, acessa outro site
    driver.get("https://roteirizador.app/#/relatorio/indicadores-roteirizacao")  # Substitua pelo site que deseja acessar
    time.sleep(15)
    # Clica no componente <ng-select> para abrir as opções
    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text' and @aria-autocomplete='list']")))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-1")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():
        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(15)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input=driver.find_element(By.CLASS_NAME, "form-control")
    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    # Cliclar no botão filtra
    time.sleep(15)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    # Botão Exportar
    time.sleep(60)
    export_button = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
 

      
    #Roterizador
    # Agora que está logado, acessa outro site
    driver.get("https://roteirizador.app/#/relatorio/roteirizacao")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    auto.click(0,1)
    # Clica no componente <ng-select> para abrir as opções
    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-autocomplete='list']")))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-1")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():
        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(30)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH,"//input[@placeholder='__/__/____']")

    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    # Filtra
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    time.sleep(30)
    # Botão Exportar
    time.sleep(60)
    export_button = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    auto.click(0,1)#Coloquei esse enter para a tela não desativar
    time.sleep(30)
    time.sleep(60)
     
    #Romaneio
    # Agora que está logado, acessa outro site na mesma janela
    auto.click(0,1)#Coloquei esse enter para a tela não desativar
    driver.get("https://roteirizador.app/#/relatorio/romaneio")  # Substitua pelo site que deseja acessar
    time.sleep(30)
    # Clica no componente quem tem as opções de DEPOSITO
    ng_select = wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@type='text']")))
    ng_select.click()
    # Localize o checkbox pelo ID
    checkbox = driver.find_element(By.ID, "item-1")
    # Verifique se o checkbox já está marcado
    if not checkbox.is_selected():
        # Se não estiver marcado, clique para marcar
        checkbox.click()
    time.sleep(30)
    #data_inicial
    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH, "//input[@placeholder='__/__/____']")    # Insira a data no formato desejado
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys(dia_anterior)
    #data_final
    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH, "/html/body/hb-app/hb-app-template/div/div[2]/div/hb-relatorio-romaneio/hb-relatorio-container/div/p-sidebar/div/div[2]/formly-form/formly-field/formly-group/formly-field[4]/formly-wrapper-form-field/div/hb-datepicker/p-calendar/span/input")
    # Insira a data no formato desejado
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys(dia_anterior)
    # Filtra
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    time.sleep(60)
    # Botão Exportar
    export_button = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()

    #Pedido

    # Agora que está logado, acessa outro site
    driver.get("https://roteirizador.app/#/pedido")  # Substitua pelo site que deseja acessar
    time.sleep(15)
    auto.click(0,1)

    time.sleep(15)
    ng_select = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox']//input[@aria-autocomplete='list']")))
    ng_select.click()
    opcao_desejada = wait.until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/ng-dropdown-panel/div/div[2]/div[2]')))
    opcao_desejada.click()

    time.sleep(15)

    # Localize o campo de entrada de data (pelo seletor adequado, neste caso uma classe)
    date_input = driver.find_element(By.XPATH,
                                     "//input[@placeholder='__/__/____']")  # Insira a data no formato desejado
    # Insira a data no formato desejado (exemplo: "15/10/2024")
    time.sleep(15)
    driver.execute_script("arguments[0].value = '';", date_input)
    time.sleep(15)
    date_input.send_keys("01012024")
    # Cliclar no botão filtraf
    time.sleep(60)
    botonfilter = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-filter')))
    botonfilter.click()
    time.sleep(15)
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)
    time.sleep(210)
    auto.click(0, 1)
    # Cliclar no botão exportar
    export_button = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()=' Exportar ']")))
    time.sleep(60)
    export_button.click()
    time.sleep(150)
    # Botão OK que aparece na tela
    button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Continuar']")))
    # Clicar no botão
    button.click()
    auto.click(0,1)
    time.sleep(210)
    auto.click(0,1)

    driver.close()


    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL', 'cliente', r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\Clientes\FILIAL\Filial.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL','relatorio-roteirizacao','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Roteirização\FILIAL\Filial.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL','indicadores-roteirizacao','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Indicadores Roteirização\FILIAL\Filial.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL','pedidos','Z:\PRODUTIVIDADE\ROTEIRIZADOR\Pedidos\FILIAL\Filial.csv')
    renomear_arquivo_por_nome(r'X:\ADMINISTRATIVO\BI\01. ANALISTA BI PEDRO\BASE\ROTERIZADOR\FILIAL', 'romaneio', r'\\192.168.1.1\arquivos\COMPARTILHADOS\WHATSAPP AUTOMÁTICO\BASES\ROTEIRIZADOR\ROTEIRIZADOR_BM.csv')

rodar_moc()
rodar_bm()







