import glob
import os
import shutil
from datetime import datetime
import time

def renomear_arquivo_por_nome(download_dir, parte_nome, novo_nome_completo):
    """
    Função que busca o arquivo mais recente que contém uma parte específica do nome e renomeia-o.

    :param download_dir: Diretório onde os arquivos estão localizados.
    :param parte_nome: Parte do nome que o arquivo deve conter para ser considerado.
    :param novo_nome_completo: O novo nome completo que o arquivo deve ter (com caminho e extensão).p
    """
    # Lista todos os arquivos no diretório de download
    files = os.listdir(download_dir)

    # Filtra os arquivos que contêm a parte do nome desejado
    arquivos_filtrados = [f for f in files if parte_nome in f]

    # Verifica se há arquivos que correspondem ao filtro
    try:
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
    except:
        pass

def substituir_arquivo(arquivo_novo, arquivo_destino):
    """
    Função para substituir um arquivo antigo por outro, sem exclusão manual.

    :param arquivo_novo: Caminho completo do novo arquivo que será copiado.
    :param arquivo_destino: Caminho completo do arquivo de destino que será sobrescrito.
    """
    # Copiar o arquivo novo para o local do arquivo de destino (substitui automaticamente)
    try:
        shutil.copy(arquivo_novo, arquivo_destino)
    except:
        pass

def substituir_arquivo_com_nome_especifico(directory, arquivo_destino, nome_parcial):
    time.sleep(1)
    """
    Função para substituir um arquivo de destino pelo arquivo mais recente de um diretório,
    contanto que o nome do arquivo mais recente contenha uma palavra-chave específica.

    :param directory: Caminho para o diretório onde os arquivos estão.
    :param arquivo_destino: Caminho completo do arquivo de destino que será substituído.
    :param nome_parcial: Parte do nome que o arquivo mais recente deve conter para ser considerado.
    """
    # Obter todos os arquivos no diretório
    try:
        full_paths = glob.glob(os.path.join(directory, "*"))

        # Pegar o arquivo mais recente
        latest_file = max(full_paths, key=os.path.getctime)

        # Verificar se o nome do arquivo mais recente contém o nome_parcial
        if nome_parcial in os.path.basename(latest_file):
            # Se o nome parcial estiver presente, copiar o arquivo para o destino (substitui automaticamente)
            shutil.copy(latest_file, arquivo_destino)
        else:
            pass
    except:
        pass


def verificar_arquivos_parciais(pasta, nome_parcial):
    arquivos_nao_baixados = []
    data_hoje = datetime.now().date()
    arquivo_encontrado_hoje = False

    # Lista todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
        if arquivo.startswith(nome_parcial):  # Verifica o nome parcial
            caminho_arquivo = os.path.join(pasta, arquivo)

            # Verifica se é um arquivo e obtém a data de modificação
            if os.path.isfile(caminho_arquivo):
                data_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho_arquivo)).date()

                # Verifica se o arquivo foi modificado hoje
                if data_modificacao == data_hoje:
                    arquivo_encontrado_hoje = True
                    print(f"Arquivo '{arquivo}' foi baixado/modificado hoje.")
                    break

    # Caso nenhum arquivo com o nome parcial tenha sido atualizado hoje
    if not arquivo_encontrado_hoje:
        arquivos_nao_baixados.append(nome_parcial)
        print(f"Arquivo com nome parcial '{nome_parcial}' precisa ser baixado novamente.")

    # Salva o log dos arquivos não baixados em um txt
    if arquivos_nao_baixados:
        with open("log_arquivos_nao_baixados.txt", "a", encoding="utf-8") as log_file:
            data_validacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_file.write(f"Validação realizada em: {data_validacao}\n")
            for arquivo in arquivos_nao_baixados:
                log_file.write(f"Arquivo não encontrado ou desatualizado:{pasta} {arquivo}\n")
            log_file.write("\n")  # Linha em branco para separação entre validações

    return arquivos_nao_baixados

def verificar_arquivo(pasta, parte_nome,rotina):
    data_hoje = datetime.now().date()
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if parte_nome in arquivo and os.path.isfile(caminho_arquivo):
            data_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho_arquivo)).date()
            if data_modificacao == data_hoje:
                print(f"O arquivo '{arquivo}' foi baixado hoje.")
                return True
    return rotina




