from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import os
import xlwings as xw
import shutil
import tempfile




download_dir = "M:\\ADM DE VENDAS PJ\\Diario Imput\\DownloadArquivos\\arquivos baixados"
chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Para abrir o navegador maximizado

prefs = {
    "profile.default_content_settings.popups": 0,  # Bloqueia popups de confirmação
    "download.default_directory": download_dir,  # Diretório de download
    "download.prompt_for_download": False,  # Desativa o prompt de confirmação de download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True  # Habilita a navegação segura (impede mensagens de segurança)
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

#função para diminuir o código
def acao_site(elemento):

    #codigo para expendir o campo para buscar data
    campo_expandir = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, elemento))
    )
    campo_expandir.click()

def expandir_painel(xpath, path):

    #encontra o painel de expansão de download e gera um click para expandir
    campo_download = driver.find_element(By.XPATH, xpath)

    # Aguarda o campo de data e abre o calendário
    campo_download = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, path))
    )
    campo_download.click()

def downloadArquivos(xpath, path):
    download_1 = driver.find_element(By.XPATH, xpath)
    download_1 = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.ID, path))
    )
    download_1.click()
    sleep(1)

try:
    #localiza a endereço e entrar no site MSGestor
    driver.get("https://msgestor.msconnect.com.br/pages/auth/login")
    sleep(2)

    #Localiza os campos de login Usuário e senha
    campo_usuario = driver.find_element(By.XPATH,"//input[@id='mat-input-26']")
    campo_senha = driver.find_element(By.XPATH,"//input[@id='mat-input-27']")
    sleep(1)
    #Preencge os campos de login usuário e senha
    campo_usuario.send_keys('allef.sousa')
    campo_senha.send_keys('98638C3')
    sleep(1)
    #Gera um click no ENTER para realizar o Login
    campo_senha.send_keys(Keys.RETURN)

    sleep(5)
except Exception as e:
    print(f"Erro durante o login: {e}")


try:
    sleep(1)
    #localiza o botão de configuração e gera um click
    acao_site("//button[contains(@class, 'btnSettings')]")

except Exception as e:
    print("Erro ao encontrar o botão:", e)



try:
    #localiza o painel de expanção para pesquisar a data e expande
    acao_site("//mat-expansion-panel-header[@id='mat-expansion-panel-header-16']")

except Exception as e:
    print("Erro ao encontrar ou expandir aba:", e)


#Aqui irá encontrar os campos de data, criar um modal para pausar a aplicação para preenchimento manual das datas
try:

    # Localiza os campos de data
    campo_data_inicio = driver.find_element(By.XPATH, "//input[@id='mat-input-28']")
    campo_data_final = driver.find_element(By.XPATH, "//input[@id='mat-input-29']")

    # Aguarda o campo de data e abre o calendário
    campo_data_inicio = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "mat-input-28"))
    )
    campo_data_inicio.click()


    # Cria uma div na página com uma mensagem para preencher as datas desejadas
    driver.execute_script("""
        var popup = document.createElement('div');
        popup.style.position = 'fixed';
        popup.style.top = '50%';
        popup.style.left = '30%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
        popup.style.color = 'white';
        popup.style.padding = '20px';
        popup.style.borderRadius = '10px';
        popup.style.fontSize = '18px';
        popup.style.zIndex = '9999';  // Definindo um z-index alto para garantir que o popup fique à frente
        popup.innerHTML = 'Selecione a data manualmente e clique em OK para continuar.';
        popup.id = 'popupAlert';

        var button = document.createElement('button');
        button.innerText = 'OK';
        button.style.marginTop = '10px';
        button.style.backgroundColor = '#4CAF50';
        button.style.color = 'white';
        button.style.padding = '10px 20px';
        button.style.border = 'none';
        button.style.borderRadius = '5px';
        button.style.cursor = 'pointer';
        button.onclick = function() {
            popup.remove();
        };

        popup.appendChild(button);
        document.body.appendChild(popup);
    """)

    # A automação pode continuar sem ser bloqueada, você pode interagir com o site normalmente
    print("O popup foi exibido. Agora você pode preencher a data manualmente.")

    # Aguarda a remoção do popup (ou até você clicar em OK)
    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element_located((By.ID, "popupAlert"))
    )

    # Após você clicar em OK no popup, a automação continua
    print("Automação retomada após a seleção da data.")

    # Após o popup sair será gerado um click no botão "Aplicar e salvar filtros"
    sleep(1)
    acao_site("//button[@class='mat-focus-indicator ng-tns-c233-107 mat-raised-button mat-button-base mat-accent']")

except Exception as e:
    print(f"Erro durante ao preencher campos de datas: {e}")

sleep(15)

try:
    expandir_painel("//div[@id='mat-select-value-97']", "mat-select-value-97")
    sleep(2)
    downloadArquivos("//mat-option[@id='mat-option-227']", 'mat-option-227')
    sleep(5)


except Exception as e:
    print(f"Erro ao fazer download: {e}")


elemento = driver.find_element(By.XPATH, '//div[@id="mat-select-value-95"]')  # Insira o XPATH correto do elemento
# Realiza o scroll até o elemento
driver.execute_script("arguments[0].scrollIntoView();", elemento)

sleep(2)

try:
    expandir_painel("//div[@id='mat-select-value-95']", "mat-select-value-95")
    sleep(2)
    downloadArquivos("//mat-option[@id='mat-option-224']", "mat-option-224")

except Exception as e:
    print(f"Erro ao expandir área de download: {e}")

sleep(3)

try:
    #downloadArquivos("//mat-option[@id='mat-option mat-focus-indicator ng-tns-c130-220']", "mat-option-221")

    expandirElemento = driver.find_element(By.ID, "mat-select-value-93")
    # Força o clique com JavaScript
    driver.execute_script("arguments[0].click();", expandirElemento)

    sleep(2)
    fazerDownload = driver.find_element(By.XPATH, "//mat-option[@value='detalhamento-msgestor']")
    fazerDownload.click()

except Exception as e:
    print(f"Erro ao expandir área de download: {e}")
finally:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # Cria uma div na página com uma mensagem para fechar o driver
    driver.execute_script("""
        var popup = document.createElement('div');
        popup.style.position = 'fixed';
        popup.style.top = '50%';
        popup.style.left = '30%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
        popup.style.color = 'white';
        popup.style.padding = '20px';
        popup.style.borderRadius = '10px';
        popup.style.fontSize = '18px';
        popup.style.zIndex = '9999';  // Definindo um z-index alto para garantir que o popup fique à frente
        popup.innerHTML = 'Quando finalizar os 3 downloads clique em OK para continuar...';

        var button = document.createElement('button');
        button.innerText = 'OK';
        button.style.marginTop = '10px';
        button.style.backgroundColor = '#4CAF50';
        button.style.color = 'white';
        button.style.padding = '10px 20px';
        button.style.border = 'none';
        button.style.borderRadius = '5px';
        button.style.cursor = 'pointer';
        button.onclick = function() {
            popup.remove();  // Remove o popup
            // Envia um evento para o Python saber que o botão foi clicado
            window.localStorage.setItem('close_browser', 'true');  // Sinaliza que o botão foi clicado
        };

        popup.appendChild(button);
        document.body.appendChild(popup);
    """)

    # Exibe a mensagem no terminal para indicar que o modal foi exibido
    print("O popup foi exibido. Clique em OK para fechar o navegador.")

    # Aguarda até que o botão OK seja clicado e a variável no localStorage seja definida
    WebDriverWait(driver, 300).until(
        lambda driver: driver.execute_script("return window.localStorage.getItem('close_browser')") == 'true'
    )

    # Após clicar em OK, fechamos o driver
    driver.quit()  # Fecha o navegador

    # A automação continua
    print("Navegador fechado, e a aplicação continuará.")

sleep(5)

# Caminho da pasta onde os arquivos Excel estão localizados
diretorio = "M:\\ADM DE VENDAS PJ\\Diario Imput\\DownloadArquivos\\arquivos baixados"

# Lista todos os arquivos da pasta
arquivos = [f for f in os.listdir(diretorio) if f.endswith('.xlsx')]

# Inicializa uma lista para armazenar os DataFrames
lista_dfs = []

# Itera sobre cada arquivo Excel encontrado
for arquivo in arquivos:
    # Lê o arquivo Excel
    caminho_arquivo = os.path.join(diretorio, arquivo)
    df = pd.read_excel(caminho_arquivo)

    # Adiciona o DataFrame à lista
    lista_dfs.append(df)

# Concatena todos os DataFrames da lista
df_final = pd.concat(lista_dfs, ignore_index=True, join="outer")

caminho_salvar = "M:\\ADM DE VENDAS PJ\\Diario Imput\\DownloadArquivos\\arquivo_unificado.xlsx"

# Salva o DataFrame final em um novo arquivo Excel no caminho especificado
df_final.to_excel(caminho_salvar, index=False)

print("Arquivos unificados com colunas diferentes!")

sleep(10)

# Exclui os arquivos originais
for arquivo in arquivos:
    caminho_arquivo = os.path.join(diretorio, arquivo)
    try:
        os.remove(caminho_arquivo)  # Apaga o arquivo
        print(f"Arquivo {arquivo} apagado com sucesso.")
    except Exception as e:
        print(f"Erro ao tentar apagar o arquivo {arquivo}: {e}")

sleep(5)

# Caminho para o arquivo unificado
arquivo_unificado = "M:\\ADM DE VENDAS PJ\\Diario Imput\\DownloadArquivos\\arquivo_unificado.xlsx"

# Caminho para o arquivo de destino (o que você vai substituir)
arquivo_destino = "M:\\ADM DE VENDAS PJ\\Diario Imput\\planilhaTeste\\DIARIO IMPUT V.37.xlsb"

# Cria uma pasta temporária única para o usuário
pasta_temp = tempfile.mkdtemp()  # Cria uma pasta temporária única
arquivo_local = os.path.join(pasta_temp, "DIARIO IMPUT V.37.xlsb")  # Caminho do arquivo temporário

# Copia o arquivo da rede para a pasta local
try:
    shutil.copy(arquivo_destino, arquivo_local)
    print(f"Arquivo copiado para a pasta temporária: {arquivo_local}")
except Exception as e:
    print(f"Erro ao copiar o arquivo da rede para local: {e}")
    exit()

# Lê o arquivo unificado
df_unificado = pd.read_excel(arquivo_unificado)

# Lê o arquivo de destino (arquivo .xlsb) usando xlwings


try:
    # Abre o arquivo .xlsb usando xlwings
    wb_destino = xw.Book(arquivo_local)

    # Acessa a aba específica pelo nome (substitua 'Nome_da_Aba' pelo nome real da aba)
    aba_destino = wb_destino.sheets['Esteira']

    # Agora você pode fazer operações nessa aba
    # Por exemplo, para ler dados de uma célula específica
    valor = aba_destino.range('A3').value
    print(f"Valor na célula A3: {valor}")

    titulos_colunas = aba_destino.range("A2:AE2").value
    df_unificado.columns = titulos_colunas

    # Define corretamente o intervalo de destino, começando da linha 3
    intervalo_final = len(df_unificado) + 2  # Começa na linha 3, então soma 2
    intervalo = f"A3:AE{intervalo_final}"

    # Substitui apenas os dados, mantendo os títulos intactos
    aba_destino.range(intervalo).value = df_unificado.values.tolist()

    # Salva e fecha o arquivo
    wb_destino.save()
    wb_destino.close()

    print("Alterações feitas com sucesso!")

except Exception as e:
    print(f"Erro ao tentar abrir o arquivo {arquivo_destino}: {e}")


# Copia o arquivo modificado de volta para a rede
try:
    # Copia o arquivo modificado de volta para a rede
    shutil.copy(arquivo_local, arquivo_destino)
    print("Arquivo copiado de volta para a rede.")
except Exception as e:
    print(f"Erro ao copiar o arquivo de volta para a rede: {e}")
    exit()

# Exclui o arquivo temporário local
try:
    os.remove(arquivo_local)
    os.rmdir(pasta_temp)
    print("Arquivo e pasta temporária removidos.")
except Exception as e:
    print(f"Erro ao limpar arquivos temporários: {e}")

# Exclui o arquivo unificado
try:
    os.remove(arquivo_unificado)  # Apaga o arquivo unificado
    print(f"Arquivo {arquivo_unificado} apagado com sucesso.")
except Exception as e:
    print(f"Erro ao tentar apagar o arquivo {arquivo_unificado}: {e}")

print("Dados substituídos com sucesso no arquivo de destino.")
