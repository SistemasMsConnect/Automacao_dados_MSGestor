"""
** Necessário instalar o selenium
** Necessário instalar o pandas
** Necessário instalar o xlwings
** Necessário instalar o pyopenxl
** Caso queria criar um executável pode instalar o pyinstaller
e então executar o pyinstaller nome_da_aplicação
"""

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
import datetime


chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Para abrir o navegador maximizado

#Solicitando caminho para download do arquivo do site.
download_dir = input(r"Insira o caminho onde sera feito o download dos 3 arquivos: ")
print('')
print('Caminho inserido com sucesso! Por favor aguarde!')

#Realizando pré-configurações no chrome
prefs = {
    "profile.default_content_settings.popups": 0,  # Bloqueia popups de confirmação
    "download.default_directory": download_dir,  # Diretório de
    "download.prompt_for_download": False,  # Desativa o prompt de confirmação de
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True  # Habilita a navegação segura (impede mensagens de segurança)
}

chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

#função para diminuir o código
def acao_site(elemento):

    #codigo para expandir o campo para buscar data
    campo_expandir = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, elemento))
    )
    campo_expandir.click()

#função para expandir o painel de cada download do site
def expandir_painel(xpath, path):

    #encontra o painel de expansão de download e gera um click para expandir
    campo_download = driver.find_element(By.XPATH, xpath)

    # Aguarda o campo de data e abre o calendário
    campo_download = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, path))
    )
    campo_download.click()

#Função para realizar download de arquivos do site
def downloadArquivos(xpath, path):
    download_1 = driver.find_element(By.XPATH, xpath)
    download_1 = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.ID, path))
    )
    download_1.click()
    sleep(1)

# Função para aguardar o término de cada download
padrao_nome = "relatorio_msgestor_detalhamento_itens-exportacao"
def esperar_download():
    arquivos_antes = set(os.listdir(download_dir))  # Lista antes do download
    while True:
        sleep(1)  # Espera um segundo para não sobrecarregar a CPU
        arquivos_depois = set(os.listdir(download_dir))
        setNnovos_arquivos = arquivos_depois - arquivos_antes  # Identifica novos arquivos

        # Verifica se há um novo arquivo e se ele terminou de baixar
        for setArquivo in setNnovos_arquivos:
            if setArquivo.startswith(padrao_nome) and not setArquivo.endswith(".crdownload"):
                print(f"Download concluído: {setArquivo}")
                return os.path.join(download_dir, setArquivo)


#Faz a conexao do python com o site especifico
try:
    #localiza a endereço e entrar no site MSGestor
    driver.get("https://msgestor.msconnect.com.br/pages/auth/login")
    sleep(2)

    #Localiza os campos de usuário e senha
    campo_usuario = driver.find_element(By.XPATH,"//input[@id='mat-input-26']")
    campo_senha = driver.find_element(By.XPATH,"//input[@id='mat-input-27']")
    sleep(1)
    #Preencge os campos de login usuário e senha
    campo_usuario.send_keys("######") # Insira o login do usuario aqui
    campo_senha.send_keys("######") # -Insira a senha do usuario aqui
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
    print('')
    print("O popup foi exibido. Agora você pode preencher a data manualmente.")
    print('')

    # Aguarda a remoção do popup (ou até você clicar em OK)
    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element_located((By.ID, "popupAlert"))
    )

    # Após você clicar em OK no popup, a automação continua
    print('')
    print("Automação retomada após a seleção da data.")
    print('')

    # Após o popup sair será gerado um click no botão "Aplicar e salvar filtros"
    sleep(1)
    acao_site("//button[@class='mat-focus-indicator ng-tns-c233-107 mat-raised-button mat-button-base mat-accent']")
except Exception as e:
    print(f"Erro durante ao preencher campos de datas: {e}")

sleep(2)
sleep(18)

#Utilizando a função para expandir o primeiro painel de download e realizar o download
try:
    expandir_painel("//div[@id='mat-select-value-97']", "mat-select-value-97")
    sleep(2)
    downloadArquivos("//mat-option[@id='mat-option-227']", 'mat-option-227')
    print("O próximo download será iniciado assim que o primeiro for concluído.")
    esperar_download()
except Exception as e:
    print(f"Erro ao fazer download: {e}")



elemento = driver.find_element(By.XPATH, '//div[@id="mat-select-value-95"]')  # Insira o XPATH correto do elemento
# Realiza o scroll até o elemento
driver.execute_script("arguments[0].scrollIntoView();", elemento)


#Utilizando a função para expandir o segundo painel de download e realizar o download
try:
    expandir_painel("//div[@id='mat-select-value-95']", "mat-select-value-95")
    downloadArquivos("//mat-option[@id='mat-option-224']", "mat-option-224")
    print("O terceiro download será iniciado assim que o segundo for concluído.")
    esperar_download()
except Exception as e:
    print(f"Erro ao expandir área de download: {e}")



#Força o JavaScript clicar no botão, pois o metodo usado em outros estava dando erro haha
try:
    expandirElemento = driver.find_element(By.ID, "mat-select-value-93")
    # Força o clique com JavaScript
    driver.execute_script("arguments[0].click();", expandirElemento)
    sleep(2)
    fazerDownload = driver.find_element(By.XPATH, "//mat-option[@value='detalhamento-msgestor']")
    fazerDownload.click()
    print("Realizando download do terceiro arquivo.")
    esperar_download()
except Exception as e:
    print(f"Erro ao expandir área de download: {e}")


print("Todos os downloads foram concluídos!")
sleep(3)
driver.quit()  # Fecha o navegador

# Caminho da pasta onde os arquivos Excel estão localizados

# Lista todos os arquivos da pasta
arquivos = [f for f in os.listdir(download_dir) if f.endswith('.xlsx')]

# Inicializa uma lista para armazenar os DataFrames
lista_dfs = []

# Itera sobre cada arquivo Excel encontrado
for arquivo in arquivos:
    # Lê o arquivo Excel
    caminho_arquivo = os.path.join(download_dir, arquivo)
    df = pd.read_excel(caminho_arquivo)

    # Adiciona o DataFrame à lista
    lista_dfs.append(df)

# Concatena todos os DataFrames da lista
df_final = pd.concat(lista_dfs, ignore_index=True, join="outer")


nome_arquivo = "arquivo_consolidado.xlsx"
print('')
caminho = input(r"Insira aqui o caminho para salvar o arquivo consolidado: ")
print('')
print('Caminho inserido com sucesso! Por favor aguarde!')
print('')

caminho_salvar = caminho+"\\"+nome_arquivo

# Salva o DataFrame final em um novo arquivo Excel no caminho especificado
df_final.to_excel(caminho_salvar, index=False)

print('')
print("Arquivos consolidados com sucesso!!")
print('')
sleep(5)

# Exclui os arquivos originais
for arquivo in arquivos:
    caminho_arquivo = os.path.join(download_dir, arquivo)
    try:
        os.remove(caminho_arquivo)  # Apaga o arquivo
        print('')
        print(f"Arquivo {arquivo} apagado com sucesso.")
        print('')
    except Exception as e:
        print(f"Erro ao tentar apagar o arquivo {arquivo}: {e}")

sleep(5)

# Caminho para o arquivo unificado
arquivo_unificado = caminho_salvar

# Caminho para o arquivo de destino (o que você vai substituir)
print('')
nome_arquivo_diario = "DIARIO IMPUT V.37.xlsb"
destino_arquivo = input(rf"Digite o caminho do arquivo DIARIO IMPUT: "  )
print('')

arquivo_destino = destino_arquivo+"\\"+nome_arquivo_diario
print('')
print('Caminho inserido com sucesso! Por favor aguarde!')
print('')
#arquivo_destino = "M:\\ADM DE VENDAS PJ\\Diario Imput\\planilhaTeste\\DIARIO IMPUT V.37.xlsb"

# Cria uma pasta temporária única para o
pasta_temp = tempfile.mkdtemp()  # Cria uma pasta temporária única
arquivo_local = os.path.join(pasta_temp, "DIARIO IMPUT V.37.xlsb")  # Caminho do arquivo temporário

# Copia o arquivo da rede para a pasta local
try:
    shutil.copy(arquivo_destino, arquivo_local)
    print('')
    print(f"Arquivo copiado para a pasta temporária: {arquivo_local}")
    print('')
except Exception as e:
    print(f"Erro ao copiar o arquivo da rede para local: {e}")
    exit()

# Lê o arquivo unificado
df_unificado = pd.read_excel(arquivo_unificado)

# Lê o arquivo de destino (arquivo .xlsb) usando xlwings
# Abre o arquivo .xlsb usando xlwings

wb_destino = xw.Book(arquivo_local)

# Acessa a aba específica pelo nome (substitua 'Nome_da_Aba' pelo nome real da aba)
aba_destino = wb_destino.sheets['Esteira']
try:
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



    print('')
    print("Alterações feitas com sucesso!")
    print('')
except Exception as e:
    print(f"Erro ao tentar abrir o arquivo {arquivo_destino}: {e}")

print("Iniciando a manipulação dos dados da planilha!")
print(" ")

#Aqui esta iniciando a manipulação de dados da planilha
try:
    # Abre o arquivo local
    wb_destino = xw.Book(arquivo_local)
    aba_destino = wb_destino.sheets['Esteira']
    print("Arquivo aberto com sucesso")

    # Desativa atualizações e cálculos
    wb_destino.app.screen_updating = False
    wb_destino.app.calculation = 'manual'
    print("Configurações do Excel ajustadas")

    # Passo 1: Formatar a coluna B com Pandas e encontrar a última linha com dados
    # Busca a última linha com dados na coluna B
    ultima_linha_real = aba_destino.range("B" + str(aba_destino.cells.last_cell.row)).end('up').row
    if ultima_linha_real < 3:
        ultima_linha_real = 3  # Garante que comece em B3
    print(f"Última linha com dados detectada na coluna B: {ultima_linha_real}")

    coluna_b = aba_destino.range(f"B3:B{ultima_linha_real}").value
    df_coluna_b = pd.Series(coluna_b)
    df_coluna_b = df_coluna_b.apply(
        lambda x: x.split(" ")[0] if isinstance(x, str) and " " in x
        else (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=x)).strftime("%d/%m/%Y") if isinstance(x, (int, float))
        else x
    )
    aba_destino.range(f"B3:B{ultima_linha_real}").value = [[v] for v in df_coluna_b]
    print("Passo 1 concluído: Coluna B formatada")

    # Passo 2: Remover dados de AF4:AS apenas até a última linha real
    aba_destino.range(f"AF4:AS{ultima_linha_real}").clear_contents()
    print("Passo 2 concluído: Dados de AF4:AS removidos")

    # Passo 3 e 5: Expandir fórmulas e colar valores em lotes menores
    formulas = aba_destino.range("AF3:AS3").formula
    print(f"Fórmulas a serem aplicadas: {formulas}")
    lote_tamanho = 50000 #em um computador mais potente pode aumentar esse valor (em pc menos potente, diminui esse valor)
    inicio = 4
    while inicio <= ultima_linha_real:
        fim = min(inicio + lote_tamanho - 1, ultima_linha_real)
        print(f"Aplicando fórmulas em AF{inicio}:AS{fim}")
        aba_destino.range(f"AF{inicio}:AS{fim}").formula = formulas
        print(f"Calculando lote AF{inicio}:AS{fim}")
        wb_destino.app.calculate()
        valores_calculados = aba_destino.range(f"AF{inicio}:AS{fim}").value
        aba_destino.range(f"AF{inicio}:AS{fim}").value = valores_calculados
        print(f"Lote processado: AF{inicio}:AS{fim}")
        #sleep(0.9)  # Pausa para liberar memória #Computador menos potente descomentar esse linha sleep(0.9)
        inicio = fim + 1
    print("Passo 3 e 5 concluídos: Fórmulas expandidas e valores colados")

    # Reativa atualizações e cálculos
    wb_destino.app.screen_updating = True
    wb_destino.app.calculation = 'automatic'
    print("Configurações do Excel restauradas")

    # Salva e fecha o arquivo local
    wb_destino.save()
    wb_destino.close()
    print("Arquivo local salvo e fechado")

    # Copia o arquivo de volta para a rede
    try:
        shutil.copy(arquivo_local, arquivo_destino)
        print(f"Arquivo copiado de volta para a rede: {arquivo_destino}")
    except Exception as e:
        print(f"Erro ao copiar o arquivo de volta para a rede: {e}")

    print("Manipulações realizadas com sucesso no arquivo local!")

except Exception as e:
    print(f"Erro ao processar a planilha: {e}")
    if 'wb_destino' in locals():
        wb_destino.close()

# Exclui o arquivo temporário local
try:
    os.remove(arquivo_local)
    os.rmdir(pasta_temp)
    print('')
    print("Arquivo e pasta temporária removidos.")
    print('')
except Exception as e:
    print(f"Erro ao limpar arquivos temporários: {e}")

# Exclui o arquivo unificado
try:
    os.remove(arquivo_unificado)  # Apaga o arquivo unificado
    print('')
    print(f"Arquivo {arquivo_unificado} apagado com sucesso.")
    print('')
except Exception as e:
    print(f"Erro ao tentar apagar o arquivo {arquivo_unificado}: {e}")
print('')
print("Dados substituídos com sucesso no arquivo de destino.")
