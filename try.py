import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
import re
from selenium.webdriver.support.ui import Select
import os
import pandas as pd
from bs4 import BeautifulSoup

caminho = "//Mac/Home/Downloads/alteracao_preco_calgatev1.xlsx"
code = "179658"

dict_grupos = {
    "CD" : "[CD] - MATRIZ",
    "20" : "[20] - LOJA 20 - JAU",
    "21" : "[21] - LOJA 21-PINHEIROS I",
    "22" : "[22] - LOJA 22-ITAIM BIBI",
    "26" : "[26] - LOJA 26-BARAO",
    "28" : "[28] - LOJA 28-JUNDIAI",
    "35" : "[35] - LOJA 35-PENHA",
    "40" : "[40] - LOJA 40-BRAGANCA I",
    "41" : "[41] - LOJA 41-BRAGANCA II",
    "42" : "[42] - LOJA 42-BARUERI BOUL",
    "52" : "[52] - LOJA 52 - COTIA II",
    "63" : "[63] - LOJA 63-RIB. SHOP",
    "67" : "[67] - LOJA 67-FRANC MORATO",
    "2" : "[2] - GRUPO 2",
    "3" : "[3] - GRUPO 3",
    "4" : "[4] - GRUPO 4",
    "5" : "[5] - GRUPO 5",
    "6" : "[6] - GRUPO 6"
}
num_loja = "26"
data = "25/10/2024"

# Carregar credenciais do arquivo
with open('credentials.txt', 'r') as file:
    lines = file.readlines()
    username = lines[0].strip()
    password = lines[1].strip()

navegador = webdriver.Chrome()
navegador.get("https://sumire-phd.homeip.net:8490/eVendas/home.faces")
user_name = navegador.find_element(By.ID, "form-login")
user_password = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "form-senha")))

user_name.send_keys(username)
user_password.send_keys(password)

button_login = navegador.find_element(By.ID, "form-submit")
button_login.click()

field_cadastro = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "opCadastros")))
field_cadastro.click()

field_precos = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'opPrecos')]")))
field_precos.click()

field_manutencao = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'opManutCustoProd')]")))
field_manutencao.click()

WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, "//select[@id='incCentral:incCentralDiversos:incCentralDiversos:formConteudo:selEmiCoCnpj']/option[contains(text(), '[CD] - MATRIZ')]"))
)

loja = dict_grupos.get(num_loja)
element_loja = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:selEmiCoCnpj")))
select_loja = Select(element_loja)
select_loja.select_by_visible_text(loja)

field_data = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'txtPvdDtIniValidade')]")))
# Usar JavaScript para limpar o campo
navegador.execute_script("arguments[0].value = '';", field_data)
        
# Usar JavaScript para definir o valor no campo
navegador.execute_script("arguments[0].value = arguments[1];", field_data, data)

# Disparar o evento 'change' para garantir que o valor seja registrado
navegador.execute_script("arguments[0].dispatchEvent(new Event('change'));", field_data)

def loading(navegador):

    # Espera até que o status de carregamento mude para "display: none"
    WebDriverWait(navegador, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='_viewRoot:status.start'][style='display: none']"))
    )

def box_message_nenhuma(navegador):
    try:
        # Espera até que o elemento com a classe 'divNenhumaLinha' esteja presente
        element_nenhuma_linha = navegador.find_element(By.XPATH, "//*[contains(@class, 'divNenhumaLinha')]")
        # Aqui você pode interagir com o elemento ou retornar o texto, por exemplo
        texto = element_nenhuma_linha.text
        return texto
    except:
        return None

def box_message_td(navegador):
    try:
        # Tenta capturar o texto no primeiro XPath
        element_tds = navegador.find_elements(By.XPATH, "//*[@id='incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody']//td[@class='tblLinha']")
        if element_tds:
            descricao = element_tds[0].text
            texto = str(descricao.split(' - ')[0])
            return texto

        # Se não encontrar o primeiro elemento, tenta capturar o segundo XPath
        element_tds = navegador.find_elements(By.XPATH, "//*[@id='incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody']//td[@class='tblLinha']")
        if element_tds:
            descricao = element_tds[0].text
            texto = str(descricao.split(' - ')[0])
            return texto
    except:
        return None

def get_value(navegador, id):
        # Obtendo o conteúdo HTML da página com Selenium
    html_content = navegador.page_source

    # Analisando o HTML com BeautifulSoup e formatando
    soup = BeautifulSoup(html_content, 'html.parser')
    pretty_html = soup.prettify()

    # Definindo o caminho do arquivo para salvar o HTML no diretório do projeto
    file_path = os.path.join(os.getcwd(), f"pagina_formatada_{id}.html")

    # Salvando o HTML formatado em um arquivo
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(pretty_html)

    try:
        value = navegador.execute_script("return document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:txtPrdCoProdutoFiltro').value;")
        return str(value)
    except:
        return None

def capturar_codigo_span(navegador):
    try:
        # Localiza o elemento span pelo texto exato
        span_element = navegador.find_element(By.XPATH, "//span[contains(text(), 'Produto pertence a um grupo de preços. Não permite alterar!')]")

        # Obtém o texto do elemento
        span_text = span_element.text

        # Usa uma expressão regular para capturar o número entre parênteses
        match = re.search(r'\((\d+)\)', span_text)
        
        if match:
            codigo = match.group(1)  # Extrai o número capturado
            return codigo
        else:
            return None
    except:
        return None

loading(navegador)
# Clica no botão para selecionar o produto
button_grupo_preco = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Selecionar o Produto']")))
button_grupo_preco.click()


field_code_grupe = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:txtPrdCoProdutoFiltro")))
field_code_grupe.click()
field_code_grupe.clear()
field_code_grupe.send_keys(code)

button_pesquisar = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:btnPsqProduto")))
button_pesquisar.click()

time.sleep(2)

# Espera até que o elemento seja clicável incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:0:j_id278
pastel_all_codes= WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")))
# Aqui estamos buscando os mesmos elementos que o 'pastel_all_codes' se refere
pastel_codes = pastel_all_codes.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")
pastel_codes[1].click()

time.sleep(2)
# Clica no botão para selecionar o produto
button_grupo_preco = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Selecionar o Produto']")))
button_grupo_preco.click()

print(get_value(navegador,1))
time.sleep(5)
print(get_value(navegador,2))



time.sleep(5)