from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
import re
import os


# Dicionário para mapear lojas fantasias para seus IDs e CNPJs
dict_lojas = {
    "AVARE CD" : "2183783000111",
    "SOROCABA" : "2183783001606",
    "SOROCABA-MP" : "2183783001517",
    "LOJA 01 - AVARE" : "2183783000545",
    "LOJA 03 - MARILIA" : "2183783000383",
    "LOJA 04 - ARACATUBA" : "2183783000200",
    "LOJA 05 - PRUDENTE" : "2183783000464",
    "LOJA 07 - MARILIA" : "2183783000626",
    "LOJA 08 - BAURU" : "2183783000898",
    "LOJA 09 - ASSIS" : "2183783000707",
    "LOJA 10 - ANDRADINA" : "2183783000979",
    "LOJA 11 - BIRIGUI" : "2183783001193",
    "LOJA 12 - ITAPEVA" : "2183783001002",
    "LOJA 13 - BAURU SHOP" : "2183783001274",
    "LOJA 14 - OURINHOS" : "2183783001355",
    "LOJA 17 - JUNDIAI" : "14298644000112",
    "LOJA 20 - JAU" : "328944000354",
    "LOJA 21 - PINHEIROS" : "8571461000126",
    "LOJA 22 - ITAIM BIBI" : "8571461000207",
    "LOJA 25 - STA CRUZ" : "2183783001860",
    "LOJA 26 - BARAO" : "2183783001789",
    "LOJA 28 - JUNDIAI" : "2183783001940",
    "LOJA 30 - ARICANDUVA" : "2183783002165",
    "LOJA 31 - MATEO BEI" : "2183783002246",
    "LOJA 32 - TEODORO" : "2183783002327",
    "LOJA 33 - ITAIM" : "2183783002408",
    "LOJA 34 - HIGIENOPOLIS" : "2183783002599",
    "LOJA 40 - BRAGANCA I" : "72714637000150",
    "LOJA 41 - BRAGANCA II" : "72714637000401",
    "LOJA 42 - BARUERI BOUL" : "2183783003218",
    "LOJA 43 - BARUERI CAMP" : "2183783003307",
    "LOJA 44 - PERUS" : "2183783003480",
    "LOJA 45 - CRUZEIRO" : "2183783003641",
    "LOJA 46 - GUARA CENTRO" : "2183783003722",
    "LOJA 47 - GUARA SHOP" : "2183783003803",
    "LOJA 48 - LORENA" : "2183783003994",
    "LOJA 49 - PINDA" : "2183783004028",
    "LOJA 50 - TAUBATE I" : "2183783004109",
    "LOJA 51 - TAUBATE II" : "2183783004290",
    "LOJA 55 - BOTUCATU 1" : "2183783004613",
    "LOJA 56 - BOTUCATU 2" : "2183783004702",
    "LOJA 57 - JAU" : "2183783004885",
    "LOJA 58 - SOROCABA 1" : "2183783002670",
    "LOJA 61 - SC V PRADO" : "2183783002912",
    "LOJA 62 - SC SHOPPING" : "2183783003056",
    "LOJA 63 - RIB SHOP" : "2183783003137",
    "LOJA 64 - IPIRANGA" : "2183783004966",
    "LOJA 66 - PIEDADE" : "2183783005180",
    "LOJA 67 - FRANC MORATO" : "2183783005261",
    "LOJA 68 - MARILIA III" : "2183783005423",
    "LOJA 71 - P FERREIRA" : "2183783005695",
}
def open_file(caminho):
    df = pd.read_excel(caminho, engine='openpyxl')
    return df

# Função para salvar o nome de usuário e a senha em um arquivo
def save_credentials():
    username = username_entry.get()
    password = password_entry.get()
    with open('credentials.txt', 'w') as file:
        file.write(f"{username}\n{password}")

# Função para carregar o nome de usuário e a senha do arquivo
def load_credentials():
    try:
        with open('credentials.txt', 'r') as file:
            lines = file.readlines()
            username = lines[0].strip()
            password = lines[1].strip()
            username_entry.insert(0, username)
            password_entry.insert(0, password)
    except FileNotFoundError:
        # Cria o arquivo se não existir
        with open('credentials.txt', 'w') as file:
            file.write("")  # Cria um arquivo vazio

# Função para realizar o login e abrir a página principal
def login():
    save_credentials()  # Salva as credenciais
    root.destroy()  # Fecha a janela de login
    open_main_page()  # Abre a página principal

def abrir_arquivo(entrada_arquivo, root):
    arquivo_selecionado = filedialog.askopenfilename(parent=root)
    if arquivo_selecionado:
        entrada_arquivo.delete(0, tk.END)  # Limpa a entrada
        entrada_arquivo.insert(0, arquivo_selecionado)  # Insere o caminho do arquivo selecionado

# Função para abrir a página principal
def open_main_page():
    # Criação da Página Principal
    main_page = Tk()
    main_page.title('Página Principal')
    main_page.configure(background='DeepPink2')
    main_page.geometry("500x200")
    main_page.resizable(True, True)
    main_page.minsize(width=500, height=200)

    frame1 = Frame(main_page)
    frame1.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

    buttom_consulta_preco = Button(frame1, text="Consulta Preço", bd=3, command=lambda: page_search_price(main_page, dict_lojas))
    buttom_consulta_preco.place(relx=0.2, rely=0.3, relwidth=0.2, relheight=0.3)
    buttom_mov_interna = Button(frame1, text="Mov. Interna", bd=3, command=lambda: page_mov_int(main_page, dict_lojas))
    buttom_mov_interna.place(relx=0.4, rely=0.3, relwidth=0.2, relheight=0.3)
    buttom_novo = Button(frame1, text="Alteraçao Preço", bd=3)
    buttom_novo.place(relx=0.6, rely=0.3, relwidth=0.2, relheight=0.3)
    # buttom_alterar = Button(frame1, text="Alterar", bd=3)
    # buttom_alterar.place(relx=0.65, rely=0.1, relwidth=0.1, relheight=0.3)
    # buttom_apagar = Button(frame1, text="Apagar", bd=3)
    # buttom_apagar.place(relx=0.75, rely=0.1, relwidth=0.1, relheight=0.3)
    main_page.mainloop()

# Criação da Página de Movimentação Interna
def page_mov_int(main_page, dict_lojas):
    root1 = Toplevel(main_page)
    root1.title("Movimentação Interna")
    root1.configure(background='DeepPink2')
    root1.geometry("500x200")
    root1.resizable(True, True)
    root1.minsize(width=500, height=200)

    # Pegar a posição da main_page
    x_main_page = main_page.winfo_x()
    y_main_page = main_page.winfo_y()

    # Definir root1 para ser aberta no mesmo local de main_page
    root1.geometry(f"+{x_main_page}+{y_main_page}")

    # Tornar root1 uma janela "filha" de main_page
    root1.transient(main_page)
    root1.lift()
    root1.grab_set()

    # Frame
    frame3 = Frame(root1)
    frame3.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

    # Inclusão do arquivo de entrada
    folder_name = Label(frame3, text="Arquivo")
    folder_name.place(relx=0.0, rely=0.1, relwidth=0.3, relheight=0.1)
    folder = Entry(frame3)
    folder.place(relx=0.30, rely=0.1, relwidth=0.5, relheight=0.1)
    buttom_search_folder = tk.Button(frame3, text='...', command=lambda: abrir_arquivo(folder, root1))
    buttom_search_folder.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.1)

    # Inclusão Combo Box Lojas Origem
    loja_origem_label = Label(frame3, text="Loja Origem")
    loja_origem_label.place(relx=0.0, rely=0.35, relwidth=0.25, relheight=0.1)
    loja_origem = ttk.Combobox(frame3, values=list(dict_lojas.keys()))
    loja_origem.place(relx=0.3, rely=0.35, relwidth=0.6, relheight=0.1)

    # Inclusão Combo Box Lojas Destino
    loja_destino_label = Label(frame3, text="Loja Destino")
    loja_destino_label.place(relx=0.0, rely=0.6, relwidth=0.25, relheight=0.1)
    loja_destino = ttk.Combobox(frame3, values=list(dict_lojas.keys()))
    loja_destino.place(relx=0.3, rely=0.6, relwidth=0.6, relheight=0.1)

    # Inclusão do botão Iniciar
    buttom_start = Button(frame3, text="Iniciar", bd=3, command=lambda: movimentacao_interna(folder.get(), dict_lojas[loja_origem.get()], dict_lojas[loja_destino.get()]))
    buttom_start.place(relx=0.7, rely=0.75, relwidth=0.2, relheight=0.15)


# Funcao para consultar o preco.
def consulta_preco(caminho, select_loja):
    import warnings
    warnings.simplefilter(action='ignore', category=FutureWarning)
    
    # Obter o CNPJ da loja selecionada
    cnpj_loja = dict_lojas.get(select_loja)
    
    def novo_caminho(caminho):
        # Extrair diretório e nome base do arquivo
        dir_name, file_name = os.path.split(caminho)
        base_name, ext = os.path.splitext(file_name)

        # Criar o novo nome de arquivo
        new_file_name = f"{base_name}_{select_loja}_preco_coletado_parcial{ext}"
        new_file_path = os.path.join(dir_name, new_file_name)

        return new_file_path

    def arquivo_final(caminho):
        # Definir o caminho do arquivo existente (parcial)
        old_file_path = novo_caminho(caminho)

        # Extrair diretório e nome base do arquivo
        dir_name, file_name = os.path.split(old_file_path)
        base_name, ext = os.path.splitext(file_name)

        # Substituir a palavra 'parcial' por 'final' no nome do arquivo
        new_file_name = file_name.replace('parcial', 'final')
        new_file_path = os.path.join(dir_name, new_file_name)
        
        # Renomear o arquivo
        os.rename(old_file_path, new_file_path)

    def abrir_arquivo_existente(caminho):
        new_file_path = novo_caminho(caminho)

        if os.path.exists(new_file_path):
            df_existente = pd.read_excel(new_file_path, dtype={'Product Code': str}, engine='openpyxl')
            return df_existente
        else:
            # Retorna um DataFrame vazio com as colunas especificadas
            colunas = ['Product Code', 'Product Description', 'Price Cust Rep', 'Price Venda',
                    'Price Promocao', 'Price Custo Cont', 'Price Fidelidade', 'Price Ecommerce']
            return pd.DataFrame(columns=colunas)  # Retorna DataFrame vazio se o arquivo não existir
    
    def enter_navegador():
        # Carregar credenciais do arquivo
        try:
            with open('credentials.txt', 'r') as file:
                lines = file.readlines()
                username = lines[0].strip()
                password = lines[1].strip()
        except:
            pass

        navegador = webdriver.Chrome()
        navegador.get("https://sumire-phd.homeip.net:8490/SistemasPHD/")
        user_name = navegador.find_element(By.ID, "form-login")
        user_password = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.ID, "form-senha"))
        )

        user_name.send_keys(username)
        user_password.send_keys(password)

        button_login = navegador.find_element(By.ID, "form-submit")
        button_login.click()

        time.sleep(2)
        navegador.find_element(By.ID, 'j_id13').click()

        time.sleep(2)
        navegador.find_element(By.ID, 'opConsultas').click()

        time.sleep(2)
        navegador.find_element(By.XPATH, "//input[starts-with(@id, 'incCentral:')]").click()

        cnpj_field = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:incCentralConsultas:formEscolheCnpj:selEmiCoCnpj'))
        )
        cnpj_field.click()

        select_cnpj_loja = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f".//option[contains(@value, '{cnpj_loja}')]"))
        )
        select_cnpj_loja.click()

        return navegador

    def collect_prices(navegador, code, descricao=None):
        # Função para localizar o valor baseado no cabeçalho, pulando para a próxima linha
        def buscar_valor_proxima_linha(soup, cabecalho):
            try:
                # Encontra a <td> que contém o cabeçalho
                td_cabecalho = soup.find('td', string=cabecalho)
                tr_cabecalho = td_cabecalho.find_parent('tr')
                tds_cabecalho = tr_cabecalho.find_all('td')

                # Encontra a próxima <tr> (a linha com os valores)
                proxima_tr = tr_cabecalho.find_next('tr')

                # Captura todas as <td> da próxima linha (a linha de valores)
                tds_valores = proxima_tr.find_all('td')

                # Descobre o índice do cabeçalho desejado (ex: 'Venda')
                for i, td in enumerate(tds_cabecalho):
                    if cabecalho.strip() == td.get_text(strip=True):
                        return tds_valores[i].text.strip()
                return None
            except:
                return None

        # Função para localizar o valor do texto dentro do 'span' baseado no cabeçalho
        def localizar_valor_por_cabecalho(soup, cabecalho):
            try:
                # Procura a linha com o cabeçalho desejado
                linha = soup.find('td', string=cabecalho)             
                # Localiza o próximo 'td'
                valor_td = linha.find_next('td')
                # Busca o valor dentro do 'span'
                valor_span = valor_td.find('span')
                if valor_span:
                    return valor_span.get_text(strip=True)  # Retorna o valor visível do 'span'
                else:
                    return None
            except:
                return None
            
        if descricao is None:
            # Remover qualquer coisa que não seja número do code, se for necessário
            code = re.sub(r'\D', '', str(code).strip())  # Converte para string e mantém apenas números
            
            WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.ID, 'incCentral:incCentralConsultas:pnlObjQtdeValores'))
            )
            
            check_code = False
            while not check_code:
                html_content = navegador.page_source

                # Criar um objeto BeautifulSoup
                soup = BeautifulSoup(html_content, 'html.parser')
                try:
                    product_code = soup.find('td', class_='colCodigoProduto cinza1').text
                except:
                    product_code = 0
                if product_code.strip() == code:
                    check_code = True
                    break
                else:
                    pass

            # Localize os valores usando o texto dos cabeçalhos
            product_description = soup.find('td', class_='colDescricaoProduto cinza1').text
            price_cust_rep = buscar_valor_proxima_linha(soup, 'Custo Rep.')
            price_venda = buscar_valor_proxima_linha(soup, 'Venda')
            price_promocao = buscar_valor_proxima_linha(soup, 'Promoção')
            price_custo_cont = buscar_valor_proxima_linha(soup, 'Custo Cont.')
            price_fidelidade = localizar_valor_por_cabecalho(soup, 'FIDELIDADE')
            price_ecommerce = localizar_valor_por_cabecalho(soup, 'e-Commerce')

            # Função para lidar com valores None e fazer a substituição ou retornar um valor padrão
            def safe_float_conversion(value):
                if value is None:
                    return 'SEM VALOR'
                return float(value.replace(',', '.'))

            # Criação de um novo DataFrame com os dados extraídos
            new_data = {
                'Product Code': [product_code],
                'Product Description': [product_description],
                'Price Cust Rep': safe_float_conversion(price_cust_rep),
                'Price Venda': safe_float_conversion(price_venda),
                'Price Promocao': safe_float_conversion(price_promocao),
                'Price Custo Cont': safe_float_conversion(price_custo_cont),
                'Price Fidelidade': safe_float_conversion(price_fidelidade),
                'Price Ecommerce': safe_float_conversion(price_ecommerce)
            }
        else:
            new_data = {
                'Product Code': [code],
                'Product Description': [descricao],
                'Price Cust Rep': [''],
                'Price Venda': [''],
                'Price Promocao': [''],
                'Price Custo Cont': [''],
                'Price Fidelidade': [''],
                'Price Ecommerce': ['']
            }
        new_df = pd.DataFrame(new_data)

        df_frame = abrir_arquivo_existente(caminho)

        # Concatenar o novo DataFrame com o DataFrame existente
        df_frame = pd.concat([df_frame, new_df], ignore_index=True)

        new_file_path = novo_caminho(caminho)

        # Salvar o DataFrame atualizado no arquivo
        df_frame.to_excel(new_file_path, index=False)

    def selecionar_item(navegador, code):
        select_item = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//input[starts-with(@id, 'incCentral:incCentralConsultas:')]"))
        )
        select_item.click()

        try:
            # Criar um objeto BeautifulSoup
            html_content = navegador.page_source        
            soup = BeautifulSoup(html_content, 'html.parser')

            # Procurar a frase "Nenhuma linha retornada!"
            mensagem = soup.find(string=re.compile("Nenhuma linha retornada!", re.IGNORECASE))

            code_field = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.ID, 'incCentral:incCentralConsultas:formPnlModalPesquisa:txtPrdCoProdutoFiltro'))
            )
            code_field.click()
            code_field.clear()
            code_field.send_keys(code)

            consulta = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.ID, 'incCentral:incCentralConsultas:formPnlModalPesquisa:btnPsqProduto'))
            )
            consulta.click()
            time.sleep(6)
        except:
            code_field = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.ID, 'incCentral:incCentralConsultas:formPnlModalPesquisa:txtPrdCoProdutoFiltro'))
            )
            code_field.click()
            code_field.clear()
            code_field.send_keys(code)

            consulta = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.ID, 'incCentral:incCentralConsultas:formPnlModalPesquisa:btnPsqProduto'))
            )
            consulta.click()

        check_code = False
        
        # Remover qualquer coisa que não seja número do code, se for necessário
        code = re.sub(r'\D', '', str(code).strip())  # Converte para string e mantém apenas números
        while not check_code:
            tentativas =+ 1
            # Criar um objeto BeautifulSoup
            html_content = navegador.page_source        
            soup = BeautifulSoup(html_content, 'html.parser')

            # Procurar a frase "Nenhuma linha retornada!"
            mensagem = soup.find(string=re.compile("Nenhuma linha retornada!", re.IGNORECASE))

            if mensagem:
                navegador.execute_script("document.getElementById('incCentral:incCentralConsultas:pnlPsqProduto').component.hide()")
                return False
            try:
                linhas = soup.find_all('td', class_='tblLinha')
                td_linha = linhas[0] if linhas else check_code == False

                # Pegar o texto da <td> e dividir para pegar apenas o primeiro valor (número)
                texto_completo = td_linha.get_text(strip=True)

                # Extrair apenas números do code_box (caso tenha texto misturado)
                code_box = texto_completo.split(' ')[0]
                code_box = re.sub(r'\D', '', code_box.strip())  # Mantém apenas números
            except:
                code_box = 0
            # Verificar se ambos são iguais
            if code_box == code:
                check_code = True
                break
            else:
                pass

        item = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[starts-with(@id, 'incCentral:incCentralConsultas:formPnlModalPesquisa:tblPsqInfoParticipanteBody:')]"))
            )
        item.click()
        return True

    # Carregar os dados da planilha original e garantir que os códigos de produto sejam strings
    lista = pd.read_excel(caminho, dtype={0: str}, engine='openpyxl')

    # Verificar se o arquivo já existe
    df_existente = abrir_arquivo_existente(caminho)

    # Remover espaços em branco ou caracteres extras dos códigos de df_existente
    df_existente['Product Code'] = df_existente['Product Code'].astype(str).str.strip()

    # Remover espaços em branco ou caracteres extras dos códigos de lista (primeira coluna)
    codigos_novos = lista.iloc[:, 0].astype(str).str.strip()

    # Comparar a primeira coluna de 'lista' com 'Product Code' de df_existente
    codigos_existentes = df_existente['Product Code'].unique()

    # Filtrar apenas os códigos que não estão no arquivo existente
    lista = lista[~codigos_novos.isin(codigos_existentes)]

    # Inicializar o navegador e coletar dados apenas para os códigos não processados
    if not lista.empty:
        nav = enter_navegador()
        for index, row in lista.iterrows():
            code = row.iloc[0]
            codigo_localizado = selecionar_item(nav, code)
            if codigo_localizado:
                collect_prices(nav, code)
            else:
                collect_prices(nav, code, descricao='NAO LOCALIZADO')
    arquivo_final(caminho)

# Criação da Página de Consulta Preco
def page_search_price(main_page, dict_lojas):
    root1 = Toplevel(main_page)
    root1.title("Consulta de Precos")
    root1.configure(background='DeepPink2')
    root1.geometry("500x200")
    root1.resizable(True, True)
    root1.minsize(width=500, height=200)

    # Pegar a posição da main_page
    x_main_page = main_page.winfo_x()
    y_main_page = main_page.winfo_y()

    # Definir root1 para ser aberta no mesmo local de main_page
    root1.geometry(f"+{x_main_page}+{y_main_page}")

    # Tornar root1 uma janela "filha" de main_page
    root1.transient(main_page)
    root1.lift()
    root1.grab_set()

    # Frame
    frame3 = Frame(root1)
    frame3.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

    # Inclusão do arquivo de entrada
    folder_name = Label(frame3, text="Arquivo")
    folder_name.place(relx=0.0, rely=0.1, relwidth=0.3, relheight=0.1)
    folder = Entry(frame3)
    folder.place(relx=0.30, rely=0.1, relwidth=0.6, relheight=0.1)
    buttom_search_folder = tk.Button(frame3, text='...', command=lambda: abrir_arquivo(folder, root1))
    buttom_search_folder.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.1)

    # Inclusão Combo Box Selecionar Loja
    select_loja_label = Label(frame3, text="Selecionar Loja")
    select_loja_label.place(relx=0.0, rely=0.35, relwidth=0.25, relheight=0.1)
    select_loja = ttk.Combobox(frame3, values=list(dict_lojas.keys()))
    select_loja.place(relx=0.3, rely=0.35, relwidth=0.6, relheight=0.1)

    # Inclusão do botão Iniciar
    buttom_start = Button(frame3, text="Iniciar", bd=3, command=lambda: consulta_preco(folder.get(), select_loja.get()))
    buttom_start.place(relx=0.7, rely=0.75, relwidth=0.2, relheight=0.15)

# Funcao para realizar a movimentacao interna
def movimentacao_interna(caminho, cnpj_origem, cnpj_destino):
    def enter_navegador():
        # Carregar credenciais do arquivo
        try:
            with open('credentials.txt', 'r') as file:
                lines = file.readlines()
                username = lines[0].strip()
                password = lines[1].strip()
        except:
            pass
        
        # Entrar no navegador
        navegador = webdriver.Chrome()
        navegador.get("https://sumire-phd.homeip.net:8490/eVendas/home.faces")
        user_name = navegador.find_element(By.ID, "form-login")
        user_password = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.ID, "form-senha"))
        )

        # Credenciais
        user_name.send_keys(username)
        user_password.send_keys(password)

        button_login = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, "form-submit"))
        )
        button_login.click()

        return navegador

    def enter_mov_int(navegador):
        navegador.find_element(By.ID, 'opMovInterna').click()

        select_field = WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located((By.ID, 'incCentral:formConteudo:selPadraoLancamento'))
        )
        select = Select(select_field)
        select.select_by_visible_text('[TRANSFERENCIA]')

        iniciar = navegador.find_element(By.ID, 'incCentral:formConteudo:btnIniciar')
        iniciar.click()

        return navegador

    def action_mov_int(navegador,cnpj_origem, cnpj_destino):
        select_cnpj = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.XPATH, f".//option[contains(@value, '{cnpj_origem}')]"))
        )
        select_cnpj.click()

        current_value = 0
        lista = open_file(caminho)
        
        for index, row in lista.iterrows():
            quantite_field = navegador.find_element(By.ID, 'incCentral:formConteudo:txtProduto')
            quantite_field.click()
            quantite_field.clear()
            quantite_field.send_keys(row.iloc[0])
           
            navegador.find_element(By.ID, 'incCentral:formConteudo:btnAdicionar').click()
            
            quantite = int(row.iloc[0].split('&')[0])
            current_value += quantite

            text_value = 0

            while current_value != text_value:
                time.sleep(2)
                html_content = navegador.page_source

                # Criar um objeto BeautifulSoup
                soup_content = BeautifulSoup(html_content, 'html.parser')

                # Usar BeautifulSoup para encontrar o elemento desejado
                item_span = soup_content.find('span', string='ITENS:')
                
                text_value = int(item_span.next_sibling.split(' ')[-1])
                
        navegador.find_element(By.ID, 'incCentral:formConteudo:btnPesqCliente').click()

        time.sleep(2)
        cnpj = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:formPnlModalPesquisaParticipante:txtPtcCoCnpjFiltro'))
        )
        cnpj.click()
        cnpj.send_keys(cnpj_destino)

        navegador.find_element(By.ID, 'incCentral:formPnlModalPesquisaParticipante:btnPsqPtcControlado').click()

        select_id_destino = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[starts-with(@id, 'incCentral:formPnlModalPesquisaParticipante:tblPsqInfoPtcControladoBody:0:')]"))
        )
        select_id_destino.click()


        cond_pag = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.ID, 'incCentral:formConteudo:selCondicaoPagto'))
        )
        select_pagamento = Select(cond_pag)
        select_pagamento.select_by_visible_text('[100] - TRANSFERENCIA')

        navegador.find_element(By.ID, 'incCentral:formConteudo:btnFinalizar').click()

        time.sleep(3)
        # Captura o HTML do elemento
        html_element = navegador.page_source

        soup_element = BeautifulSoup(html_element, 'html.parser')

        item_class = soup_element.find('li', class_='okMessageGrande')
        
        value_element = item_class.text.split(' ')[2]

        # Selecionar Vendas
        navegador.find_element(By.ID, 'opVendas').click()

        # Selecionar Pedidos
        navegador.find_element(By.XPATH, "//*[contains(@id, 'opPedidoVenda')]").click()

        # Habilitar o campo CNPJ
        habilitar_cnpj = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:incCentralVenda:formConteudo:formEmitente:selFiltroEmiCoCnpj'))
        )
        habilitar_cnpj.click()

        # Selecionar o CNPJ
        select_cnpj2 = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f".//option[contains(@value, '{cnpj_origem}')]"))
        )
        select_cnpj2.click()

        # Pesquisar
        navegador.find_element(By.ID, 'incCentral:incCentralVenda:formConteudo:btnPedPesquisar').click()

        # Localizar posicao do numero do pedido
        num_order = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//td[text()= '{value_element}']"))
        )
        position = num_order.location
        y_num_order = position['y']

        # Localizar todos os elementos "Pastel" que seguem o padrão no XPath
        elementos_pastel = navegador.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralVenda:formConteudo:tblPrdBodyPesquisa:')]")

        # Iterar sobre todos os elementos localizados
        for elemento in elementos_pastel:
            # Pegar a localização do elemento "Pastel"
            posicao_pastel = elemento.location
            y_pastel = posicao_pastel['y']

            # Comparar a coordenada y do pedido com o elemento pastel
            if y_pastel == y_num_order + 2:
                # Se as posições y forem iguais, clique no elemento
                elemento.click()
                break  # Para o loop após encontrar e clicar no elemento correto

        # Clicar em Liberar Faturamento
        faturamento = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:incCentralVenda:formConteudo:btnPedLiberar'))
        )
        faturamento.click()

        time.sleep(2)
        # Esperar até que o alerta apareça e interagir com ele
        alerta = Alert(navegador)

        # Aceitar o alerta clicando no botão "OK"
        alerta.accept()

        time.sleep(3)

        # Selecionar Faturamento
        navegador.find_element(By.XPATH, "//*[contains(@id, 'opFaturamento')]").click()

        # Habilitar o campo CNPJ
        habilitar_cnpj2 = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:incCentralVenda:formConteudo:formEmitente:selFiltroEmiCoCnpj'))
        )
        habilitar_cnpj2.click()

        # Selecionar o CNPJ
        select_cnpj3 = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f".//option[contains(@value, '{cnpj_origem}')]"))
        )
        select_cnpj3.click()

        # Pesquisar
        navegador.find_element(By.ID, 'incCentral:incCentralVenda:formConteudo:btnPdfPesquisar').click()

        # Localizar posicao do numero do pedido
        num_order = WebDriverWait(navegador, 20).until(
            EC.presence_of_element_located((By.XPATH, f"//td[text()= '{value_element}']"))
        )
        position = num_order.location
        y_num_order = position['y']

        # Localizar todos os elementos "Operacao" que seguem o padrão no XPath
        elementos_operacao = navegador.find_elements(By.XPATH, "//*[contains(@id, 'btnSelOperacaoPed')]")

        # Iterar sobre todos os elementos localizados
        for elemento in elementos_operacao:
            # Pegar a localização do elemento "Operacao"
            posicao_operacao = elemento.location
            y_operacao = posicao_operacao['y']

            # Comparar a coordenada y do pedido com o elemento operacao
            if y_operacao == y_num_order + 2:
                # Se as posições y forem iguais, clique no elemento
                elemento.click()
                break  # Para o loop após encontrar e clicar no elemento correto

        time.sleep(2)
        # Esperar até que o alerta apareça e interagir com ele
        alerta = Alert(navegador)

        # Aceitar o alerta clicando no botão "OK"
        alerta.accept()

        time.sleep(1)
        # Selecionar operacao
        navegador.find_element(By.ID, 'incCentral:incCentralVenda:frmmodalOperacao:rgnSelOperacaomodalOperacao').click()

        # Aguarde o select aparecer
        field_ope = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralVenda:frmmodalOperacao:selOpeCoOperacao')]"))
        )

        # Re-encontrar o elemento antes de usar Select
        select_ope = Select(field_ope)
        select_ope.select_by_visible_text('[10] - TRANSFERENCIA DE MERCADORIA')

        # navegador.find_element(By.XPATH, "//input[@value='Selecionar Operação']").click()

        time.sleep(10)
        navegador.close()

    navegador = enter_navegador()
    enter_mov_int(navegador)
    action_mov_int(navegador, cnpj_origem, cnpj_destino)

# Janela de login
root = Tk()
root.title('Login')
root.configure(background='DeepPink2')
root.geometry("500x200")
root.resizable(True, True)
root.minsize(width=500, height=200)

frame1 = Frame(root)
frame1.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

# Nome de usuário
Label(frame1, text="Nome de Usuário").place(relx=0.2, rely=0.2)
username_entry = Entry(frame1)
username_entry.place(relx=0.5, rely=0.2)

# Senha
Label(frame1, text="Senha").place(relx=0.2, rely=0.4)
password_entry = Entry(frame1, show="*")
password_entry.place(relx=0.5, rely=0.4)

# Botão para logar
login_button = Button(frame1, text="Logar", command=login)
login_button.place(relx=0.5, rely=0.6, relwidth=0.2)

# Carregar credenciais ao iniciar o programa
load_credentials()

root.mainloop()