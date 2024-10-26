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

caminho = "//Mac/Home/Downloads/alteracao_preco_teste_PR.xlsx"

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

# Funcao para alterar o preco
def alteracao_preco(caminho):
    import warnings
    warnings.simplefilter(action='ignore', category=FutureWarning)
    
    def loading(navegador):
        time.sleep(1)
        # Espera até que o status de carregamento mude para "display: none"
        WebDriverWait(navegador, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='_viewRoot:status.start'][style='display: none']"))
        )

    def enter_navegador():
        # Carregar credenciais do arquivo
        with open('credentials.txt', 'r') as file:
            lines = file.readlines()
            username = lines[0].strip()
            password = lines[1].strip()

        navegador = webdriver.Chrome()
        navegador.get("https://sumire-phd.homeip.net:8490/eVendas/home.faces")
        wait = WebDriverWait(navegador,10)
        user_name = navegador.find_element(By.ID, "form-login")
        user_password = wait.until(EC.presence_of_element_located((By.ID, "form-senha")))

        user_name.send_keys(username)
        user_password.send_keys(password)

        button_login = navegador.find_element(By.ID, "form-submit")
        button_login.click()

        field_cadastro = wait.until(EC.element_to_be_clickable((By.ID, "opCadastros")))
        field_cadastro.click()

        field_precos = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'opPrecos')]")))
        field_precos.click()

        field_manutencao = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'opManutCustoProd')]")))
        field_manutencao.click()
        return navegador
    
    def box_message_nenhuma(navegador):
        try:
            # Espera até que o elemento com a classe 'divNenhumaLinha' esteja presente
            element_nenhuma_linha = navegador.find_element(By.XPATH, "//*[contains(@class, 'divNenhumaLinha')]")
            # Aqui você pode interagir com o elemento ou retornar o texto, por exemplo
            texto = element_nenhuma_linha.text
            return texto
        except:
            return None
    
    def box_message_td(navegador, tipo):
        if tipo == "grupo":
            try:
                # Tenta capturar o texto no primeiro XPath
                element_tds = navegador.find_elements(By.XPATH, "//*[@id='incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody']//td[@class='tblLinha']")
                if element_tds:
                    descricao = element_tds[0].text
                    texto = str(descricao.split(' - ')[0])
                    return texto
            except:
                return None
        else:
            try:
                # Se não encontrar o primeiro elemento, tenta capturar o segundo XPath
                element_tds = navegador.find_elements(By.XPATH, "//*[@id='incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody']//td[@class='tblLinha']")
                if element_tds:
                    descricao = element_tds[0].text
                    texto = str(descricao.split(' - ')[0])
                    return texto
            except:
                return None
    
    def get_value(navegador, tipo):
        if tipo == "grupo":
            try:
                # Encontra o elemento pelo ID
                elemento_input = navegador.find_element(By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:txtGrpCoGrupoFiltro")
                if elemento_input:
                    # Obtém o valor do atributo 'value'
                    valor = elemento_input.get_attribute('value')
                    return str(valor)
            except:
                return None
        else:
            try:
                elemento_input = navegador.find_element(By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:txtPrdCoProdutoFiltro")
                if elemento_input:
                    # Obtém o valor do atributo 'value'
                    valor = elemento_input.get_attribute('value')
                    return str(valor)
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
    
    def seleciona_loja(navegador, num_loja):
        time.sleep(1)
        try:
            loja = dict_grupos.get(num_loja)
            element_loja = WebDriverWait(navegador,10).until(EC.visibility_of_element_located((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:selEmiCoCnpj")))
            select_loja = Select(element_loja)
            select_loja.select_by_visible_text(loja)
            return navegador, True
        except:
            return navegador, False
    
    def inclui_data_inicio(navegador, data):
        field_data = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'txtPvdDtIniValidade')]")))
        # Usar JavaScript para limpar o campo
        navegador.execute_script("arguments[0].value = '';", field_data)
        
        # Usar JavaScript para definir o valor no campo
        navegador.execute_script("arguments[0].value = arguments[1];", field_data, data)
        
        # Disparar o evento 'change' para garantir que o valor seja registrado
        navegador.execute_script("arguments[0].dispatchEvent(new Event('change'));", field_data)
        return navegador

    def selecionar_grupo_preco(navegador, code, tipo):
        loading(navegador)
        code = str(code)
        
        # Clica no botão para selecionar o grupo de preço
        button_grupo_preco = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Sel. o Grupo Preço']")))
        button_grupo_preco.click()
        
        time.sleep(2)

        while True:
            if get_value(navegador, tipo) == "":
                break
            if box_message_nenhuma(navegador) is not None or box_message_td(navegador, tipo) is not None:
                break
            time.sleep(1)

        value = get_value(navegador, tipo)
        first_message_td = box_message_td(navegador, tipo)
        first_message_nenhum = box_message_nenhuma(navegador)
        print(f"{code} value - {value}")
        print(f"{code} td - {first_message_td}")
        print(f"{code} nenhum - {first_message_nenhum}")

        if value == code:
            if first_message_nenhum != None:
                navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqGrupoPreco').component.hide()")
                return navegador, f"ERRO - {first_message_nenhum}"
            else:             
                pastel_all_codes = WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody:')]")))
                pastel_codes = pastel_all_codes.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody:')]")
                # Executa o clique no segundo elemento usando JavaScript
                navegador.execute_script("arguments[0].click();", pastel_codes[1])
                return navegador, first_message_td

        field_code_grupe = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:txtGrpCoGrupoFiltro")))
        field_code_grupe.click()
        field_code_grupe.clear()
        field_code_grupe.send_keys(code)

        button_pesquisar = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:btnPsqGrupoPreco")))
        button_pesquisar.click()

        time.sleep(3)

        while True:
            if box_message_td(navegador, tipo) is not None:
                if code == box_message_td(navegador, tipo):
                    break
            if box_message_nenhuma(navegador) is not None:
                break
            time.sleep(1)
        
        mensagem_td = box_message_td(navegador, tipo)
        mensagem_nenhuma = box_message_nenhuma(navegador)  
        print(f"{code} td - {mensagem_td}")
        print(f"{code} nenhum - {mensagem_nenhuma}")

        if mensagem_nenhuma is not None:
            navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqGrupoPreco').component.hide()")
            return navegador, f"ERRO - {mensagem_nenhuma}"
        else:
            if code == mensagem_td:
                pastel_all_codes = WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody:')]")))
                pastel_codes = pastel_all_codes.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaGpr:tblPsqInfoGrupoPrecoBody:')]")
                # Executa o clique no segundo elemento usando JavaScript
                navegador.execute_script("arguments[0].click();", pastel_codes[1])
                return navegador, mensagem_td
            else:
                navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqGrupoPreco').component.hide()")
                return navegador, "ERRO"

    def selecionar_produto(navegador, code, tipo):
        loading(navegador)
        code = str(code)
        
        # Clica no botão para selecionar o produto
        button_produto = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Selecionar o Produto']")))
        button_produto.click()

        time.sleep(2)

        while True:
            if get_value(navegador, tipo) == "":
                break
            if box_message_nenhuma(navegador) is not None or box_message_td(navegador, tipo) is not None:
                break
            time.sleep(1)

        first_message_td = box_message_td(navegador, tipo)
        first_message_nenhum = box_message_nenhuma(navegador)
        value = get_value(navegador, tipo)
        print(f"{code} value - {value}")
        print(f"{code} td - {first_message_td}")
        print(f"{code} nenhum - {first_message_nenhum}")

        if value == code:
            if first_message_nenhum != None:
                navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqProduto').component.hide()")
                return navegador, f"ERRO - {first_message_nenhum}"
            else:
                # Espera até que o elemento seja clicável incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:0:j_id278
                pastel_all_codes= WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")))
                # Aqui estamos buscando os mesmos elementos que o 'pastel_all_codes' se refere
                pastel_codes = pastel_all_codes.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")
                # Executa o clique no segundo elemento usando JavaScript
                navegador.execute_script("arguments[0].click();", pastel_codes[1])
                time.sleep(2)
                span = capturar_codigo_span(navegador)
                if span != None:
                    return navegador, f"ERRO - Grupo de preco {span}"
                return navegador, first_message_td

        field_code_grupe = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:txtPrdCoProdutoFiltro")))
        field_code_grupe.click()
        field_code_grupe.clear()
        field_code_grupe.send_keys(code)

        button_pesquisar = WebDriverWait(navegador,10).until(EC.element_to_be_clickable((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:btnPsqProduto")))
        button_pesquisar.click()

        time.sleep(3)

        while True:
            if box_message_td(navegador, tipo) is not None:
                if code == box_message_td(navegador, tipo):
                    break
            if box_message_nenhuma(navegador) is not None:
                break
            time.sleep(1)
    
        mensagem_td = box_message_td(navegador, tipo)
        mensagem_nenhuma = box_message_nenhuma(navegador)
        print(f"{code} td - {mensagem_td}")
        print(f"{code} nenhum - {mensagem_nenhuma}")

        if mensagem_nenhuma is not None:
            navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqProduto').component.hide()")
            return navegador, f"ERRO - {mensagem_nenhuma}"
        else:
            if mensagem_td == code:
                    # Espera até que o elemento seja clicável incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:0:j_id278
                    pastel_all_codes = WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")))
                    # Aqui estamos buscando os mesmos elementos que o 'pastel_all_codes' se refere
                    pastel_codes = pastel_all_codes.find_elements(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:incCentralDiversos:formPnlModalPesquisaPrd:tblPsqInfoParticipanteBody:')]")
                    # Executa o clique no segundo elemento usando JavaScript
                    navegador.execute_script("arguments[0].click();", pastel_codes[1])
                    time.sleep(2)
                    span = capturar_codigo_span(navegador)
                    if span != None:
                        return navegador, f"ERRO - {span}"
                    return navegador, mensagem_td
            else:
                navegador.execute_script("document.getElementById('incCentral:incCentralDiversos:incCentralDiversos:pnlPsqProduto').component.hide()")
                return navegador, "ERRO"
            
    def validacao_dados(navegador):
        WebDriverWait(navegador,10).until(EC.visibility_of_element_located((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:pnlPrecoVendaForm")))
        start_time_loja = time.time()
        check_loja = ''
        while True:
            # Localiza o elemento <select> pelo ID 
            select_element = navegador.find_element(By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:selEmiCoCnpj")

            # Localiza o <option> que está selecionado
            selected_option = select_element.find_element(By.XPATH, "./option[@selected='selected']")

            # Obtém o texto do <option> selecionado
            option_text = selected_option.text
            # Usa uma expressão regular para capturar o número entre colchetes
            match = re.search(r'\[(.*?)\]', option_text)

            if match:
                check_loja = match.group(1)  # Extrai o valor capturado
                break
            
            if time.time() - start_time_loja > 5:
                break
        
        start_time_code = time.time()
        check_code = ''
        while True:
            # Localiza o elemento span usando o ID
            span_element = navegador.find_element(By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:pnlNoPrdFiltro")

            # Obtém o texto do elemento
            span_text = span_element.text

            # Usa uma expressão regular para capturar o número dentro dos colchetes
            match = re.search(r'\[(\d+)\]', span_text)
            
            if match:
                check_code = str(match.group(1))  # Extrai o número capturado
                break
            if time.time() - start_time_code > 5:
                break
        
        return check_loja, check_code

    def atualiza_preco(navegador, vl_custo, vl_revenda):
        try:
            field_custo = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:incCentralDiversos:formConteudo:txtPvdVlCustoReposicao')
        except:
            field_custo = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:incCentralDiversos:formConteudo:txtGpvVlCustoReposicao')
        field_custo.click()
        field_custo.clear()
        field_custo.send_keys(vl_custo)

        try:
            field_revenda = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:incCentralDiversos:formConteudo:txtPvdVlVendaRevenda')
        except:
            field_revenda = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:incCentralDiversos:formConteudo:txtGpvVlVendaRevenda')
        field_revenda.click()
        field_revenda.clear()
        field_revenda.send_keys(vl_revenda)

        button_salvar = navegador.find_element(By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:btngpvPrecoAddEdt")
        button_salvar.click()

        start_time_save = time.time()
        while True:
            # Aguardar até que o elemento seja encontrado ou o tempo limite seja atingido
            ul_element = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, "incCentral:incCentralDiversos:incCentralDiversos:formConteudo:msgEndGlobal")))

            # Localiza o <li> com a classe 'okMessage' dentro do <ul>
            li_element = ul_element.find_element(By.CLASS_NAME, "okMessage")

            # Obtém o texto do <li>
            mensagem_ok = li_element.text

            if mensagem_ok == 'Salvo com sucesso!':
                return navegador, True
            if time.time() - start_time_save > 10:
                return navegador, False

    def novo_nome_csv():
        # Extrair diretório e nome base do arquivo Excel
        dir_name, file_name = os.path.split(caminho)
        base_name, ext = os.path.splitext(file_name)

        # Criar o novo nome de arquivo CSV no mesmo diretório
        csv_file_name = f"{base_name}_alteracao_preco_parcial.csv"  # Altera a extensão para .csv
        csv_file_path = os.path.join(dir_name, csv_file_name)
        return csv_file_path
    
    def xml_csv():
        # Carregar a planilha
        df = pd.read_excel(caminho, engine='openpyxl')
        
        # Normalizar o case da coluna 'Tipo do Codigo' para aceitar qualquer valor independente de maiúsculas/minúsculas
        df['Tipo do Codigo'] = df['Tipo do Codigo'].str.lower()  # Converte para minúsculas para padronizar

        try:
            # Substituir vírgula por ponto em 'Vl. Custo' e 'Vl. Revenda' se necessário
            df['Vl. Custo'] = df['Vl. Custo'].astype(str).str.replace(',', '.').astype(float)
        except:
            pass
        try:
            df['Vl. Revenda'] = df['Vl. Revenda'].astype(str).str.replace(',', '.').astype(float)
        except:
            pass
        try:
            # Converter a coluna 'Data inicio' para datetime e depois para o formato desejado
            df['Data inicio'] = pd.to_datetime(df['Data inicio'], errors='coerce')  # Converter para datetime
            df['Data inicio'] = df['Data inicio'].dt.strftime('%d/%m/%Y')  # Formatar como dia/mês/ano
        except:
            pass

        csv_file_path = novo_nome_csv()

        # Salvar o DataFrame como CSV no mesmo diretório
        df.to_csv(csv_file_path, sep=";", index=False)
        return df

    def arquivo_final():
        # Definir o caminho do arquivo existente (parcial)
        old_file_path = novo_nome_csv()

        # Extrair diretório e nome base do arquivo
        dir_name, file_name = os.path.split(old_file_path)
        base_name, ext = os.path.splitext(file_name)

        # Substituir a palavra 'parcial' por 'final' no nome do arquivo
        new_file_name = file_name.replace('parcial', 'final')
        new_file_path = os.path.join(dir_name, new_file_name)
        
        # Renomear o arquivo
        os.rename(old_file_path, new_file_path)

    def analisar_linha(df):
        navegador = enter_navegador()
        # Iterar sobre as linhas da planilha
        for idx, row in df.iterrows():
            # Verificar o status da linha
            data = row['Data inicio']
            status = str(row['Status'])
            codigo = row['Produto/Grupo']
            lojas_total = row['Loja/Grupo']
            lojas = lojas_total.split(',')
            tipo = row['Tipo do Codigo']
            vl_custo = row['Vl. Custo']
            vl_revenda = row['Vl. Revenda']

            if status.startswith("OK") or status.startswith("ERRO"):
                # Se o status for OK ou ERRO, pular para a próxima linha
                continue

            elif status.startswith("PARCIAL"):
                # Se o status for PARCIAL, obter as lojas já analisadas
                lojas_analisadas = status.split('-')[-1].split(',')
                # Encontrar a próxima loja/grupo a ser analisada
                lojas_pendentes = [loja for loja in lojas if loja not in lojas_analisadas]
                analisadas = status
                if len(lojas_pendentes) == 0:
                    df.at[idx, 'Status'] = "OK"
                    df.to_csv(novo_nome_csv(), sep=";" ,index=False)
                    continue
            else:
                # Se não for PARCIAL, todas as lojas/grupos estão pendentes
                lojas_pendentes = lojas
                analisadas = "PARCIAL-"

            inclui_data_inicio(navegador, data)

            # Verificar se ainda há lojas/grupos pendentes
            if lojas_pendentes:
                for loja in lojas_pendentes:
                    navegador, bool_loja = seleciona_loja(navegador, loja)
                    if bool_loja is False:
                        analisadas_parcial = analisadas.split("-")[0] + f" Loja {loja} nao localizada."
                        if analisadas.split("-")[-1]:
                            analisadas = analisadas_parcial + "-" + analisadas.split("-")[-1]
                        else:
                            analisadas = analisadas_parcial + "-"
                        df.at[idx, 'Status'] = analisadas
                        df.to_csv(novo_nome_csv(), sep=";" ,index=False)
                        continue
                    else:
                        if tipo == 'produto':
                            print("PRODUTO")
                            navegador, mensagem = selecionar_produto(navegador, codigo, tipo)
                        else:
                            print("GRUPO")
                            navegador, mensagem = selecionar_grupo_preco(navegador, codigo, tipo)
                        # Atualizar o status com a mensagem retornada
                        if mensagem.startswith("ERRO"):
                            df.at[idx, 'Status'] = mensagem
                            df.to_csv(novo_nome_csv(), sep=";" ,index=False)
                            break
                        else:
                            check_loja, check_code = validacao_dados(navegador)
                            if check_loja == loja and check_code == str(codigo):
                                navegador, bool_status = atualiza_preco(navegador, vl_custo, vl_revenda)
                                if bool_status is False:
                                    continue
                                else:
                                    if analisadas.split("-")[-1]:
                                        analisadas = analisadas + "," + loja
                                    else:
                                        analisadas = analisadas + loja
                                    df.at[idx, 'Status'] = analisadas
                                    df.to_csv(novo_nome_csv(), sep=";" ,index=False)
                            else:
                                continue

                new_status = analisadas.split("-")[-1]
                if lojas_total == new_status:
                    df.at[idx, 'Status'] = "OK"
                    df.to_csv(novo_nome_csv(), sep=";" ,index=False)

    def analisar_planilha():
        try:
            df = pd.read_csv(novo_nome_csv(), sep=";")
        except:
            df = xml_csv()
    
        # Verificar se há algum status "PARCIAL" ou em branco
        if df['Status'].isnull().any() or any(df['Status'].str.startswith("PARCIAL")):
            
            # Rodar a função para continuar o processamento
            analisar_linha(df)
        else:
            # Se todos os status forem "OK" ou "ERRO", rodar arquivo_final()
            arquivo_final()
            

    analisar_planilha()
alteracao_preco(caminho)
        # # Obtendo o conteúdo HTML da página com Selenium
        # html_content_pos = navegador.page_source

        # # Analisando o HTML com BeautifulSoup e formatando
        # soup_pos = BeautifulSoup(html_content_pos, 'html.parser')
        # pretty_html_pos = soup_pos.prettify()

        # # Definindo o caminho do arquivo para salvar o HTML no diretório do projeto
        # file_path_pos = os.path.join(os.getcwd(), f"pagina_formatada_{code}.html")

        # # Salvando o HTML formatado em um arquivo
        # with open(file_path_pos, "w", encoding="utf-8") as file:
        #     file.write(pretty_html_pos)