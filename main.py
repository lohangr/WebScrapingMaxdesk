from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from time import sleep
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.support.ui import Select
import openpyxl

# Configuracoes do navegador
options = Options()
options.add_argument("--headless=new")

# Inicia o navegador
navegador = webdriver.Chrome(options=options)

link = "https://nome_empresa.maxdesk.us/admin/login"
navegador.get(url=link)
sleep(1)

# Preenche campo email
inputEmail = navegador.find_element(By.ID, "f-email_address")
inputEmail.send_keys("seu_email")
sleep(1)

# Preenche campo senha
inputSenha = navegador.find_element(By.ID, "f-password")
inputSenha.send_keys("sua_senha")
sleep(1)

# Clica no botao Login
botaoLogin = navegador.find_element(By.CSS_SELECTOR, ".button.large.radius")
botaoLogin.click()
sleep(1)

# Clica no botao Clientes
botaoClientes = navegador.find_element(By.CSS_SELECTOR, ".ic-customers")
botaoClientes.click()
sleep(1)

# Funcao para coletar observacoes de um cliente
def coletar_observacoes():
    # Clica na aba 'Observacoes'
    abaObservacoes = navegador.find_element(By.ID, "js-customer-notes-tab")
    abaObservacoes.click()
    sleep(1)
    # Rolar pagina para cima
    navegador.execute_script("window.scrollTo(0, 0)")

    # Seleciona exibicao de 50 itens
    select = Select(navegador.find_element(By.XPATH, '//div[@id="DataTables_Table_4_length"]/label/select'))
    select.select_by_value("50")
    sleep(1)

    # Funcao que procura todos os campos de observacao
    def procurar_campos_observacao():
        camposObservacao = navegador.find_elements(By.CSS_SELECTOR, ".notes-note-content")
        observacoes = []
        for campo in camposObservacao:
            htmlContent = campo.get_attribute("innerHTML")
            soup = BeautifulSoup(htmlContent, "html.parser")
            paragrafos = soup.find_all("p")
            note_content = "\n".join([paragrafo.get_text(strip=True) for paragrafo in paragrafos])
            # Verificar se ha a tag <br> no conteudo
            if soup.find("br"):
                # Quebra uma linha para cada tag <br>
                note_content = note_content.replace("<br>", "\n")
            observacoes.append(note_content)
        return observacoes

    # Coleta todos campos de observacao da pagina atual
    observacoes_pagina = procurar_campos_observacao()

    return observacoes_pagina

# Lista para armazenar os dados coletados
data = []

# Funcao que coleta os clientes em uma pagina
def coletar_clientes():
    # Obtem o código HTML da pagina atual
    html = navegador.page_source

    # Cria um objeto BeautifulSoup para analisar o HTML
    soup = BeautifulSoup(html, "html.parser")

    # Encontra todas tags 'a' dentro de 'td' com a classe 'sorting_1', na pagina de Clientes
    links = soup.select('td.sorting_1 a[href]')

    # Itera sobre os links encontrados para acessar os clientes
    for link in links:
        href = link['href']
        url_completa = "https://nome_empresa.maxdesk.us" + href

        # Texto de identificacao do href
        identificacao = link.get_text(strip=True)

        # Acessa o cliente
        navegador.get(url=url_completa)
        sleep(1)

        # Coleta as observacoes do cliente
        observacoes_cliente = coletar_observacoes()

        # Verifica se o cliente possui observacoes
        if observacoes_cliente:
            # Adiciona o nome do cliente seguido pelas observacoes coletadas
            data.append("%" + identificacao + "%")
            data.extend(observacoes_cliente)
            # Adiciona uma linha em branco apos cada campo de observacao
            data.append("")
                    
        # Retorna para a pagina 'Clientes'
        botaoClientes = navegador.find_element(By.CSS_SELECTOR, ".ic-customers")
        botaoClientes.click()
        sleep(1)

# Coleta clientes da pagina 1
coletar_clientes()

# Encontra o botao da pagina 2
botaoPagina2 = navegador.find_element(By.XPATH, '//a[@class="button white pagination" and text()="2"]')
botaoPagina2.click()
sleep(1)

# Coleta clientes da pagina 2
coletar_clientes()

# Encontra o botao da pagina 3
botaoPagina3 = navegador.find_element(By.XPATH, '//a[@class="button white pagination" and text()="3"]')
botaoPagina3.click()
sleep(1)

# Coleta clientes da pagina 3
coletar_clientes()

# Salvar os dados em um arquivo XLSX
df = pd.DataFrame(data, columns=["Observacoes"])
df.to_excel("Observacoes.xlsx", index=False, sheet_name="Sheet1", engine="openpyxl")