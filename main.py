import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client as win32


def verificar_termo_banido(nome, termos_banidos):
    # verificar sem tem termos banidos
    return any(palavra in nome for palavra in termos_banidos)


def verificar_todos_nomes(nome, nome_produto):
    return all(palavra in nome for palavra in nome_produto.split())


def unificar_bases_dados(df_google, df_buscape):
    # Criando dataframe vazio
    df_geral = pd.DataFrame(columns=['Produto', 'Preço', 'Link'])

    # Lista para acumular os DataFrames (quando usamos concat com apenas um dataframe, substitui o anterior)
    dataframes = []

    if not df_google.empty:
        dataframes.append(df_google)
    if not df_buscape.empty:
        dataframes.append(df_buscape)

    if dataframes:
        df_geral = pd.concat(dataframes, axis=0)
        df_geral = df_geral.sort_values('Preço', ascending=True, ignore_index=True)

    return df_geral


def google_shopping(navegador, nome_produto, termos_banidos, minimo_preco, maximo_preco):

    lista_produtos = []

    # abrir no google
    SITE_GOOGLE = r'https://www.google.com.br'
    navegador.get(SITE_GOOGLE)

    # pesquisar o item da tabela
    CLASS_BARRA_PESQUISA_GOOGLE = 'gLFyf'
    WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, CLASS_BARRA_PESQUISA_GOOGLE)))
    navegador.find_element(By.CLASS_NAME, CLASS_BARRA_PESQUISA_GOOGLE).send_keys(nome_produto, Keys.ENTER)

    # clicar em shopping
    CSS_BARRA_NAVEGACAO = "[role='listitem']"
    WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, CSS_BARRA_NAVEGACAO)))   
    barra_navegacao = navegador.find_elements(By.CSS_SELECTOR, CSS_BARRA_NAVEGACAO) 
    for opcao in barra_navegacao:
        if 'Shopping' in opcao.text:
            opcao.click()
            break

    # pegando todos os produtos
    CLASS_GERAL_PRODUTOS = 'i0X6df'
    lista_produtos_site = navegador.find_elements(By.CLASS_NAME, CLASS_GERAL_PRODUTOS)
    for produto in lista_produtos_site:
        CLASS_PRECO_PRODUTO = 'a8Pemb'
        CLASS_LINK_PRODUTO = 'bONr3b'
        CLASS_NOME_PRODUTO = 'tAxDx'
        preco = produto.find_element(By.CLASS_NAME, CLASS_PRECO_PRODUTO).text
        link = produto.find_element(By.CLASS_NAME, CLASS_LINK_PRODUTO).find_element(By.XPATH, '..').get_attribute('href')
        nome = produto.find_element(By.CLASS_NAME, CLASS_NOME_PRODUTO).text.lower()

        # verificar sem tem termos banidos
        tem_termos_banidos = verificar_termo_banido(nome, termos_banidos)

        # verificar se todas as palavras do produto estão no nome
        tem_todos_nomes = verificar_todos_nomes(nome, nome_produto)

        if tem_termos_banidos or not tem_todos_nomes:
            continue

        # tratar os dados do preço
        try:
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco)
        except ValueError:
            continue

        if preco < minimo_preco or preco > maximo_preco:
            continue

        dicionario_produto = {
            'Produto': nome,
            'Preço': preco,
            'Link': link,
        }

        lista_produtos.append(dicionario_produto)

    df_produtos = pd.DataFrame(lista_produtos)
    return df_produtos


def buscape(navegador, nome_produto, termos_banidos, minimo_preco, maximo_preco):

    lista_produtos = []

    # abrir o buscapé
    SITE_BUSCAPE = r'https://www.buscape.com.br/'
    navegador.get(SITE_BUSCAPE)

    # pesquisar pelo produto
    XPATH_BARRA_PESQUISA_BUSCAPE = '/html/body/div[1]/main/header/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input'
    WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.XPATH, XPATH_BARRA_PESQUISA_BUSCAPE)))
    navegador.find_element(By.XPATH, XPATH_BARRA_PESQUISA_BUSCAPE).send_keys(nome_produto, Keys.ENTER)

    # pegando todos os resultados
    CLASS_GERAL_PRODUTOS = 'ProductCard_ProductCard_Inner__gapsh'
    WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, CLASS_GERAL_PRODUTOS)))
    lista_produtos_site = navegador.find_elements(By.CLASS_NAME, CLASS_GERAL_PRODUTOS)
    for produto in lista_produtos_site:
        CLASS_PRECO_PRODUTO = 'Text_MobileHeadingS__HEz7L'
        CLASS_NOME_PRODUTO = 'ProductCard_ProductCard_Name__U_mUQ'
        nome = produto.find_element(By.CLASS_NAME, CLASS_NOME_PRODUTO).text.lower()
        preco = produto.find_element(By.CLASS_NAME, CLASS_PRECO_PRODUTO).text
        link = produto.get_attribute('href')

        # verificar sem tem termos banidos
        tem_termos_banidos = verificar_termo_banido(nome, termos_banidos)

        # verificar se todas as palavras do produto estão no nome
        tem_todos_nomes = verificar_todos_nomes(nome, nome_produto)

        if tem_termos_banidos or not tem_todos_nomes:
            continue

        # tratar os dados do preço
        try:
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco)
        except ValueError:
            continue

        if preco < minimo_preco or preco > maximo_preco:
            continue

        dicionario_produto = {
            'Produto': nome,
            'Preço': preco,
            'Link': link,
        }

        lista_produtos.append(dicionario_produto)

    df_produtos = pd.DataFrame(lista_produtos)
    return df_produtos


def enviar_email(df_geral, html_body):
    # Mandar email
    if not df_geral.empty:
        outlook = win32.Dispatch('outlook.application')

        mail = outlook.CreateItem(0)
        mail.To = 'tuliovcr2+COTACOES@gmail.com'
        mail.Subject = 'Cotações na faixa de preço desejadas.'
        mail.HTMLBody = f"""
        <p>Prezado,</p>
        <p></p>
        <p>Segue em anexo as cotações dos produtos solicitados.</p>
        <p></p>
        {html_body}
        <p></p>
        <p>Atenciosamente,</p>
        <p>Túlio Rocha</p>
        """
        mail.Send()


# pegar os dados da planilha
ARQUIVO_BUSCA = r'buscas.xlsx'

buscas_df = pd.read_excel(ARQUIVO_BUSCA)

html_body = ''

# abrir o navegador
navegador = webdriver.Chrome()

# para cada produto dentro da planilha
for linha in buscas_df.index:

    nome_produto = buscas_df.loc[linha, 'Nome'].lower()
    termos_banidos = buscas_df.loc[linha, 'Termos banidos'].lower().split(" ")
    minimo_preco = float(buscas_df.loc[linha, 'Preço mínimo'])
    maximo_preco = float(buscas_df.loc[linha, 'Preço máximo'])

    # fazendo a pesquisa no google shopping
    df_google = google_shopping(navegador, nome_produto, termos_banidos, minimo_preco, maximo_preco)
    df_buscape = buscape(navegador, nome_produto, termos_banidos, minimo_preco, maximo_preco)  

    df_geral = unificar_bases_dados(df_google, df_buscape)  

    html_body += f"""<p></p>
    <p>{df_geral.head().to_html(index=False)}</p>
    <p></p>
    """

navegador.close()

enviar_email(df_geral, html_body)





