# Sistema de Cotação Automatizada

Este projeto automatiza a busca de produtos em sites como **Google Shopping** e **Buscapé**, unificando os resultados e enviando um e-mail com as cotações de produtos que atendem a critérios predefinidos.

## Funcionalidades

- **Pesquisa Automatizada**: Realiza pesquisas em dois sites (Google Shopping e Buscapé) para encontrar produtos com base em palavras-chave.
- **Filtragem Avançada**: Verifica termos banidos e limita os resultados a uma faixa de preços especificada.
- **Unificação de Resultados**: Consolida os resultados de ambos os sites, removendo duplicatas e ordenando por preço.
- **Envio Automatizado de E-mails**: Envia os resultados das cotações por e-mail, formatados em HTML.

## Bibliotecas Utilizadas

- **[Selenium](https://www.selenium.dev/)**: Utilizado para automação de navegação e coleta de dados de produtos e preços em sites de e-commerce.
- **[Pandas](https://pandas.pydata.org/)**: Manipulação de dados, criação e ordenação de DataFrames, e consolidação dos resultados.
- **[win32com.client](https://pypi.org/project/pywin32/)**: Utilizado para integração com o Outlook e envio de e-mails automáticos.

## Como Funciona

1. **Leitura da Planilha de Produtos**: O sistema lê os produtos a serem pesquisados a partir de um arquivo `buscas.xlsx`, que contém o nome do produto, termos banidos e faixa de preço.
2. **Pesquisa nos Sites**: Utilizando o Selenium, o sistema realiza a pesquisa em ambos os sites, extrai os nomes dos produtos, links e preços.
3. **Verificação e Filtragem**: Cada produto é verificado para garantir que não contenha termos banidos e que seu preço esteja dentro do intervalo especificado.
4. **Unificação de Bases de Dados**: Os resultados das pesquisas no Google Shopping e Buscapé são unificados em um DataFrame, ordenados por preço.
5. **Envio de E-mail**: Um e-mail contendo as cotações é enviado automaticamente, com o conteúdo formatado em HTML.

## Estrutura do Projeto

- **google_shopping**: Função que automatiza a pesquisa de produtos no Google Shopping.
- **buscape**: Função que automatiza a pesquisa de produtos no Buscapé.
- **verificar_termo_banido**: Função para verificar se o nome do produto contém termos banidos.
- **verificar_todos_nomes**: Função para garantir que todas as palavras-chave do produto estejam no nome.
- **unificar_bases_dados**: Função que consolida os resultados de pesquisa dos dois sites.
- **enviar_email**: Função responsável por enviar o e-mail com as cotações.

## Como Executar

### Requisitos

- **Python 3.8+**
- **Google Chrome** instalado e o **ChromeDriver** correspondente configurado no PATH.
- Pacotes necessários:

```bash
pip install selenium pandas pywin32 openpyxl
```

## Executando o Script

1. Preencha o arquivo `buscar.xlsx` com os seguintes campos
    - **Nome:** Nome do produto a ser pesquisado
    - **Termos banidos:** Palavras que não devem aparecer no nome do produto (separadas por espaço)
    - **Preço mínimo:** Preço mínimo desejado.
    - **Preço máximo:** Preço máximo desejado
2. Execute o script

    ```bash
    python main.py
    ```
3. Os resultados serão enviados por e-mail automaticamente.

## Contribuição

Sinta-se à vontade para abrir issues ou pull requests com melhorias, correções ou sugestões para o projeto.# cotacoes_google_buscape
