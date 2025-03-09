from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import pandas as pd
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Configuração do WebDriver
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)

try:
    # Acesso ao site
    navegador.get('https://quotes.toscrape.com/js-delayed/')
    WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "quote")))
    quotes = navegador.find_elements(By.CLASS_NAME, "quote")

    # Extração dos dados
    dados = []
    for quote in quotes:
        texto = quote.find_element(By.CLASS_NAME, "text").text
        autor = quote.find_element(By.CLASS_NAME, "author").text
        tags = [tag.text for tag in quote.find_elements(By.CLASS_NAME, "tag")]
        dados.append([texto, autor, ';'.join(tags)])

    # Salvar em CSV
    with open("quotes.csv", 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Citação', 'Autor', 'Tags'])
        writer.writerows(dados)

finally:
    navegador.quit()

# Processamento dos dados do CSV
def processar_dados_csv():
    df = pd.read_csv("quotes.csv")
    df.columns = df.columns.str.strip()
    
    qtd_citacoes = df.shape[0]
    autor_mais_recorrente = df['Autor'].mode()[0]
    df['Tags'] = df['Tags'].fillna('')
    tot_tags = df['Tags'].str.split(';').explode()
    tag_mais_recorrente = tot_tags.mode()[0]
    
    return qtd_citacoes, autor_mais_recorrente, tag_mais_recorrente

# Exibir dados processados
qtd_citacoes, autor_mais_recorrente, tag_mais_recorrente = processar_dados_csv()
print(f'Total de citações: {qtd_citacoes}')
print(f'Autor mais recorrente: {autor_mais_recorrente}')
print(f'Tag mais utilizada: {tag_mais_recorrente}')

# Função para envio de e-mail
def enviar_email():
    servidor = None  # Inicializar como None para garantir que o quit só será chamado se o servidor for criado
    try:
        
        usuario =  'albuquerquedevback@gmail.com'
        senha = "cvgx wrpb mldi ngdm" 

        if not usuario or not senha:
            print("Erro: Credenciais de e-mail não configuradas.")
            return

        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(usuario, senha)

        mensagem = MIMEMultipart()
        mensagem['From'] = usuario
        mensagem['To'] = 'paulo.andre@parvi.com.br, thiago.jose@parvi.com.br'
        mensagem['Subject'] = 'Relatório de Citações'

        corpo_email = f"""
        Relatório de Citações:
        
        Quantidade de Citações: {qtd_citacoes}
        Autor Mais Recorrente: {autor_mais_recorrente}
        Tag Mais Utilizada: {tag_mais_recorrente}

        O arquivo com todas as citações está anexado.
        """
        mensagem.attach(MIMEText(corpo_email, 'plain'))

        if not os.path.exists('quotes.csv'):
            print("Erro: O arquivo 'quotes.csv' não foi encontrado.")
            return

        with open('quotes.csv', 'rb') as arquivo:
            anexo = MIMEBase('application', 'octet-stream')
            anexo.set_payload(arquivo.read())
            encoders.encode_base64(anexo)
            anexo.add_header('Content-Disposition', 'attachment', filename='quotes.csv')
            mensagem.attach(anexo)

        servidor.send_message(mensagem)
        print('E-mail enviado com sucesso!')

    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

    finally:
        if servidor:  # Só chama quit se servidor foi iniciado
            servidor.quit()
            print('Conexão SMTP fechada.')

enviar_email()
