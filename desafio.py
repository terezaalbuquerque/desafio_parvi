from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import csv 

# Configuração e acesso ao site


navegador = webdriver.Chrome()
navegador.get('https://quotes.toscrape.com/js-delayed/')

WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "quote")))
quotes = navegador.find_elements(By.CLASS_NAME, "quote")


# Extrair e salvar dados em arquivo CSV, encontrando todas as citações, autores e tags 
dados = []
for quote in quotes: 
    texto = quote.find_element(By.CLASS_NAME, "text").text
    autor = quote.find_element(By.CLASS_NAME, "author").text
    tags = [tag.text for tag in quote.find_elements(By.CLASS_NAME, "tag")]
    dados.append([texto, autor, ';'.join(tags)])


with open("quotes.csv", 'w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(['Citação', 'Autor', 'Tags'])
    writer.writerows(dados)

navegador.quit()

# Leitura e exibição_df do arquivo CSV em Pandas

import pandas as pd

def proc_dados():
    df = pd.read_csv("quotes.csv")
    df.columns = df.columns.str.strip()
    print(df.head())

    numero_quotes = df.shape[0]
    autor_mais_frequente = df['Autor'].mode()[0]
    tot_tags = df ['Tags'].str.split(';').explode()
    tag_mais_frequente = tot_tags.mode()[0]

    print(f'Total de citações: {numero_quotes}')
    print(f'Autor mais recorrente: {autor_mais_frequente}')
    print(f'Tag mais utilizada: {tag_mais_frequente}')


proc_dados()

# Adicionar um log detalhado de depuração para verificar a leitura do arquivo
def enviar_email():
    try:
        usuario = os.getenv('EMAIL_USER')
        senha = os.getenv('EMAIL_PASSWORD')
        
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        print('Conexão com o servidor SMTP estabelecida.')
        
        servidor.login(usuario, senha)

        # Obter os dados processados para o corpo do e-mail
        qtd_citacoes, autor_mais_recorrente, tag_mais_recorrente = processar_dados_csv() # type: ignore

        if qtd_citacoes is None:
            print("Erro ao processar o arquivo CSV. E-mail não enviado.")
            return

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

        try:
            with open('quotes.csv', 'rb') as arquivo:
                arquivo_conteudo = arquivo.read()
                
                if arquivo_conteudo is None or len(arquivo_conteudo) == 0:
                    print("Erro: O arquivo 'quotes.csv' está vazio ou não pôde ser lido.")
                    return
                else:
                    print(f"Conteúdo do arquivo lido com sucesso. Tamanho do conteúdo: {len(arquivo_conteudo)} bytes.")

                anexo = MIMEBase('application', 'octet-stream')
                anexo.set_payload(arquivo_conteudo)
                encoders.encode_base64(anexo)
                anexo.add_header('Content-Disposition', 'attachment', filename='quotes.csv')
                mensagem.attach(anexo)
        except Exception as e:
            print(f"Erro ao abrir o arquivo 'quotes.csv': {e}")
            return

        servidor.send_message(mensagem)
        print('E-mail enviado com sucesso!')

    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

    finally:
        try:
            if 'servidor' in locals():
                servidor.quit()
                print('Conexão SMTP fechada.')
        except Exception as e:
            print(f"Erro ao fechar a conexão SMTP: {e}")