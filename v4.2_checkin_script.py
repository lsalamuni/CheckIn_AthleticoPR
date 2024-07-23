# -*- coding: utf-8 -*-
"""
Extracting selected emails from gmail account

1. Need to enable IMAP on gmail settings

2. In case of 2-factor authentication, it is needed to create an application specific password

"""

# Libraries
import time
import imaplib
import email
from email.header import decode_header
import os
import yaml # to load saved login credentials from a yaml file
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
import pandas as pd

# Setting a new directory
new_directory = 'C:\\Users\\lsala\\OneDrive\\Lucas Salamuni\\Geral\\Projetos\\Check-in'
os.chdir(new_directory)

# Initializing df
df = pd.DataFrame(columns=["Nome", "Data do Checkin", "Data do Jogo", "Autenticação"])

# Reading the yaml file with user and password
with open('email_user_password.yml') as f:
    content = f.read()

# Importing username and password from the given file
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# Loading the username as well as the password from the yaml file
user = my_credentials['user']
password = my_credentials['password']

# URL for IMAP connection
imap_url = 'imap.gmail.com'

# Further procedures  
try:
    # connecting with gmail via SSL
    my_mail = imaplib.IMAP4_SSL(imap_url) 
    my_mail.login(user, password)  # loging in with credentials

    my_mail.select('inbox')  # selecting the Inbox to fetch the messages

    # defining the filters
    from_filter = 'FROM "noreply@athletico.com.br"'
    subject_filter = 'HEADER SUBJECT "check-in"'
    result, data = my_mail.search(None, from_filter, subject_filter)

    if result == 'OK':
        mail_id_list = data[0].split()  # IDs of all emails to fetch

        if mail_id_list:
            latest_email_id = mail_id_list[-1]  # the last and most recent email
            result, data = my_mail.fetch(latest_email_id, '(RFC822)')

            if result == 'OK':
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                subject, encoding = decode_header(msg['subject'])[0]  # decoding the email's subject

                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or 'utf-8')

                from_, encoding = decode_header(msg['from'])[0]  # decoding the email's sender (from)

                if isinstance(from_, bytes):
                    from_ = from_.decode(encoding or 'utf-8')

                email_date = msg['date']
                print(f"Subject: {subject}")
                print(f"From: {from_}")
                print(f"Date: {email_date}")

                # extracting the email's link (i.e., the link for check-in)
                link = None
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True)
                        soup = BeautifulSoup(body, 'html.parser')
                        try:
                            area_tags = soup.find_all('area')
                            for area_tag in area_tags:
                                if area_tag.has_attr('alt') and 'check-in' in area_tag['alt'].lower():
                                    if area_tag.has_attr('href'):
                                        link = area_tag['href']
                                        break
                        except Exception as e:
                            print(f"Erro ao extrair o link: {e}")
                            link = None

                if link:
                    checkin_link = link
                    print(f"Link: {checkin_link}")
                else:
                    print("Nenhum link encontrado no email.")

            else:
                print(f"Erro ao buscar o e-mail com ID: {latest_email_id}")
        else:
            print("Nenhum e-mail encontrado com os critérios especificados.")
    else:
        print(f"Erro na busca de emails: {result}")

except imaplib.IMAP4.error as e:
    print(f"Erro de IMAP: {e}")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
finally:
    my_mail.logout()

# Lendo o arquivo YAML com os CPFs
with open('CPFs.yml') as f:
    content = f.read()
my_cpfs = yaml.load(content, Loader=yaml.FullLoader)
cpfs = my_cpfs['CPFs']  # Assumindo que os CPFs estão listados sob a chave 'CPFs'

# Loop para processar cada CPF
for cpf in cpfs:
    # Inicializando o Selenium para cada CPF
    options = Options()
    options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Navegando para o link de check-in
    driver.get(checkin_link)

    try:
        # Configurar uma espera explícita com um timeout de 30 segundos
        wait = WebDriverWait(driver, 5)
        input_field = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "input")))
        
        # Enviar CPF para o campo de entrada, se encontrado
        input_field.send_keys(cpf)

    except (NoSuchElementException, TimeoutException) as e:
        print(f"Elemento não encontrado ou tempo esgotado ao tentar encontrar o campo de entrada para CPF: {e}")
        driver.quit()
        continue

    try:
        # Clicar no botão de pesquisa
        click_pesquisar = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "button")))
        click_pesquisar.click()

        # Condicional: já fez check-in VS ainda não feito
        try:
            # Verificar a presença do elemento "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN"
            message_element_check_in = wait.until(
                EC.presence_of_element_located((By.XPATH, "//h3[contains(text(), 'Selecione os contratos que deseja fazer o check-in')]"))
            )

            if "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN" in message_element_check_in.text:
                print("Mensagem de seleção de contratos encontrada.")
                
                # Localize o checkbox e clique nele
                checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox']")))
                checkbox.click()
                
                time.sleep(2)
                
                # Localize o botão "CONFIRMAR" e clique nele
                confirmar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='button']")))
                confirmar_button.click()
                
                time.sleep(2)
                
                print("Check-in realizado com sucesso!")
            else:
                print("Check-in já realizado.")

        except (NoSuchElementException, TimeoutException, StaleElementReferenceException):
            print("Elemento de seleção de contratos não encontrado.")
            print(f"Check-in já realizado para o CPF {cpf}")

    except Exception as e:
        print("Ocorreu um erro:", e)
    finally:
        driver.quit()








