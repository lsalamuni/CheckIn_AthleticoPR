# -*- coding: utf-8 -*-
"""
Extracting selected emails from gmail account

1. Need to enable IMAP on gmail settings

2. In case of 2-factor authentication, it is needed to create an application specific password

"""

# Libraries
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
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import pandas as pd
import re

# Setting a new directory
new_directory = 'C:\\Users\\lsala\\OneDrive\\Lucas Salamuni\\Geral\\Projetos\\Check-in'
os.chdir(new_directory)

# Initializing df
df = pd.DataFrame(columns = ["Name", "Checkin_Date", "Authentication"])

# Reading the yaml file with user and password
with open('email_user_password.yml') as f:
    content = f.read()

# Importing username and password from the given file
my_credentials = yaml.load(content, 
                           Loader = yaml.FullLoader)

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
                                        print(f"Link: {link}")
                                        break
                        except Exception as e:
                            print(f"Erro ao extrair o link: {e}")
                            link = None

                if link:
                    checkin_link = link
                    print(f"Link de check-in armazenado: {checkin_link}")
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

# Initializing selenium
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()),
                          options = options)

# Navigating to the check-in link
driver.get(checkin_link)

# Inserting the CPF into the input field
with open('CPFs.yml') as f:
    content = f.read()

my_cpfs = yaml.load(content, 
                    Loader = yaml.FullLoader)

cpf = my_cpfs['CPF']

try:
    # Configurar uma espera explícita com um timeout de 10 segundos
    wait = WebDriverWait(driver, 10)
    input_field = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "input")))
    
    # Enviar CPF para o campo de entrada, se encontrado
    input_field.send_keys(cpf)

except NoSuchElementException as e:
    print(f"Elemento não encontrado: {e}")

# Clicking the search button
click_pesquisar = driver.find_element(By.CLASS_NAME, "button")
click_pesquisar.click()

# Conditional: already checked-in VS not yet done
try:
    # Aguarde até que um dos elementos com as mensagens especificadas esteja presente na página
    wait = WebDriverWait(driver, 10)
    
    # Verificar a presença do elemento "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN"
    try:
        message_element_check_in = wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div/div[2]/div/div[1]/div[1]"))
        )
        if "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN" in message_element_check_in.text:
            print("Mensagem de seleção de contratos encontrada.")
            # Localize o checkbox e clique nele
            checkbox = driver.find_element(By.XPATH, "//input[@type='checkbox']")
            checkbox.click()
            
            # Localize o botão "CONFIRMAR" e clique nele
            confirmar_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div/div[2]/div/div[1]/div[3]/button")
            confirmar_button.click()
            
            print("Check-in realizado com sucesso!")
        else:
            name_element = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='name']"))
                )
            checkin_date_element = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='hour large']"))
                )
            authentication_element = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='text-center chave']"))
                )
            
            name = name_element.text
            checkin_date = checkin_date_element.text
            pattern = r"(\d{2}/\d{2}/\d{4})"
            match = re.search(pattern, checkin_date)
            if match:
                checkin_date = match.group(1)
            else:
                checkin_date = ""
            authentication = authentication_element.text
            
            new_row = pd.DataFrame({"Name": [name], "Checkin_Date": [checkin_date], "Authentication": [authentication]})
            
            df = pd.concat([df, new_row], ignore_index = True)
            
            driver.quit()
            
            print("Check-in já foi realizado. Fechando a página.")
            print(df)

    except NoSuchElementException:
        print("Elemento de seleção de contratos não encontrado.")

except Exception as e:
    print("Ocorreu um erro:", e)