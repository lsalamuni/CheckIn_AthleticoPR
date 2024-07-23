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
import yaml
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
from openpyxl import load_workbook
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
import smtplib

# Setting a new directory
new_directory = 'C:\\Users\\lsala\\OneDrive\\Lucas Salamuni\\Geral\\Projetos\\Check-in'
os.chdir(new_directory)

# Initializing df
df = pd.DataFrame(columns=["Nome", "Data do Checkin", "Data do Jogo", "Campeonato", "Autenticação"])

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
    my_mail.login(user, password)  # logging in with credentials

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
            print(f"Check-in já realizado para o CPF {cpf}")

    except Exception as e:
        print("Ocorreu um erro:", e)
    finally:
        driver.quit()

# Loop para processar cada CPF
for cpf in cpfs:
    # Inicializando o Selenium para cada CPF
    options = Options()
    options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Navegando para o link de check-in
    driver.get(checkin_link)

    try:
        # Configurar uma espera explícita com um timeout de 10 segundos
        wait = WebDriverWait(driver, 30)
        input_field = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "input")))

        # Enviar CPF para o campo de entrada, se encontrado
        input_field.send_keys(cpf)

    except NoSuchElementException as e:
        print(f"Elemento não encontrado: {e}")
        driver.quit()
        continue

    # Clicar no botão de pesquisa
    click_pesquisar = driver.find_element(By.CLASS_NAME, "button")
    click_pesquisar.click()

    try:
        # Aguarde até que um dos elementos com as mensagens especificadas esteja presente na página
        wait = WebDriverWait(driver, 100)

        # Verificar a presença do elemento "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN"
        try:
            message_element_check_in = wait.until(
                EC.presence_of_element_located((By.XPATH, "//h3[contains(text(), 'Seu check-in já foi realizado')]"))
            )
            if "SEU CHECK-IN JÁ FOI REALIZADO" in message_element_check_in.text:
                print("Mensagem encontrada.")

                try:
                    comprovante_button = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[@class='button button-comprovante']"))
                    )
                    comprovante_button.click()

                    time.sleep(2)

                    name_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='name']"))
                    )
                    checkin_date_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='hour large']"))
                    )
                    match_date_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='data']"))
                    )
                    tournament_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='campeonato']"))
                    )
                    authentication_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='text-center chave']"))
                    )

                    time.sleep(2)

                    name = name_element.text
                    checkin_date = checkin_date_element.text
                    match_date = match_date_element.text
                    pattern = r"(\d{2}/\d{2}/\d{4})"
                    match = re.search(pattern, checkin_date)
                    if match:
                        checkin_date = match.group(1)
                    else:
                        checkin_date = ""
                    tournament = tournament_element.text
                    authentication = authentication_element.text

                    new_row = pd.DataFrame({"Nome": [name], "Data do Checkin": [checkin_date],
                                            "Data do Jogo": [match_date],
                                            "Campeonato": [tournament],
                                            "Autenticação": [authentication]})

                    df = pd.concat([df, new_row], ignore_index=True)

                    print("Check-in já foi realizado.")
                    print(df)

                except TimeoutException:
                    print("Período de check-in terminou")
            else:
                print("Erro no XPATH: mensagem não encontrada.")

        except NoSuchElementException:
            print("Elemento de seleção de contratos não encontrado.")
            
    except Exception as e:
        print("Fim do período de check-in.")
    
    finally:
        driver.quit()

# Loop para processar cada CPF
for cpf in cpfs:
    # Inicializando o Selenium para cada CPF
    options = Options()
    options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Navegando para o link de check-in
    driver.get(checkin_link)

    try:
        # Configurar uma espera explícita com um timeout de 10 segundos
        wait = WebDriverWait(driver, 20)
        input_field = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "input")))

        # Enviar CPF para o campo de entrada, se encontrado
        input_field.send_keys(cpf)

    except NoSuchElementException as e:
        print(f"Elemento não encontrado: {e}")
        driver.quit()
        continue

    # Clicar no botão de pesquisa
    click_pesquisar = driver.find_element(By.CLASS_NAME, "button")
    click_pesquisar.click()

    try:
        # Aguarde até que um dos elementos com as mensagens especificadas esteja presente na página
        wait = WebDriverWait(driver, 20)

        # Verificar a presença do elemento "SELECIONE OS CONTRATOS QUE DESEJA FAZER O CHECK-IN"
        try:
            message_element_check_in = wait.until(
                EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Comprovante de check-in')]"))
            )
            if "COMPROVANTE DE CHECK-IN" in message_element_check_in.text:
                print("Mensagem encontrada.")

                try:
                    name_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='name']"))
                    )
                    checkin_date_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='hour large']"))
                    )
                    match_date_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='data']"))
                    )
                    tournament_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='campeonato']"))
                    )
                    authentication_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//div[@class='text-center chave']"))
                    )

                    time.sleep(2)

                    name = name_element.text
                    checkin_date = checkin_date_element.text
                    match_date = match_date_element.text
                    pattern = r"(\d{2}/\d{2}/\d{4})"
                    match = re.search(pattern, checkin_date)
                    if match:
                        checkin_date = match.group(1)
                    else:
                        checkin_date = ""
                    tournament = tournament_element.text
                    authentication = authentication_element.text

                    new_row = pd.DataFrame({"Nome": [name], "Data do Checkin": [checkin_date],
                                            "Data do Jogo": [match_date],
                                            "Campeonato": [tournament],
                                            "Autenticação": [authentication]})

                    df = pd.concat([df, new_row], ignore_index=True)

                    print("Check-in já foi realizado.")
                    print(df)

                except TimeoutException:
                    print("Período de check-in terminou")
            else:
                print("Erro no XPATH: mensagem não encontrada.")

        except NoSuchElementException:
            print("Elemento de seleção de contratos não encontrado.")

    except Exception as e:
        print("Ocorreu um erro:", e)
        
    finally:
        driver.quit()

# Save the DataFrame to a xlsx file
df.sort_values(by="Nome", inplace=True)
df.to_excel('checkin_data.xlsx', index=False)
print("Dados salvos no arquivo 'checkin_data.xlsx'")


# Save the DataFrame to a xlsx file
df.sort_values(by="Nome", inplace=True)
df.to_excel('checkin_data.xlsx', index=False)
print("Dados salvos no arquivo 'checkin_data.xlsx'")


# Save the DataFrame to a xlsx file
df.sort_values(by="Nome", inplace=True)
df.to_excel('checkin_data.xlsx', index=False)
print("Dados salvos no arquivo 'checkin_data.xlsx'")

# Email content
with open('email_content.yml', encoding="utf-8") as f:
    email_content = yaml.load(f, Loader=yaml.FullLoader)

file_path = email_content["path"]

# Load the xlsx file in openpyxl
wb = load_workbook(file_path)
ws = wb.active

# Adjusting the xlsx columns
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

wb.save(file_path)

# Reading the YAML file with user and password
with open('email_user_password_ii.yml') as f:
    content = f.read()

# Importing username and password from the given file
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# Loading the username as well as the password from the YAML file
user = my_credentials['user']
password = my_credentials['password']

# Organizing the email content
recipient = email_content['recipient'].split(", ")
title = email_content['title']
message = email_content['message']

# Setting up the MIME
msg = MIMEMultipart()
msg["From"] = user
msg["To"] = ", ".join(recipient)
msg["Subject"] = title

# Attaching the body with the msg instance
msg.attach(MIMEText(message, "plain", "utf-8"))

# Guessing the MIME type and subtype of the attachment
mime_type, _ = mimetypes.guess_type(file_path)
mime_type, mime_subtype = mime_type.split('/')

# Open the file to be sent
with open(file_path, "rb") as file:
    mime_base = MIMEBase(mime_type, mime_subtype)
    mime_base.set_payload(file.read())
    encoders.encode_base64(mime_base)
    mime_base.add_header("Content-Disposition", f"attachment; filename={file_path}")
    msg.attach(mime_base)

try:
    # Creating the server connection
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(user, password)

    # Sending the email
    text = msg.as_string()
    server.sendmail(user, recipient, text)
    server.quit()

    print("E-mail enviado com sucesso!")

except Exception as e:
    print(f"Falha ao enviar o e-mail: {e}")