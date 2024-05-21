import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pythoncom
import win32com.client as win32

# function that returns last week monday's date
def mondayOfLastWeek():
    
    today = datetime.now().date()
    todayWeekday = today.weekday()
    lastMonday = today - timedelta(days=todayWeekday)
    lastWeekMonday = lastMonday - timedelta(weeks=1)
    f_lastWeekMonday = lastWeekMonday.strftime('%d/%m/%Y')
    
    return f_lastWeekMonday

# function that returns last sunday's date
def lastSunday():

    today = datetime.now().date()
    todayWeekday = today.weekday()
    
    match todayWeekday:
        case 0:
            days = 1
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 1:
            days = 2
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 2:
            days = 3
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 3:
            days = 4
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 4:
            days = 5
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 5:
            days = 6
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday
        case 6:
            days = 7
            sunday = today - timedelta(days=days)
            sunday = sunday.strftime('%d/%m/%Y')
            return sunday

# function that sends the email with the attached file
def sendEmail():
    
    # initializes COM library
    pythoncom.CoInitialize()
    
    try:
        # creating integration with outlook
        outlook = win32.Dispatch('Outlook.Application')
        
        # creating an email object
        email = outlook.CreateItem(0)
        
        # configuring email information
        email.To = "iluccafx@hotmail.com"
        email.Subject = f'Relatório Mapa RH - {startDate} a {endDate}'
        email.HTMLBody = f'''
        <p>Bom dia!</p> 

        <p>Segue em anexo Relatório de Mapeamento referente a semana anterior.</p>

        <p>Att,</p>
        '''
        
        # adds file
        file = r'c:\Users\PC\Downloads\ExportERP.xls'
        email.Attachments.Add(file)
        
        # sending email
        email.Send()
        print('Email sent successfully!')
    
    except Exception as e:
        print(f'Error: {e}')
    
    finally:
        # ends COM library
        pythoncom.CoUninitialize()

# opening the browser and logging into the Extranet
browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
browser.implicitly_wait(10)
browser.maximize_window()
browser.get('https://extranet.lopesrio.com.br/ERP/RelatorioMapaRH.aspx')
log = browser.find_element(By.XPATH, '//input[@name="tLogin"]')
login = 'xxxxxxxxxxx'
for i in login:
    log.send_keys(i)
    time.sleep(0.01)
browser.find_element(By.XPATH, '//input[@name="tSenha"]').send_keys('xxxxxxx')
browser.find_element(By.XPATH, '//a[@id="btLogin"]').click()

# defining search period
startDate = mondayOfLastWeek()
endDate = lastSunday()

# selecting search parameters
browser.find_element(By. XPATH, '/html/body/form/main/div/div[1]/div[3]/div/div/div/div/div/div[1]/select/option[2]').click()

browser.find_element(By. XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tini"]').send_keys(startDate)

browser.find_element(By. XPATH, '//*[@id="ctl00_ContentPlaceHolder1_tFim"]').send_keys(endDate)

browser.find_element(By. XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlDiretor"]/option[5]').click()

browser.find_element(By. XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btFiltro"]').click()

# downloading file
browser.find_element(By. XPATH, '//*[@id="wrapper"]/div/div[1]/div[3]/div/div/div[2]/div[3]/div/div[2]/button').click()

time.sleep(10)

# sending email
sendEmail()

time.sleep(5)

# deleting file after sent by email
file = r'c:\Users\PC\Downloads\ExportERP.xls'

try:
    os.remove(file)
    print(f"'{file}' successfully deleted!")
except FileNotFoundError:
    print(f"Error: '{file}' not found.")
except PermissionError:
    print(f"Error: permission denied to delete '{file}'.")
except Exception as e:
    print(f"An error occurred when trying to delete the file: {e}")
