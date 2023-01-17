from selenium import webdriver
import chromedriver_autoinstaller
from credentials import *
from selenium.webdriver.common.by import By
import time
import os
import shutil
import os
import pandas as pd
import xlwings as xw
clear = lambda: os.system('cls')


def tratamento_base():
    try:
        for files in os.listdir(path_venv_tel):
            if 'xls' in files and not 'xlsx' in files:
                split = files.split('.')
                df = pd.read_html(path_venv_tel + files)
                df[1].to_excel(path_venv_tel + split[0] + str('.xlsx'))
    except:
        pass

    try:
        for files in os.listdir(path_venv_tel):
            if '.xlsx' in files:
                df = pd.read_excel(path_venv_tel + files)
                df = df[
                            ~(df['DATA'] == 'TOTAL')
                        ]
                df.to_excel(path_venv_tel + files)  
    except:
        pass
             
    for files in os.listdir(path_venv_tel):
        if '.xls' in files and not 'xlsx' in files:
            os.remove(path_venv_tel + files)   
            
def login(url, login, passw, template):
    driver.get(url)
    driver.find_element(By.XPATH, '//*[@id="UserTxt"]').send_keys(login)
    driver.find_element(By.XPATH, '//*[@id="Password"]').send_keys(passw)
    driver.find_element(By.XPATH, '//*[@id="BtnOK"]').click()
    driver.find_element(By.XPATH, '//*[@id="ctl00_PageMenu_LinkButton4"]').click()
    driver.find_element(By.XPATH, '//*[@id="PageMenu_lblMenuLatReports"]').click()
    driver.find_element(By.XPATH, '//*[@id="PageMenu_menu1__labelMenuTitle_reports_agent_view"]').click()
    driver.find_element(By.XPATH, '//*[@id="PageMenu_menu1_submenu_reports_agents_day_agent"]').click()
    driver.find_element(By.XPATH, '//*[@id="PageContent_search1_StartDate"]').send_keys(data_considerada)
    driver.find_element(By.XPATH, '//*[@id="PageContent_search1_EndDate"]').send_keys(data_considerada)
    driver.find_element(By.XPATH, '//*[@id="PageContent_search1_EndDate"]').click()
    driver.find_element(By.XPATH, template).click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="PageContent_search1_btn_xls"]').click()
    waiting_file = True
    while waiting_file == True:
        for files in os.listdir(path_download):
            if 'Report' in files and not 'crd' in files:
                shutil.copy(path_download + files, path_venv_tel + files)
                waiting_file = False
                break
    windows = driver.window_handles        
    for window in windows:
        driver.switch_to.window(window)
        driver.close()
            

            
def rename_file(name_file):
    for files in os.listdir(path_venv_tel):
        if 'Report' in files:
            split = files.split('.')
            os.rename(path_venv_tel + files, path_venv_tel + name_file + str('.') + split[1])
                    
def remove_files():
    for files in os.listdir(path_download):
        os.remove(path_download + files)
                
def remove_logs():
    try:
        for files in os.listdir(path_venv_tel):
            if '.xls' in files:
                os.remove(path_venv_tel + files)
    except:
        print('Arquivos não encontrado')  



while True:
    if time.strftime('%H:%M') == '18:37':
        valor = int(time.strftime('%d')) - 1
        data_considerada = str(valor)+time.strftime('/%m/%Y')
        path_venv_tel = '\\\\TELFSBAR01\\Planejamento$\\2022\\43 - MIS\\22 - Relatórios Olos\\'
        path_download = 'C:\\Users\\Roboplan\\Downloads\\'


        chromedriver_autoinstaller.install()
        driver = webdriver.Chrome()
            
        remove_logs()
            
        for position in range(len(urls)):
            if position == 0:
                remove_files()
                login(urls[position], login_sac, passw_sac, '//*[@id="PageContent_search1_DDTemplate"]/option[15]')
                rename_file('LOG N1 - SAC')
            if position == 1:
                driver = webdriver.Chrome()
                remove_files()
                login(urls[position], login_olos, passw_olos, '//*[@id="PageContent_search1_DDTemplate"]/option[30]')
                rename_file('LOG N1 - TEL')
            if position == 2:
                driver = webdriver.Chrome()
                remove_files()
                login(urls[position], login_cob, passw_cob, '//*[@id="PageContent_search1_DDTemplate"]/option[28]')
                rename_file('LOG COB - TEL')
                
        tratamento_base()
    else:
        clear()
        print(time.strftime('%H:%M:%S'))
        time.sleep(1)

                


    
        
    