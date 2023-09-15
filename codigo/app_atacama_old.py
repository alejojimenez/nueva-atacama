import os
import time
import shutil
import requests
import pandas as pd
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains

class Scraper_Atacama():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None    

    def login(self):
        
        driver_exe = '.domain\\chromedriver.exe'
        credencials = '.\\config\\credenciales.xlsx'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password
        
        options = webdriver.ChromeOptions()
        options.add_extension('C:\\roda\\nueva-atacama\\config\\33.9_0.crx')
        options.add_experimental_option('prefs', {
        "download.default_directory": "C:\\roda\\nueva-atacama\\input\\", #Change default directory for downloads
        "download.prompt_for_download": False, #To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
        })
            
        self.driver = webdriver.Chrome(driver_path, options=options)
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        # Seleccionar usuario dentro de la pagina
        intentos = 0
        usuario = True
        while (usuario):
            try:
                print('Try en la pestaña usuario..', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_usuario= self.driver.find_element(By.ID, 'username')
                element_usuario.clear()
                element_usuario.click()
                element_usuario.send_keys(email)
                usuario = False
            except:    
                print('Exception en la pestaña usuario')
                print('----------------------------------------------------------------------')
                usuario = intentos <= 3                

        # Seleccionar clave dentro de la pagina
        intentos = 0
        clave = True
        while (clave):
            try:
                print('Try en la pestaña clave..', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                time.sleep(5)
                element_clave= self.driver.find_element(By.ID, 'password')
                element_clave.clear()
                element_clave.click()
                element_clave.send_keys(password)
                clave = False   
            except:
                print('Exception en la pestaña clave')
                print('----------------------------------------------------------------------')
                clave = intentos <= 3 

        # Seleccionar clave dentro de la pagina
        intentos = 0
        ingresar = True
        while (ingresar):
            try:
                print('Try en el borton de ingreso...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                time.sleep(5)
                element_ingreso= self.driver.find_element(By.ID, 'submitIngresar')
                if element_ingreso.is_displayed():
                    action_chains = ActionChains(self.driver)
                    action_chains.click(element_ingreso).perform()
                
                #element_clave.click()
                ingresar = False   
            except:
                print('Exception en el borton de ingreso')
                print('----------------------------------------------------------------------')
                ingresar = intentos <= 3

        print('Hicimos login, comenzamos scrapping')
         
    def scrapping_codiner(self):

        print('Entrando en la funcion Scrapping...')
        print('----------------------------------------------------------------------')
        
        folder_path_config = './config/'
        
        # Especifica la ruta de tu archivo Excel
        excel_file = folder_path_config + "clientes.xlsx"

        # Especifica el nombre de la hoja en la que se encuentran los datos
        hoja_excel = "Hoja1"

        # Carga los datos de Excel en un DataFrame
        df = pd.read_excel(excel_file, sheet_name=hoja_excel)
        print('Dataframe ', df)
        print('----------------------------------------------------------------------')

        #Esperamos detectar iframe para poder obtener los datos de la tabla
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "tbody")))
        
        tabla = element_ingreso= self.driver.find_element(By.TAG_NAME, 'tbody')
        filas = tabla.find_elements(By.TAG_NAME, 'tr')
        
        for index,boton in enumerate(filas):
            print(index,'-',boton.get_attribute('innerText'))
        
        largo = len(filas)
        
        #Iteramos para obtener las url de descarga y la fecha
        i=2
        while i < largo+1:

            #Obtenemos la fecha
            intentos = 0
            get_fecha = True
            while (get_fecha):
                try:
                    print('Try en obtener la fecha...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_fecha = self.driver.find_element(By.XPATH, f'/html/body/div[7]/main/div/div[1]/div/div[2]/table/tbody/tr[{i}]/td[1]')
                    fecha  = element_fecha.text
                    mes,año = fecha.split()
                    time.sleep(5)

                    get_fecha = False   
                except:
                    print('Exception en el borton de ingreso')
                    print('----------------------------------------------------------------------')
                    get_fecha = intentos <= 3
                    
            #Obtenemos el boton de descarga
            intentos = 0
            get_descarga = True
            while (get_descarga):
                try:
                    print('Try en hacer click sobre boton de descarga...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_descarga = self.driver.find_element(By.XPATH, f'/html/body/div[7]/main/div/div[1]/div/div[2]/table/tbody/tr[{i}]/td[2]/a')
                    element_descarga.click()
                    time.sleep(5)
                    get_descarga = False   
                except:
                    print('Exception en hacer click sobre boton de descarga')
                    print('----------------------------------------------------------------------')
                    get_descarga = intentos <= 3
 
            #Entreamos un tiempo para que descargue
            time.sleep(10)

            folder_path = './input/'
            
            # Cruce datos faltantes para ontener
            for index, row in df.iterrows():
                    
                df_nro_cliente = df.loc[index, 'nro_cliente']
                df_sucursal = df.loc[index, 'sucursal']
                print('Nro. Cliente: ', df_nro_cliente, 'Sucursal: ', df_sucursal)
                print('--------------------------------------------------------------------------')

            #Buscamos el archivo que se haya descargado que comience con el nombre boleta para poder moverlo a la carpeta input
            for filename in os.listdir(folder_path):
                if filename.startswith("RE") and filename.endswith('.pdf'):
                    nombre_archivo = filename
                    print(nombre_archivo)
                    print('Numero factura: ', nombre_archivo[21:-22])
                    print('----------------------------------------------------------------------')
                    try:
                        os.rename(folder_path+nombre_archivo,folder_path+f'{df_nro_cliente}_{nombre_archivo[21:-22]}_{año}.pdf')
                    except:
                        print('archivo corrupto o no disponible')

            i+=1        
            print('pasamos al siguiente archivo')
            
        print('Terminamos descargas')
