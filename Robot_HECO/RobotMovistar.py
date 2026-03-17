# -*- coding: utf-8 -*-
"""
This code belongs to Web Builders company. 
Developed by Ramon Marino Solis.
Python 3.14 / Selenium 3.141.0
"""

import pandas as pd
import io
import sys
import json
import os
import warnings
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

# Desactivar advertencias
warnings.filterwarnings("ignore", category=DeprecationWarning) 

# --- FUNCIONES ---

def click_element_by_name(browser, element_name):
    try:
        WebDriverWait(browser, 9).until(
            EC.element_to_be_clickable((By.NAME, element_name))
        )
        browser.find_element(By.NAME, element_name).click()
    except Exception as e:
        print("Error al clicar el elemento: " + str(element_name))

def get_searcher(browser, element_id, frame_name_containing_element):
    try:
        element = browser.find_element(By.ID, element_id)
    except: 
        try:
            WebDriverWait(browser, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, frame_name_containing_element))
            )
            element = browser.find_element(By.ID, element_id)
        except:
            print('El frame o la barra de busqueda no estan disponibles')
            element = None
    return element

def sent_customer_data(browser, element, cif):
    if element:
        WebDriverWait(browser, 15).until(
            EC.visibility_of_element_located((By.NAME, 'Buscar')))
        element.clear()
        element.send_keys(cif)
        click_element_by_name(browser, 'Buscar')

# --- LOGICA PRINCIPAL ---

if len(sys.argv) < 4:
    print("Uso: python RobotMovistar.py <user> <pwd> <archivo.xlsx> [driver_path]")
    sys.exit()

username, password, path = sys.argv[1].strip(), sys.argv[2].strip(), sys.argv[3].strip()


config_path = os.path.join(os.path.dirname(__file__), 'config.json')
try:
    with open(config_path) as f:
        config = json.load(f)
    driver_path = os.path.expanduser(config['driver_path'])
except Exception as e:
    print(f"Error leyendo config.json: {e}")
    sys.exit()

options = EdgeOptions()

browser = None

try:
    service = EdgeService(executable_path=driver_path)
    browser = webdriver.Edge(service=service, options=options)
    
    try:
        lista_CIFs = pd.read_excel(path)
        print('Excel cargado correctamente.')
    except Exception as e:
        print('Error al cargar el Excel:', e)
        sys.exit()

    lista_CIFs = lista_CIFs.astype('object').dropna()
    lista_CIFs.columns = ['CIF']
    
    listaCIFs_boletines = pd.DataFrame(columns=[
        'Tipo de documento','Numero de documento','SEGMENTO',
        'LISTADO DE BOLETINES FECHA','LISTADO DE BOLETINES ACTUACION',
        'LISTADO DE BOLETINES DE FUSION FECHA PRIMERA','LISTADO DE BOLETINES DE FUSION ACTUACION PRIMERA',
        'LISTADO DE BOLETINES DE FUSION FECHA ULTIMA','LISTADO DE BOLETINES DE FUSION ACTUACION ULTIMA'
    ])

    browser.get("https://" + username + ":" + password + "@ecanal.telefonica.es/HECO")

    for row_index, row in lista_CIFs.iterrows():
        cif = row['CIF']
        segmento = boletines_fecha = boletines_actuacion = ' '
        f_f_pri = f_a_pri = f_f_ult = f_a_ult = ' '

        try:
            searcher = get_searcher(browser, 'dato', 'Cabecera')
            sent_customer_data(browser, searcher, cif)
            
            browser.switch_to.default_content()
            browser.switch_to.frame('Principal')
            
            tables = browser.find_elements(By.CLASS_NAME, 'sortable')

            if len(tables) >= 1:
                try:
                    html_t1 = tables[0].get_attribute('outerHTML')
                    df_t1 = pd.read_html(io.StringIO(html_t1), header=0, decimal=',', thousands='.', flavor='html5lib')[0]
                    segmento = df_t1.iloc[0]['Segmento']
                    boletines_fecha = df_t1.iloc[0]['Fecha']
                    boletines_actuacion = df_t1.iloc[0]['Actuación']
                except:
                    pass

                if len(tables) == 2:
                    try:
                        html_t2 = tables[1].get_attribute('outerHTML')
                        df_t2 = pd.read_html(io.StringIO(html_t2), header=0, decimal=',', thousands='.', flavor='html5lib')[0]
                        f_f_pri = df_t2.iloc[0]['Fecha']
                        f_a_pri = df_t2.iloc[0]['Actuación']
                        f_f_ult = df_t2.iloc[-1]['Fecha']
                        f_a_ult = df_t2.iloc[-1]['Actuación']
                    except:
                        print("Sin datos de fusion para: " + str(cif))

            data_to_add = {
                'Tipo de documento': 'CIF', 'Numero de documento': cif,
                'SEGMENTO': segmento, 'LISTADO DE BOLETINES FECHA': boletines_fecha,
                'LISTADO DE BOLETINES ACTUACION': boletines_actuacion,
                'LISTADO DE BOLETINES DE FUSION FECHA PRIMERA': f_f_pri,
                'LISTADO DE BOLETINES DE FUSION ACTUACION PRIMERA': f_a_pri,
                'LISTADO DE BOLETINES DE FUSION FECHA ULTIMA': f_f_ult,
                'LISTADO DE BOLETINES DE FUSION ACTUACION ULTIMA': f_a_ult
            }
            listaCIFs_boletines = pd.concat([listaCIFs_boletines, pd.DataFrame([data_to_add])], ignore_index=True)

        except Exception as e:
            print("Error procesando CIF " + str(cif) + ": " + str(e))
        
        browser.switch_to.default_content()

except Exception as main_e:
    print("Error critico: " + str(main_e))

finally:
    if browser:
        print('\nExportando resultados...')
        try:
            out = path.replace('.xlsx', '_boletines.xlsx')
            listaCIFs_boletines.to_excel(out, index=False)
            print("Guardado en: " + out)
        except:
            print("Error al guardar Excel")
        
        browser.quit()