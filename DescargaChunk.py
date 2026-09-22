import pandas as pd
import time
import os
import glob
import math
import sys
from dotenv import load_dotenv
import shutil
from collections import Counter
from datetime import datetime

#Subimos un nivel de carpetas, para poder hacer el import AutoDescargaChunk...
#Recordar que la ejecucion archivo.py coloca "Carpeta/" antes del archivo
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from AutoDescargaChunk.Conversion_parquet import ejecucion

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException

import concurrent.futures
import requests
import subprocess

from azureEntraID import carga_a_sharepoint

from FUNCIONES_DESCARGA.Manejo_ventana import *
from FUNCIONES_DESCARGA.Iniciar_driver import *
from FUNCIONES_DESCARGA.Funciones_logica_principal import *
from FUNCIONES_DESCARGA.Manejo_botones import *

def descargar_session_individual(session_uid, url_web, ruta_descarga, canal, fechas, opciones, estado):
        
    pag_carga = "//div[contains(@class, 'ant-modal-content')]"
    div_excel = '/html/body/div[3]/div/ul/li[1]'
    div_export = '/html/body/div/div/div[1]/div[2]/button[2]'
    div_filtro = "/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/div[2]/div/span/span"
    div_input_filtro_multi = '/html/body/div/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[3]/div/div[3]/div/div/div/div/div[1]/div/input[2]'
    div_fila_css = "div.ag-row[row-index='0']"

    div_apply_filtro_survey = "//button[contains(@class, 'ant-btn') and .//span[text()='Apply Filters']]"

    
    load_dotenv(dotenv_path='credenciales.env')
    usuario_arca = os.getenv('usuario_arca')
    password = os.getenv('contraseña_arca')
    
    driver_vacio = None
    ruta_completa_descarga = ruta_descarga + canal
    band = True

    try:

        driver = iniciar_driver(session_uid, ruta_completa_descarga, url_web, driver_vacio)
        inicio_sesion(driver, usuario_arca, password)
        esperar_invisibilidad(driver, pag_carga, timeout=120)
        click_survey(driver)
        time.sleep(1)

        resultado = manejo_botones(driver, estado, opciones, fechas, div_apply_filtro_survey, div_fila_css, div_excel, ruta_completa_descarga, session_uid, div_filtro, div_input_filtro_multi, pag_carga, div_export, band)

        return resultado

    except (WebDriverException, TimeoutException) as e:
        print(f"[{session_uid}] \nFALLO: Error de Selenium o Timeout: {e}")
        return f"Fallo: {session_uid} \n- Error de Selenium/Timeout."
        
    except Exception as e:
        print(f"[{session_uid}] \nFALLO: Ocurrió un error inesperado: {e}")
        return f"Fallo: {session_uid} \n- Error inesperado."

    finally:
        if driver:
            driver.quit()


# --- LÓGICA PRINCIPAL DE PRE-PROCESAMIENTO Y EJECUCIÓN PARALELA ---

if __name__ == '__main__':

    def Concurrencia(lista_uids):

        print(f"\n--- 2. INICIO DE DESCARGAS PARALELAS con {MAX_PROCESOS} procesos ---")

        chunks_uids = list(dividir_en_chunks(lista_uids, 6))

        with concurrent.futures.ProcessPoolExecutor(max_workers=MAX_PROCESOS) as executor:
            
            resultados = executor.map(
                descargar_session_individual,
                chunks_uids,
                [url_web] * len(chunks_uids),
                [ruta_descarga] * len(chunks_uids),
                [canal] * len(chunks_uids),
                [fechas] * len(chunks_uids),
                [opciones] * len(chunks_uids),
                ["1"] * len(chunks_uids)
            )
            
            for resultado in resultados:
                print(f"Resultado final: {resultado}")
                
        print("\n--- 3. PROCESO PARALELO FINALIZADO ---\n")


    eliminar_archivos_pasados('NaN', True)

#-----------------------------------------------------------------------------------------------------------------------------------------  
#    VARIABLES
#-----------------------------------------------------------------------------------------------------------------------------------------

    fecha_actual = datetime.now().strftime("%m/%d/%Y")
    load_dotenv(dotenv_path='credenciales.env')
    url_web = os.getenv('ruta_web')
    ahora = datetime.now()
    hora = ahora.hour
    minutos = ahora.minute

    if hora < 20 :
        fechas = [fecha_actual, fecha_actual] #"mm/dd/yyyy"
        #fechas = ['09/09/2026', '09/09/2026'] #"mm/dd/yyyy"

    else:
        hoy = datetime.now()
        primer_dia_actual = hoy.replace(day=1)
        primer_dia = primer_dia_actual.strftime("%m/%d/%Y")

        fechas = [primer_dia, fecha_actual] #"mm/dd/yyyy"

    lista_opciones = ['AUTOSERVICIO', 'C-STORE']#, 'AUTOAUDITORIA']
    #lista_opciones = ['AUTOSERVICIO']
    MAX_PROCESOS = 5 # Número de navegadores/procesos concurrentes


    for opcion in lista_opciones:
    
        ruta_descarga = r'C:\Users\bbartolome.DICHTER.000\Downloads'
        canal = rf"\{opcion}"

        opciones = ['',f'{opcion}']
        #opciones = [f'{opcion}']

        #-----------------------------------------------------------------------------------------------------------------------------------------
        # --- 1. PROCESO SECUENCIAL PREVIO (Obtener lista de UIDs) ---
        #-----------------------------------------------------------------------------------------------------------------------------------------
        
        eliminar_archivos_pasados(opcion, False)
        descargar_session_individual("NaN", url_web, ruta_descarga, canal, fechas, opciones, "0")

        #-----------------------------------------------------------------------------------------------------------------------------------------

        try:
            ruta_descargas_carpetas = ruta_descarga + canal
            patron = os.path.join(ruta_descargas_carpetas, 'Survey*.XLSX')
            archivos_encontrados = glob.glob(patron)
            
            Dataframe = subir_excel(archivos_encontrados, patron)
            
            if Dataframe is None:
                print("\nDataframe no se inicializó\n")
                continue
                #sys.exit(1)
            
            if Dataframe.empty:
                print("\nDataframe cargado está vacío\n")
                continue
                
            Dataframe_validos = Dataframe[(Dataframe['Session Review Status'] != 'Rejected') & (Dataframe['Survey Status'] != 'InComplete')].reset_index(drop=True)
            lista_uids = Dataframe_validos['Session Uid'].tolist()
            Dataframe_incompletos = Dataframe[(Dataframe['Session Review Status'] != 'Rejected') & (Dataframe['Survey Status'] == 'InComplete')].reset_index(drop=True)
            lista_incompletos = Dataframe_incompletos[['Outlet Code', 'Outlet Name', 'Survey Start Time']]
            lista_incompletos.to_excel(fr'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\VSCODE\Session Incomplete\INCOMPLETOS_{opcion}.xlsx', index=False)

            print(f"--- 1. Éxito: {len(lista_uids)} Session Uids válidos encontrados para descargar ---")


        except Exception as e:
            print(f"ERROR FATAL: Fallo en el pre-procesamiento para obtener la lista de UIDs. {e}")
            continue
            #sys.exit(1)
        
    #-----------------------------------------------------------------------------------------------------------------------------------------
        Concurrencia(lista_uids)
        mover_descargas(opcion)
        lista_rezagados = verificar_resultados(opcion, lista_uids)

        while len(lista_rezagados) > 0:

            print(f"No fueron descargadas {len(lista_rezagados)} mediciones. Se retoma concurrencia\n")
            Concurrencia(lista_rezagados)
            mover_descargas(opcion)
            lista_rezagados = verificar_resultados(opcion, lista_uids)

        print("Todas las mediciones fueron descargadas")
        
        try:
            requests.post(
            f"https://ntfy.sh/descarga_{opcion.lower()}",
            data=f"El proceso de descarga {opcion} de mediciones terminó con {len(lista_uids)} UIDs descargados",
            timeout=10
            )

        except requests.exceptions.RequestException as e:
            print(f"[WARN] No se pudo enviar notificación ntfy: {e}")

#Inicio de ejecucion de reporteria

    if (hora == 12 and minutos >= 30) or (hora == 15 and minutos >= 30) or (hora == 19):

        try:
            subprocess.run(
                ["python", "datos.py"],
                cwd=r"C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\VSCODE\Reporte_OT",
                check=True
            )
        except Exception as e:
            print(f"Error generando reportes: {e}")
    
    else: 
        print("\n--- SE GENERARÁN LOS REPORTES DESPUÉS DE LAS 12:30 PM / 3:30 PM ---")

#Inicio de ejecucion de carga a SharePoint
    
    if (hora >= 20):
        print("\n--- INICIANDO CARGA A SHAREPOINT ---")
        try:
            subcarpetas = ["CARPETA1","CARPETA2"]

            for opc in subcarpetas:
                ejecucion(opc)
                
            carga_a_sharepoint()

            requests.post(
                "https://ntfy.sh/descarga_auto_chunk",
                data="Descarga Y carga a SharePoint terminaron",
                timeout=10
            )
        except Exception as e:
            print(f"[ERROR] Fallo la carga a SharePoint: {e}")
    
    else:
        print("\n--- SE CARGARÁ A SHAREPOINT DESPUÉS DE LAS 8:00 PM ---")
