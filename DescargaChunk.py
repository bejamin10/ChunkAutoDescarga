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
from Conversion_parquet import ejecucion

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


def descargar_session_individual(session_uid, url_web, ruta_descarga, canal, fechas, opciones, estado):
    
    # def esperar_invisibilidad(driver, ruta, timeout=120):
    #     try:
    #         WebDriverWait(driver, timeout).until(EC.invisibility_of_element_located((By.XPATH, ruta)))
    #         return True
    #     except TimeoutException:
    #         print(f"[{session_uid}] \nDemasiada espera por invisibilidad.")
    #         return True 

    def esperar_invisibilidad(driver,ruta , timeout=120):
        wait = WebDriverWait(driver, timeout)
        
        try:
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "anticon-loading")))
        
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ant-modal-wrap")))
            
            return True
        
        except TimeoutException:
            print("El modal no desapareció correctamente")
            return False

    def tipo_elemento(driver, ruta, elemento, timeout=60):
        diccionario = {
            'clickable': WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, ruta))),
            'existente': WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, ruta)))
        }
        return diccionario[elemento]

    def tipo_elemento_css(driver, ruta, elemento, timeout=120):
        diccionario = {
            'css': WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, ruta)))
        }
        return diccionario[elemento]

    def inicio_sesion(driver, user, pwd):
        div_usuario = '/html/body/div/div/div[3]/div/form/div/div[2]/div[1]/div/input'
        div_password = '/html/body/div/div/div[3]/div/form/div/div[2]/div[2]/div/input'
        div_login = '/html/body/div/div/div[3]/div/form/div/div[3]/div[1]/button'
        
        time.sleep(2)
        input_usuario = tipo_elemento(driver, div_usuario, 'existente')
        input_password = tipo_elemento(driver, div_password, 'existente')
        button_login = tipo_elemento(driver, div_login, 'existente')
        #time.sleep(2)
        input_usuario.send_keys(user)
        #time.sleep(2)
        input_password.send_keys(pwd)
        time.sleep(0.5)
        button_login.click()
        time.sleep(2)

    def opciones_div(valor, opcion):
        
        diccionario = {1: 'Client', 2:'Organization', 4: 'From Date', 5: 'To Date', 8: 'Sub Trade Channel'}

        return {'filtro':f"//div[contains(@class, 'filter-modal') or contains(@class, 'ant-modal')]//label[text()='{diccionario[valor]}']/..//input",
                'div_export_filtro':f'/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[1]/div/div/div[2]/div/button[{valor}]'}[opcion]


    def filtro_survey(driver):
        div_filtro_survey = opciones_div(1, 'div_export_filtro')
        boton_filtro_survey = tipo_elemento(driver, div_filtro_survey, 'clickable')
        boton_filtro_survey.click()

    def click_survey(driver):
        xpath_mng = "//div[contains(@data-menu-id, '/SurveyManagement')]"
        xpath_rvw = "//li[contains(@data-menu-id, '/SurveyManagement/SurveyReview')]"

        boton_survey_mng = tipo_elemento(driver, xpath_mng, 'clickable', timeout=30)
        boton_survey_mng.click()
        time.sleep(1)
        boton_survey_rvw = tipo_elemento(driver, xpath_rvw, 'clickable', timeout=30)
        boton_survey_rvw.click()
        filtro_survey(driver)
        

    pag_carga = "//div[contains(@class, 'ant-modal-content')]"
    div_excel = '/html/body/div[3]/div/ul/li[1]'
    #div_export = '/html/body/div[1]/div/div[1]/div[2]/button[4]'
    div_export = '/html/body/div/div/div[1]/div[2]/button[3]'
    div_filtro = "/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/div[2]/div/span/span"
    #div_input_filtro = '/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[3]/div/div[3]/div/div/div/div/div[1]/div/input[1]'
    div_input_filtro_multi = '/html/body/div/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[3]/div/div[3]/div/div/div/div/div[1]/div/input[2]'
    div_fila_css = "div.ag-row[row-index='0']"

    #NUEVOS PATHS
    #div_apply_filtro_survey = '/html/body/div[4]/div/div/div/div[1]/div/div[3]/button[2]'
    div_apply_filtro_survey = "//button[contains(@class, 'ant-btn') and .//span[text()='Apply Filters']]"

    
    load_dotenv(dotenv_path='credenciales.env')
    usuario_arca = os.getenv('usuario_arca')
    password = os.getenv('contraseña_arca')
    
    driver = None
    ruta_completa_descarga = ruta_descarga + canal
    band = True

    try:
        print(f"[{session_uid}] \nIniciando proceso de descarga...")
        
        options_driver = webdriver.ChromeOptions()
        options_driver.add_argument('--start-maximized')
        options_driver.add_argument("--disable-extensions")
        #options_driver.add_argument("--headless=new")
        #options_driver.add_argument("--window-size=1920,1080")
        options_driver.add_argument("--disable-gpu")
        options_driver.add_argument("--no-sandbox")
        options_driver.add_argument("--disable-dev-shm-usage")

        options_driver.add_experimental_option("prefs", {
            "download.default_directory": ruta_completa_descarga,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        driver = webdriver.Chrome(options=options_driver)
        driver.get(url_web)

        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": ruta_completa_descarga}
        )

        inicio_sesion(driver, usuario_arca, password)
        esperar_invisibilidad(driver, pag_carga, timeout=120)
        click_survey(driver)
        time.sleep(1)

        for i in range(2,9):

            if estado == "1":

                if i == 8:
                    div_org_c_store = opciones_div(i,'filtro')
                    input_org_c_tore = tipo_elemento(driver,div_org_c_store,'clickable')
                    input_org_c_tore.send_keys(opciones[i-7])
                    time.sleep(1)
                    input_org_c_tore.send_keys(Keys.ENTER)
                    time.sleep(1)

                elif i == 4 or i == 5:
                    div_fechas = opciones_div(i, 'filtro')
                    input_fecha = tipo_elemento(driver, div_fechas,'clickable')
                    input_fecha.send_keys(Keys.CONTROL + 'a')
                    time.sleep(1)
                    input_fecha.send_keys(Keys.CLEAR)
                    time.sleep(1)
                    input_fecha.send_keys(fechas[i - 4])
                    input_fecha.send_keys(Keys.ENTER)
                    time.sleep(1)

        # for i in range(2,9):

            if estado == "0":

                if i == 2 or i == 8:
                    div_org_c_store = opciones_div(i,'filtro')
                    input_org_c_tore = tipo_elemento(driver,div_org_c_store,'clickable')
                    input_org_c_tore.send_keys(opciones[math.floor(math.sqrt(i))-1])
                    time.sleep(2)
                    input_org_c_tore.send_keys(Keys.ENTER)
                    time.sleep(1)

                elif i == 4 or i == 5:
                    div_fechas = opciones_div(i, 'filtro')
                    input_fecha = tipo_elemento(driver, div_fechas,'clickable')
                    input_fecha.send_keys(Keys.CONTROL + 'a')
                    time.sleep(1)
                    input_fecha.send_keys(Keys.CLEAR)
                    time.sleep(1)
                    input_fecha.send_keys(fechas[i - 4])
                    input_fecha.send_keys(Keys.ENTER)
                    time.sleep(2)
        
        button_search = tipo_elemento(driver, div_apply_filtro_survey,'clickable')
        button_search.click()
        time.sleep(3)

        try:
            fila_aparece_1 = tipo_elemento_css(driver, div_fila_css, 'css', timeout=120)
            fila_aparece_1.text

        except:
            print("Error al encontrar la fila")
            return f"Fallo: {fechas[0]} - {fechas[1]} \n- No se encontraron resultados con los filtros aplicados."
        
        
        if estado ==  "0":

            #click al botón exportar
            div_export_survey = opciones_div(2, 'div_export_filtro')
            button_export_survey = tipo_elemento(driver, div_export_survey,'clickable')
            button_export_survey.click()
            time.sleep(1)

            #click opcion excel
            button_excel = tipo_elemento(driver,div_excel,'existente')
            button_excel.click()
            
            while band:

                listado = os.listdir(ruta_completa_descarga)
                es_excel = any(f.endswith(".XLSX") for f in listado)

                if len(listado) != 0 and es_excel:
                    band = False
                
                time.sleep(2)

            print("Se logró descargar listado \n")

        else: 

            sessions_str = ",".join(session_uid)

            filtro = driver.find_element(By.XPATH,div_filtro)
            driver.execute_script("arguments[0].click();", filtro)
            time.sleep(0.5)

            input_id_session = tipo_elemento(driver,div_input_filtro_multi,'clickable')
            time.sleep(0.5)
            #input_id_session.send_keys(Keys.CONTROL + 'a')
            #time.sleep(0.5)
            #input_id_session.send_keys(Keys.DELETE)
            #time.sleep(0.5)
            
            input_id_session.send_keys(sessions_str) 
            time.sleep(0.5)
            input_id_session.send_keys(Keys.ENTER)
            time.sleep(0.5)

            fila_aparece_2 = tipo_elemento_css(driver, div_fila_css, 'css', timeout=120)
            fila_aparece_2.text
            
            for i in range(len(session_uid)):

                iconos = tipo_elemento(driver,f"/html/body/div/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[3]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[{i+1}]/div[1]","clickable")
                iconos.click()
                time.sleep(0.5)
                
            time.sleep(8)

            cod_ventana_principal = driver.current_window_handle
            cod_ventanas = driver.window_handles

            #ventanas_listas = []

            #while True:

            #    for h in cod_ventanas:
            #        if h == cod_ventana_principal:
            #            continue

            #        driver.switch_to.window(h)

            #        try:
            #            WebDriverWait(driver, 10).until(
            #                lambda d: d.current_url != "about:blank"
            #            )

            #           WebDriverWait(driver, 10).until(
            #                lambda d: d.execute_script("return document.readyState") == "complete"
            #            )

            #            ventanas_listas.append(h)

            #        except TimeoutException:
            #            print(f"Ventana {h} no lista, se ignora")
                
            #    if len(ventanas_listas) == len(session_uid):

            #        break

            if len(cod_ventanas) > 1:

                #nueva_ventana = [h for h in cod_ventanas if h != cod_ventana_principal][0]

                for h in cod_ventanas:

                    if h != cod_ventana_principal:

                        driver.switch_to.window(h)

                        esperar_invisibilidad(driver, pag_carga, timeout=60)
                        time.sleep(0.5)
                        
                        button_export = tipo_elemento(driver, div_export,'clickable')
                        button_export.click()
                        
                        time.sleep(0.5)
                
                for h in cod_ventanas:

                    if h != cod_ventana_principal:    
                        
                        driver.switch_to.window(h)

                        esperar_invisibilidad(driver, pag_carga, timeout=60)
                        time.sleep(0.5)

                        driver.close()

                driver.switch_to.window(cod_ventana_principal)
                time.sleep(0.5)
                
                #print(f"[{session_uid}] \nDescarga finalizada exitosamente.")
                return f"Éxito: {session_uid}"
            else:
                print(f"[{session_uid}] \nERROR: No se abrió la ventana secundaria.")
                return f"Fallo: {session_uid} \n- No se abrió la ventana secundaria."

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

    #lista_opciones = ['AUTOSERVICIO', 'C-STORE']#, 'AUTOAUDITORIA']
    lista_opciones = ['C-STORE']

    def eliminar_archivos_pasados(canal, incompleto = False):
        
        carpeta_incompletos = glob.glob(r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Session Incomplete\*')

        if canal == 'AUTOSERVICIO':    
            ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER\Downloads\AUTOSERVICIO\*')
            ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA2\*')
        elif canal == 'C-STORE':
            ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER\Downloads\C-STORE\*')
            ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA1\*' )
        #else:
        #    ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER\Downloads\AUTOAUDITORIA\*')
        #    ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA3\*' )
        else:
            ruta_canal = []
            ruta_descarga = []

        if not incompleto:
            archivos_totales = ruta_canal + ruta_descarga

        else:
            archivos_totales = carpeta_incompletos

        for f in archivos_totales:
            try:
                os.remove(f)

            except OSError:
                print("Error al borrar, es una subcarpeta!!")
        
        print(f"Se eliminaron {len(archivos_totales)} archivos")

#-----------------------------------------------------------------------------------------------------------------------------------------

    def mover_descargas(opcion):
        
        rt2 = rf'\{opcion}'
        rt1 = r"C:\Users\bbartolome.DICHTER\Downloads"
        ruta_origen = rt1 + rt2

        if opcion == 'C-STORE':
            ruta_destino = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA1'

        elif opcion == 'AUTOSERVICIO':
            ruta_destino = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA2'
            
        else:
            ruta_destino = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA3'
            

        patron_excel = 'Session_export*.xlsx'
        ruta_patron = os.path.join(ruta_origen, patron_excel)
        archivos_a_mover = glob.glob(ruta_patron)

        if not archivos_a_mover:
            print("No se encontraron archivos")
        else:
            print(f"{len(archivos_a_mover)} archivos encontrados.")
            
            if not os.path.exists(ruta_destino):
                os.makedirs(ruta_destino)
                print("carpeta creada")

            for ruta_origen_archivo in archivos_a_mover:
                
                nombre_archivo = os.path.basename(ruta_origen_archivo)
                ruta_destino_archivo = os.path.join(ruta_destino, nombre_archivo)
                
                try:
                    shutil.move(ruta_origen_archivo, ruta_destino_archivo)

                except Exception as e:
                    print(f"Error al mover {nombre_archivo}: {e}")

#-----------------------------------------------------------------------------------------------------------------------------------------

    def subir_excel(archivos_encontrados, patron):
        if not archivos_encontrados:
            print(f"No se encontró ningún archivo que coincida con el patrón: {patron}")
            return None
        else:
            archivo_a_cargar = archivos_encontrados[0]
            try:
                df = pd.read_excel(archivo_a_cargar)
                print("Archivo de UIDs maestro cargado exitosamente.")
                return df
            except Exception as e:
                print(f"Error al intentar cargar el archivo: {e}")
                return None

#-----------------------------------------------------------------------------------------------------------------------------------------

    def verificar_resultados(opcion, lista_uid, final = False):

        if opcion == 'C-STORE':
            ruta_carga = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA1'

        elif opcion == 'AUTOSERVICIO':
            ruta_carga = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA2'

        else:
            ruta_carga = r'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\CARPETA3'

        archivos_encontrados = glob.glob(os.path.join(ruta_carga, '*.xlsx'))
        
        if final:

            cant_descargados = len(archivos_encontrados)

            return cant_descargados
        
        dataframes = []

        for archivo in archivos_encontrados:

            Dataframe_resultados = pd.read_excel(archivo)
            dataframes.append(Dataframe_resultados)

        if dataframes:
            df_completo = pd.concat(dataframes, ignore_index= True)

        lista_resultados_uids = df_completo['SessionUID'].tolist()
        lista_faltantes = list((Counter(lista_uid) - Counter(lista_resultados_uids)).elements())

        return lista_faltantes
        
    #-----------------------------------------------------------------------------------------------------------------------------------------

    def dividir_en_chunks(lista, tamanio):
        for i in range(0,len(lista), tamanio):
            yield lista[i:i + tamanio]

    #-----------------------------------------------------------------------------------------------------------------------------------------

    MAX_PROCESOS = 5 # Número de navegadores/procesos concurrentes

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
                # chunks_uids,
                # [url_web] * len(lista_uids),
                # [ruta_descarga] * len(lista_uids),
                # [canal] * len(lista_uids),
                # [fechas] * len(lista_uids),
                # [opciones] * len(lista_uids),
                # ["1"] * len(lista_uids)
            )
            
            for resultado in resultados:
                print(f"Resultado final: {resultado}")
                
        print("\n--- 3. PROCESO PARALELO FINALIZADO ---\n")

#-----------------------------------------------------------------------------------------------------------------------------------------
    eliminar_archivos_pasados('NaN', True)

    for opcion in lista_opciones:

        # def Canal_Lindley():
            
        #     while True:
        #         print("\n--- MENÚ DE CANALES ---")
        #         print("1. Autoservicio")
        #         print("2. C-stores")
        #         print("3. Autoauditoria")
        #         print("0. Salir")
        #         print("------------------------")

        #         entrada_usuario = input("Ingrese su opción (0 para salir): ").strip()
                
        #         try:
        #             opcion = int(entrada_usuario)
                    
        #             if opcion == 0:
        #                 print("Saliendo del menú de canales.")
        #                 return None
        #             elif opcion == 1:
        #                 return "AUTOSERVICIO"
        #             elif opcion == 2:
        #                 return "C-STORE"
        #             elif opcion == 3:
        #                 return "AUTOAUDITORIA"
        #             else:
        #                 print(f"Opción inválida")
                
        #         except ValueError:
        #             print(f"Entrada inválida: No es un número. Por favor, ingrese un número.")

    #-----------------------------------------------------------------------------------------------------------------------------------------  
    #    VARIABLES
    #-----------------------------------------------------------------------------------------------------------------------------------------

        #opcion = Canal_Lindley()
        ruta_descarga = r'C:\Users\bbartolome.DICHTER\Downloads'
        canal = rf"\{opcion}"
        
        hoy = datetime.now()

        fecha_actual = hoy.strftime("%m/%d/%Y")

        #fechas = [fecha_actual, fecha_actual] #"mm/dd/yyyy"
        fechas = ['07/01/2026', '07/21/2026'] #"mm/dd/yyyy"
        opciones = ['',f'{opcion}']
        #opciones = [f'{opcion}']
        
        load_dotenv(dotenv_path='credenciales.env')
        url_web = os.getenv('ruta_web')

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
            lista_incompletos.to_excel(fr'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Session Incomplete\INCOMPLETOS_{opcion}.xlsx', index=False)

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


    # ahora = datetime.now()
    # hora = ahora.hour
    # minutos = ahora.minute

    # if (hora == 12 and minutos >= 30) or (hora == 15 and minutos >= 30) or (hora >= 19):

    #     try:
    #         subprocess.run(
    #             ["python", "datos.py"],
    #             cwd=r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Reporte_OT",
    #             check=True
    #         )
    #     except Exception as e:
    #         print(f"Error generando reportes: {e}")
    
    # else: 
    #     print("\n--- SE GENERARÁN LOS REPORTES DESPUÉS DE LAS 12:30 PM / 3:30 PM ---")

    
    # if (hora >= 19):
    #     print("\n--- INICIANDO CARGA A SHAREPOINT ---")
    #     try:
    #         subcarpetas = ["CARPETA1","CARPETA2"]

    #         for opc in subcarpetas:
    #             ejecucion(opc)
                
    #         carga_a_sharepoint()

    #         requests.post(
    #             "https://ntfy.sh/descarga_auto_chunk",
    #             data="Descarga Y carga a SharePoint terminaron",
    #             timeout=10
    #         )
    #     except Exception as e:
    #         print(f"[ERROR] Fallo la carga a SharePoint: {e}")
    
    # else:
    #     print("\n--- SE CARGARÁ A SHAREPOINT DESPUÉS DE LAS 8:00 PM ---")
