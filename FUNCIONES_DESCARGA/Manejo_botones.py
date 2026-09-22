import time
import os
import math

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from FUNCIONES_DESCARGA.Manejo_ventana import * 

def manejo_botones(driver, estado, opciones, fechas, div_apply_filtro_survey, div_fila_css, div_excel, ruta_completa_descarga, session_uid, div_filtro, div_input_filtro_multi, pag_carga, div_export, band):

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

        if len(cod_ventanas) > 1:

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
            
            return f"Éxito: {session_uid}"
        else:
            print(f"[{session_uid}] \nERROR: No se abrió la ventana secundaria.")
            return f"Fallo: {session_uid} \n- No se abrió la ventana secundaria."

