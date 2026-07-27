import pandas as pd
import time
import os
import glob
import shutil
import sys
import warnings
import numpy as np
from datetime import datetime
from Integridad_data import buscar_faltantes, eliminar_faltantes
import requests

ahora = datetime.now()
hora = ahora.hour
fecha_str = ahora.strftime("%d.%m")

def barra_de_carga(actual, total, largo_barra=40, mensaje="Procesando..."):
    porcentaje_completo = (actual / total)
    num_caracteres_llenos = int(porcentaje_completo * largo_barra)
    caracter_lleno = '█'
    caracter_vacio = '-'

    barra = (caracter_lleno * num_caracteres_llenos) + (caracter_vacio * (largo_barra - num_caracteres_llenos))

    texto_salida = f"{mensaje} |{barra}| {int(porcentaje_completo * 100)}% ({actual}/{total})"
    
    print(texto_salida, end='\r', file=sys.stdout)


def subir_excel(hoja, sub_carpeta):
    carpeta = fr'C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\CODBARRA\{sub_carpeta}'
    archivos = glob.glob(os.path.join(carpeta, '*.xlsx'))
    dataframes = []
    total_archivos = len(archivos)

    if not archivos:    
        print(f"Carpeta {sub_carpeta} vacia, no se hace carga para {hoja}.")
        return None

    print(f"\n--- Iniciando consolidación para la hoja: **{hoja}")

    for i, archivo in enumerate(archivos):
        try:
            df = pd.read_excel(archivo, sheet_name=hoja)
            df = df.astype(str)
            #if not df.empty:
            nombre_archivo = os.path.basename(archivo)
            df["NOMBRE ARCHIVO"] = nombre_archivo
            
            dataframes.append(df)
            barra_de_carga(actual=i + 1, total=total_archivos, mensaje=f"Leyendo archivos de '{hoja}'")
        except Exception as e:
            print(f"Error: {e}")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", FutureWarning)
        
        if dataframes:
            df_completo = pd.concat(dataframes, ignore_index=True)
            
        else:
            df_completo = pd.DataFrame(dataframes)

    return df_completo


def ejecucion(sub_carpeta):

    hojas = ['IR_Actual', 'IR_InventoryPricing', 'IR_Metrics', 'IR_MQ', 'IR_Scenes', 'IR_Sessions']
    lista_df = []

    for hoja in hojas:
        df = subir_excel(hoja, sub_carpeta)
        lista_df.append(df)

    faltantes = buscar_faltantes(lista_df)
    print("\n")
    print(len(lista_df))
    print(hora)
    if len(faltantes) > 0 and hora >= 19:
        print("Estoy dentro de faltantes")
        print(f"faltante {faltantes}")
        requests.post(
            "https://ntfy.sh/Evaluacion_parquets",
            data=f"Hay sesssion Uid faltantes en {sub_carpeta}: {faltantes}",
            timeout=10
        )

        lista_df = eliminar_faltantes(lista_df, faltantes)

        for df, hoja in zip(lista_df,hojas):
            if df is not None and not df.empty:

                if sub_carpeta == "CARPETA2":
                    df.to_parquet(fr'.\PARQUETS AUTOSERVICIO\{hoja}-{fecha_str}.parquet', index = False)
                else:
                    df.to_parquet(fr'.\PARQUETS C-STORES\{hoja}-{fecha_str}.parquet', index = False)
                time.sleep(4)  
    
    else:  

        for df, hoja in zip(lista_df,hojas):
            if df is not None and not df.empty:

                if sub_carpeta == "CARPETA2":
                    df.to_parquet(fr'.\PARQUETS AUTOSERVICIO\{hoja}-{fecha_str}.parquet', index = False)
                else:
                    df.to_parquet(fr'.\PARQUETS C-STORES\{hoja}-{fecha_str}.parquet', index = False)
                time.sleep(4)  

if __name__ == '__main__':
    #subcarpetas = ["CARPETA1","CARPETA2"]
    subcarpetas = ["CARPETA1"]
    for opc in subcarpetas:
        ejecucion(opc)