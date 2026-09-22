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


sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from Reporte_OT.CARGAS.FUNCIONES.Funciones_Carga import subir_excel

ahora = datetime.now()
hora = ahora.hour
fecha_str = ahora.strftime("%d.%m")


def ejecucion(sub_carpeta):

    hojas = ['IR_Actual', 'IR_InventoryPricing', 'IR_Metrics', 'IR_MQ', 'IR_Scenes', 'IR_Sessions']
    lista_df = []

    for hoja in hojas:
        df = subir_excel(hoja, sub_carpeta, funcion="parquet")
        lista_df.append(df)

    faltantes = buscar_faltantes(lista_df)
    print("\n")
    print(len(lista_df))
    #print(hora)
    if len(faltantes) > 0 :

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

if __name__ == '__main__':
    #subcarpetas = ["CARPETA1","CARPETA2"]
    subcarpetas = ["CARPETA2"]
    for opc in subcarpetas:
        ejecucion(opc)