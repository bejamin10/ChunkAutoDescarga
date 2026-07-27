import pandas as pd
import time
import os
import glob
import shutil
import sys
import warnings
import numpy as np

#1. Subir las mediciones
#2. Consolidar mediciones
#3. Lista = [distinct(hoja1['SessionUid]),distinct(hoja2['SessionUid]), distinct(hoja3['SessionUid']), ...]
#4. band = False
#5. for i in range(0,len(Lista)):  if len(Lista[0]) == len(Lista[i+1]): band = True
#6. return band


def buscar_faltantes(lista_df):

    sesiones = []

    for i, df in enumerate(lista_df):

        col = 'SessionUID' if 'SessionUID' in df.columns else 'SessionUId'

        sesiones.append(set(df[col].dropna().unique()))

    sesiones_totales = set.union(*sesiones)

    faltantes = []

    for sesiones_df in sesiones:

        faltantes_df = sesiones_totales - sesiones_df

        if faltantes_df:
            faltantes.append(sorted(faltantes_df))


    return faltantes

def eliminar_faltantes(lista_df, faltantes):

    for i, df in enumerate(lista_df):

        try:
            col = 'SessionUID' if 'SessionUID' in df.columns else 'SessionUId'

            lista_df[i] = df[~df[col].isin(faltantes)]

        except Exception as e:
            print(f'Error en dataframe {i}: {e}')

    return lista_df