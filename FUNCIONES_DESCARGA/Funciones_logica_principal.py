import pandas as pd
import os
import glob
from dotenv import load_dotenv
import shutil
from collections import Counter

def eliminar_archivos_pasados(canal, incompleto = False):
    
    carpeta_incompletos = glob.glob(r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\VSCODE\Session Incomplete\*')

    if canal == 'AUTOSERVICIO':    
        ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER.000\Downloads\AUTOSERVICIO\*')
        ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA2\*')
    elif canal == 'C-STORE':
        ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER.000\Downloads\C-STORE\*')
        ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA1\*' )
    #else:
    #    ruta_canal = glob.glob(r'C:\Users\bbartolome.DICHTER.000\Downloads\AUTOAUDITORIA\*')
    #    ruta_descarga = glob.glob(r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA3\*' )
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
    rt1 = r"C:\Users\bbartolome.DICHTER.000\Downloads"
    ruta_origen = rt1 + rt2

    if opcion == 'C-STORE':
        ruta_destino = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA1'

    elif opcion == 'AUTOSERVICIO':
        ruta_destino = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA2'
        
    else:
        ruta_destino = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA3'
        

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
        ruta_carga = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA1'

    elif opcion == 'AUTOSERVICIO':
        ruta_carga = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA2'

    else:
        ruta_carga = r'C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\ESCRITORIO\CODBARRA\CARPETA3'

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