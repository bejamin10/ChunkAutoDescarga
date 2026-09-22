from dotenv import load_dotenv
import requests
import os
from datetime import datetime

def carga_a_sharepoint():

    carpeta_AASS = r"C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\NUEV ESCRITORIO\VSCODE\AutoDescargaChunk\PARQUETS AUTOSERVICIO"
    carpeta_CSTORES = r"C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\NUEV ESCRITORIO\VSCODE\AutoDescargaChunk\PARQUETS C-STORES"
    carpeta_REPORTES = r"C:\Users\bbartolome.DICHTER.000\OneDrive - Dichter & Neira\NUEV ESCRITORIO\VSCODE\Reporte_OT\output"

    dia = datetime.now().day
    dia_str = f"{dia:02d}"
    fecha = datetime.now().month
    mes_num = f"{fecha:02d}"
    meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    mes_str = meses[datetime.now().month]

    load_dotenv(dotenv_path='credenciales.env')

    TENANT_ID_NUEVO = os.getenv('tenant_id_sharepoint')
    CLIENT_ID_NUEVO = os.getenv('client_id_sharepoint')
    CLIENT_SECRET_NUEVO = os.getenv('client_secret_sharepoint')

    #-----------------------------
    #-------NUEVO SHAREPOINT------
    #-----------------------------

    url_s = f"https://login.microsoftonline.com/{TENANT_ID_NUEVO}/oauth2/v2.0/token"

    data_s = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID_NUEVO,
        "client_secret": CLIENT_SECRET_NUEVO,
        "scope": "https://graph.microsoft.com/.default"
    }

    response_s = requests.post(url_s, data=data_s)

    print("STATUS TOKEN:", response_s.status_code)
    #print("RESPONSE TOKEN:", response.text)

    token_s = response_s.json().get("access_token")

    if not token_s:
        raise Exception("No se pudo obtener token")

    headers_s = {
        "Authorization": f"Bearer {token_s}"
    }

    res_s = requests.get(
        "https://graph.microsoft.com/v1.0/sites?search=Lock_LindleyPeru",
        headers=headers_s
    )

    data_1_s = res_s.json()

    for site in data_1_s["value"]:
        #print(site["name"], "->", site["id"])
        if site["name"] == "Lock_LindleyPeru":
            site_id_s = site["id"]

    res_s = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id_s}/drives", headers=headers_s)
    data_2_s = res_s.json()

    for drive in data_2_s["value"]:
        #print(drive["name"], "->", drive["id"])
        if drive["name"] == "PBI Data Moderno":
            drive_id_s = drive["id"]


    def crearCarpetasNueva(drive_id_s, carpeta_remota, headers_s, dia_str, tipo, mes_num, mes_str):

        carpeta = dia_str if tipo == 'R' else mes_str

        url_base = f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{carpeta_remota}"
                                                                                                      
        if tipo == 'R':
            url = f"{url_base}/2026/{mes_num}. {mes_str}:/children"
            carpeta = dia_str

        elif tipo == 'M' and dia_str == "01":    #Solo crea la carpeta mensual si es el primer día del mes,                                                
            url = f"{url_base}/2026:/children"   #si no verificamos el dia, haria una verificacion de existencia de carpeta mensual todos los días y no es necesario.
            carpeta = mes_str

        else: 
            print("No se creará carpeta. No es 01 del mes o no es tipo 'R' o 'M'.")
            return

        # NOTA CORREGIDA: 
        # Para crear carpetas mediante POST, microsoft graph exige que la URL termine estrictamente en ':/children'.
        # o sea que la URL solo define la carpeta padre donde queremos guardar las cosas.
        # aqui el campo 'name' dentro del body no pasa a segundo plano; es el unico que define el nombre
        # de la nueva carpeta que se va a crear ('R' o 'M').

        body = {
            "name": carpeta, 
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        }

        res_s = requests.post(url, headers=headers_s, json=body)

        if res_s.status_code == 201:
            print(f"Carpeta '{carpeta}' creada correctamente en {carpeta_remota}.")

        elif res_s.status_code == 409:
            print(f"La carpeta '{carpeta}' ya existe en {carpeta_remota}.")

        else:
            print(f"Error creando carpeta en sharepoint nuevo: {mes_str}/{res_s.status_code}")
            print(res_s.text)     


    def eliminarArchivos(drive_id_s, ruta_carpeta, mes_str, headers_s):

        url_listar = f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{ruta_carpeta}/2026/{mes_str}:/children" #hijos
        res_listar = requests.get(url_listar, headers=headers_s)

        if res_listar.status_code != 200:
            print(f"Error al obtener el contenido de '{ruta_carpeta}': {res_listar.status_code}")
            print(res_listar.text)
            return 

        elementos = res_listar.json().get("value", [])

        if not elementos:
            print(f"La carpeta '{ruta_carpeta}' está vacía o no existe.")
            return

        print(f"Iniciando eliminación en: {ruta_carpeta}")
        eliminados = 0

        for item in elementos: #elimino por elemento
            item_id = item["id"]
            nombre = item["name"]

            url_eliminar = f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/items/{item_id}"
            res_delete = requests.delete(url_eliminar, headers=headers_s)

            if res_delete.status_code == 204:
                print(f" Eliminado: {nombre}")
                eliminados += 1
            else:
                print(f" Error al eliminar '{nombre}': {res_delete.status_code}")

        print(f"Total de elementos eliminados: {eliminados}\n")


    #Flujo principal
    def cargarArchivosSharepointNuevo(carpeta_local, drive_id_s, carpeta_remota, tipo):

        crearCarpetasNueva(drive_id_s, carpeta_remota, headers_s, dia_str, tipo, mes_num,mes_str)

        if tipo == 'M':                   
            eliminarArchivos(drive_id_s, carpeta_remota, mes_str, headers_s)

        dic_ruta = {"R": f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{carpeta_remota}/2026/{mes_num}. {mes_str}/{dia_str}/",
                    "M": f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{carpeta_remota}/2026/{mes_str}/"}
        
        print(carpeta_remota)
        i=0
        
        for archivo in os.listdir(carpeta_local):
            if archivo.endswith((f"{dia_str}.{mes_num}.parquet", f"{dia_str}.{mes_num}-Cierre.pdf")):
                ruta_archivo = os.path.join(carpeta_local, archivo)

                with open(ruta_archivo, "rb") as f:
                    contenido = f.read()
            
                upload_url = dic_ruta[tipo] + f"{archivo}:/content"
                res = requests.put(upload_url, headers=headers_s, data=contenido)
                
                print(f"{archivo} -> {res.status_code}")
                i = i + 1

        print(f"Archivos evaluados: {i}")
        print("\n")

    cargarArchivosSharepointNuevo(carpeta_AASS, drive_id_s, "Data Autoservicios","M")
    cargarArchivosSharepointNuevo(carpeta_CSTORES, drive_id_s, "Data C-Stores","M")
    cargarArchivosSharepointNuevo(carpeta_REPORTES, drive_id_s, "Reporte OT's","R")

    print("\nCarga a SharePoint finalizada.")
    
if __name__ == "__main__":
    carga_a_sharepoint()