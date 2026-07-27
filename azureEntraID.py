from dotenv import load_dotenv
import requests
import os
from datetime import datetime

def carga_a_sharepoint():
    carpeta_AASS = r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\AutoDescargaChunk\PARQUETS AUTOSERVICIO"
    carpeta_CSTORES = r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\AutoDescargaChunk\PARQUETS C-STORES"
    carpeta_REPORTES = r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Reporte_OT\output"

    dia = datetime.now().day
    dia_str = f"{dia:02d}"
    fecha = datetime.now().month
    fecha_str = f"{fecha:02d}"

    load_dotenv(dotenv_path='credenciales.env')
    #TENANT_ID = os.getenv('tenant_id')
    #CLIENT_ID = os.getenv('client_id')
    #CLIENT_SECRET = os.getenv('client_secret')

    TENANT_ID_NUEVO = os.getenv('tenant_id_sharepoint')
    CLIENT_ID_NUEVO = os.getenv('client_id_sharepoint')
    CLIENT_SECRET_NUEVO = os.getenv('client_secret_sharepoint')


    #-------------------------------
    #-------ANTIGUO SHAREPOINT------
    #-------------------------------

    # url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    # data = {
    #     "grant_type": "client_credentials",
    #     "client_id": CLIENT_ID,
    #     "client_secret": CLIENT_SECRET,
    #     "scope": "https://graph.microsoft.com/.default"
    # }

    # response = requests.post(url, data=data)

    # print("STATUS TOKEN:", response.status_code)
    # #print("RESPONSE TOKEN:", response.text)

    # token = response.json().get("access_token")

    # if not token:
    #     raise Exception("No se pudo obtener token")

    # headers = {
    #     "Authorization": f"Bearer {token}"
    # }

    # res = requests.get(
    #     "https://graph.microsoft.com/v1.0/sites?search=lockyasociados",
    #     #"https://graph.microsoft.com/v1.0/sites?search=Lock_LindleyPeru",
    #     headers=headers
    # )

    # data_1 = res.json()

    # for site in data_1["value"]:
    #     #print(site["name"], "->", site["id"])
    #     if site["name"] == "Lock_LindleyPeru":
    #         site_id = site["id"]

    # res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives", headers=headers)
    # data_2 = res.json()

    # for drive in data_2["value"]:
    #     #print(drive["name"], "->", drive["id"])
    #     if drive["name"] == "PBI Data Moderno":
    #         drive_id = drive["id"]


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
        #"https://graph.microsoft.com/v1.0/sites?search=lockyasociados",
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


    # def crearCarpetasAntiguo(drive_id, carpeta_remota, headers, dia_str):

    #     parent_path = f"{carpeta_remota}/2026/06. Junio"

    #     url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}:/children"

    #     body = {
    #         "name": dia_str,
    #         "folder": {},
    #         "@microsoft.graph.conflictBehavior": "fail"
    #     }

    #     res = requests.post(url, headers=headers, json=body)

    #     if res.status_code == 201:
    #         print(f"Carpeta '{dia_str}' creada correctamente en sharepoint antiguo.")

    #     elif res.status_code == 409:
    #         print(f"La carpeta '{dia_str}' ya existe en sharepoint antiguo.")

    #     else:
    #         print(f"Error creando carpeta en sharepoint antiguo: {res.status_code}")
    #         print(res.text)


    def crearCarpetasNueva(drive_id_s, carpeta_remota, headers_s, dia_str):

        parent_path = f"{carpeta_remota}/2026/06. Junio"

        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{parent_path}:/children"

        body = {
            "name": dia_str,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        }

        res_s = requests.post(url, headers=headers_s, json=body)

        if res_s.status_code == 201:
            print(f"Carpeta '{dia_str}' creada correctamente en sharepoint nuevo.")

        elif res_s.status_code == 409:
            print(f"La carpeta '{dia_str}' ya existe en sharepoint nuevo.")

        else:
            print(f"Error creando carpeta en sharepoint nuevo: {res_s.status_code}")
            print(res_s.text)     


    # def cargarArchivosSharepointAntiguo(carpeta_local, drive_id, carpeta_remota, tipo):

    #     if tipo == "R" or tipo == "R2":
    #         crearCarpetasAntiguo(drive_id, carpeta_remota, headers, dia_str)

    #     dic_ruta = {"M": f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{carpeta_remota}/2026/Junio/",
    #                 "R": f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{carpeta_remota}/2026/06. Junio/{dia_str}/"}
                    

    #     print(carpeta_remota)
    #     i=0
        
    #     for archivo in os.listdir(carpeta_local):
    #         if archivo.endswith((".xlsx", f"{dia_str}.{fecha_str}-Cierre.pdf")):
    #             ruta_archivo = os.path.join(carpeta_local, archivo)

    #             with open(ruta_archivo, "rb") as f:
    #                 contenido = f.read()
                
    #             #upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/Mediciones Diario/2026/Marzo/31/{archivo}:/content"
    #             upload_url = dic_ruta[tipo] + f"{archivo}:/content"
    #             res = requests.put(upload_url, headers=headers, data=contenido)
                
    #             print(f"{archivo} -> {res.status_code}")
    #             i = i + 1

    #     print(f"Archivos evaluados: {i}")
    #     print("\n")

    def cargarArchivosSharepointNuevo(carpeta_local, drive_id_s, carpeta_remota, tipo):

        if tipo == "R" :
            crearCarpetasNueva(drive_id_s, carpeta_remota, headers_s, dia_str)

        dic_ruta = {"R": f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{carpeta_remota}/2026/06. Junio/{dia_str}/",
                    "M": f"https://graph.microsoft.com/v1.0/drives/{drive_id_s}/root:/{carpeta_remota}/2026/Junio/"}
                    

        print(carpeta_remota)
        i=0
        
        for archivo in os.listdir(carpeta_local):
            if archivo.endswith((f"{dia_str}.{fecha_str}.parquet", f"{dia_str}.{fecha_str}-Cierre.pdf")):
                ruta_archivo = os.path.join(carpeta_local, archivo)

                with open(ruta_archivo, "rb") as f:
                    contenido = f.read()
                
                #upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/Mediciones Diario/2026/Marzo/31/{archivo}:/content"
                upload_url = dic_ruta[tipo] + f"{archivo}:/content"
                res = requests.put(upload_url, headers=headers_s, data=contenido)
                
                print(f"{archivo} -> {res.status_code}")
                i = i + 1

        print(f"Archivos evaluados: {i}")
        print("\n")

    cargarArchivosSharepointNuevo(carpeta_AASS, drive_id_s, "Data Autoservicios","M")
    cargarArchivosSharepointNuevo(carpeta_CSTORES, drive_id_s, "Data C-Stores","M")
    #cargarArchivosSharepointAntiguo(carpeta_REPORTES, drive_id, "Reporte OT's","R")
    #cargarArchivosSharepointNuevo(carpeta_REPORTES, drive_id_s, "Reporte OT's","R")

    print("\nCarga a SharePoint finalizada.")
    
if __name__ == "__main__":
    carga_a_sharepoint()