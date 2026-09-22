import requests
import os

TOKEN = "8716177913:AAEYrNj8zb3-yk1gpQQlvBtCj5NDaQKSxO0"
CHAT_ID = "7928253154"                     

ruta = [r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Reporte_OT\output\reporte_aass_11.06-cierre.pdf",
        r"C:\Users\bbartolome.DICHTER\OneDrive - Lock & Asociados\ESCRITORIO\VSCODE\Reporte_OT\output\reporte_cstores_11.06-cierre.pdf"]

for r in ruta:
    with open(r, "rb") as f:
        response = requests.post(
            f"https://api.telegram.org/bot{TOKEN}/sendDocument",
            data={"chat_id": CHAT_ID},
            files={"document": (os.path.basename(r), f, "application/pdf")}
    )

print(response.status_code, response.json())