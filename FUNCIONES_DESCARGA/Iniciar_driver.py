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


def iniciar_driver(session_uid, ruta_completa_descarga, url_web, driver):

    # try:
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

    return driver

    # except (WebDriverException, TimeoutException) as e:
    #     print(f"[{session_uid}] \nFALLO: Error de Selenium o Timeout: {e}")
    #     return f"Fallo: {session_uid} \n- Error de Selenium/Timeout."
        
    # except Exception as e:
    #     print(f"[{session_uid}] \nFALLO: Ocurrió un error inesperado: {e}")
    #     return f"Fallo: {session_uid} \n- Error inesperado."

    # finally:
    #     if driver:
    #         driver.quit()

