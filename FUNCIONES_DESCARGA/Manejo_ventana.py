import time
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

def esperar_invisibilidad(driver,ruta , timeout=120):
    wait = WebDriverWait(driver, timeout)
    
    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "anticon-loading")))
    
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ant-modal-wrap")))
        
        return True
    
    except TimeoutException:
        print("El modal no desapareció correctamente")
        return False

#--------------------------------------------------------------------------------------------------

def tipo_elemento(driver, ruta, elemento, timeout=60):
    diccionario = {
        'clickable': WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, ruta))),
        'existente': WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, ruta)))
    }
    return diccionario[elemento]

#--------------------------------------------------------------------------------------------------

def tipo_elemento_css(driver, ruta, elemento, timeout=120):
        diccionario = {
            'css': WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, ruta)))
        }
        return diccionario[elemento]

#--------------------------------------------------------------------------------------------------

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

#--------------------------------------------------------------------------------------------------

def opciones_div(valor, opcion):
    
    diccionario = {1: 'Client', 2:'Organization', 4: 'From Date', 5: 'To Date', 8: 'Sub Trade Channel'}

    return {'filtro':f"//div[contains(@class, 'filter-modal') or contains(@class, 'ant-modal')]//label[text()='{diccionario[valor]}']/..//input",
            'div_export_filtro':f'/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[1]/div/div/div[2]/div/button[{valor}]'}[opcion]

#--------------------------------------------------------------------------------------------------

def filtro_survey(driver):
    div_filtro_survey = opciones_div(1, 'div_export_filtro')
    boton_filtro_survey = tipo_elemento(driver, div_filtro_survey, 'clickable')
    boton_filtro_survey.click()

#--------------------------------------------------------------------------------------------------    

def click_survey(driver):
    xpath_mng = "//div[contains(@data-menu-id, '/SurveyManagement')]"
    xpath_rvw = "//li[contains(@data-menu-id, '/SurveyManagement/SurveyReview')]"

    boton_survey_mng = tipo_elemento(driver, xpath_mng, 'clickable', timeout=30)
    boton_survey_mng.click()
    time.sleep(1)
    boton_survey_rvw = tipo_elemento(driver, xpath_rvw, 'clickable', timeout=30)
    boton_survey_rvw.click()
    filtro_survey(driver)


#--------------------------------------------------------------------------------------------------