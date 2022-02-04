from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from tkinter import messagebox

import time
# from config import URL_AFIP
from functions_GCA.config import URL_AFIP

def realizarFactura(driver, cuil, clave, empresa, fechaComprobanteHASTA, periodoFacturadoDESDE, cuilReceptor, precio, servicio):
    # Launch URL and maximize window
    driver.get(URL_AFIP)
    driver.maximize_window()
    # Identify CUIL User box and write CUIL
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"F1:username\"]"))).send_keys(cuil)
    # Click Submit CUIL / Siguiente
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"F1:btnSiguiente\"]"))).click()
    # Wait for TU CLAVE text
    WebDriverWait(driver, 32).until(EC.text_to_be_present_in_element((By.XPATH, "//*[@id=\"F1\"]/div/label"), "TU CLAVE"))
    # Identify CLAVE User box and write CLAVE
    driver.find_element(By.XPATH, "//*[@id=\"F1:password\"]").send_keys(clave)
    # Click Ingresar and wait for the link to change
    driver.find_element(By.XPATH, "//*[@id=\"F1:btnIngresar\"]").click()
    WebDriverWait(driver, 32).until(EC.url_contains("https://serviciosjava2.afip.gob.ar/rcel/jsp/index_bis.jsp;jsessionid="))
    # Wait for the Empresa button that matches this function argument to appear and click it
    WebDriverWait(driver, 32).until(EC.text_to_be_present_in_element((By.XPATH, "//*[@id=\"contenido\"]/form/table/tbody/tr[2]/td/b"), "Seleccione la Empresa a representar:"))
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@value='{}'and @type='button']".format(empresa)))).click()
    # Wait for the Generar Comprobantes text to appear, and then it clicks it
    WebDriverWait(driver, 32).until(EC.text_to_be_present_in_element((By.XPATH, "//*[@id=\"btn_gen_cmp\"]/span[2]"), "Generar Comprobantes"))
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//span[@class='ui-button-text' and text()=\"Generar Comprobantes\"]"))).click()
    # Wait for "Puntos de Ventas y Tipos de Comprobantes habilitados para impresión" to appear on the screen, then selects first value from dropdown
    WebDriverWait(driver, 32).until(EC.text_to_be_present_in_element((By.XPATH, "//*[@id=\"contenido\"]/div[2]"), "Puntos de Ventas y Tipos de Comprobantes habilitados para impresión"))
    Select(driver.find_element(By.ID, "puntodeventa")).select_by_value('1')
    time.sleep(0.5)
    # Waits for and clicks "Continuar", then it waits for the "Moneda Extranjera" checkbox
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Continuar >'and @type='button']"))).click()
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@name='monedaExtranjera'and @type='checkbox']")))
    # Writes fechaComprobanteHASTA
    comprobanteDate = WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"fc\"]")))
    comprobanteDate.clear()
    comprobanteDate.send_keys(fechaComprobanteHASTA)
    Select(driver.find_element(By.XPATH, "//*[@id=\"idconcepto\"]")).select_by_visible_text("\u00A0Servicios")
    desde_PF = WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"fsd\"]")))
    desde_PF.clear()
    desde_PF.send_keys(periodoFacturadoDESDE)
    hasta_FC = driver.find_element(By.XPATH, "//*[@id=\"fsh\"]")
    hasta_FC.clear()
    hasta_FC.send_keys(fechaComprobanteHASTA)
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Continuar >'and @type='button']"))).click()
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"email\"]")))
    Select(driver.find_element(By.XPATH, "//*[@id=\"idivareceptor\"]")).select_by_visible_text("\u00A0Consumidor Final")
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"nrodocreceptor\"]"))).send_keys(cuilReceptor)
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"formadepago7\"]"))).click()
    time.sleep(0.8)
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Continuar >'and @type='button']"))).click()
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"idoperacion\"]/tbody/tr[2]/td[1]/input[1]")))
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"detalle_cantidad1\"]"))).send_keys("1")
    Select(driver.find_element(By.XPATH, "//*[@id=\"detalle_medida1\"]")).select_by_visible_text(" unidades")
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"detalle_precio1\"]"))).send_keys(precio)
    WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"detalle_descripcion1\"]"))).send_keys(servicio)
    # WebDriverWait(driver, 32).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Continuar >'and @type='button']"))).click()
    messagebox.showwarning("Se termino el proceso de Realizar Factura", "Asegurese de antes de cerrar este mensaje, interactuar con la ventana de Safari y elegir la opción de Detener Sesión")
    

if __name__ == '__main__':
    list_DatosParaComprobante = [ 
        webdriver.Safari(), # Using a Safari driver
        "20123456789", # cuil Example
        "ejemploClave", # clave Example
        "APELLIDOS NOMBRES", # empresa Example
        "24/01/2022", # fechaComprobanteHASTA Example
        "17/01/2022", # periodoFacturadoDESDE Example
        "20987654321", # cuilReceptor Example
        "3200", # precio Example
        "Consulta de Test" # servicio Example
        ]
        
    realizarFactura(*list_DatosParaComprobante)
