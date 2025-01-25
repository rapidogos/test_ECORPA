
import pyautogui    
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time

def initialize_driver():
    # Crea una instancia del navegador
    driver = webdriver.Chrome()
    return driver

def main():
    driver = initialize_driver()
    driver.get("https://www.dane.gov.co/index.php/estadisticas-por-tema/precios-y-costos/precios-de-venta-al-publico-de-articulos-de-primera-necesidad-pvpapn")
    # Espera a que la página cargue completamente
    time.sleep(3)  # Puedes usar WebDriverWait en lugar de sleep para una espera más controlada

    # Encuentra el enlace con el atributo 'title' específico
    enlace = driver.find_element(By.XPATH, '//*[@title="Anexo referencias mas vendidas"]')

    # Realiza un clic sobre el enlace
    ActionChains(driver).move_to_element(enlace).click().perform()

    time.sleep(4)

    # Maximizar la ventana para asegurarte de que todo se vea
    driver.maximize_window()
    # Tomar la captura de pantalla y guardarla
    screenshot_path = 'captura_pantalla.png'
    pyautogui.screenshot(screenshot_path)

    print(f"Captura de pantalla guardada en: {screenshot_path}")


if __name__ == '__main__':
    main()

