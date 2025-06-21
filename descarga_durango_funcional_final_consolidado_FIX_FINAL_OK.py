from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import shutil
from datetime import datetime

# Ruta de descarga
from pathlib import Path
download_dir = str(Path.home() / "Downloads")

options = webdriver.ChromeOptions()
prefs = {"download.default_directory": download_dir}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get("https://nube.agricultura.gob.mx/avance_agricola/")
print("🌐 Abriendo página de Avance Agrícola...")
print("⏳ Esperando 15 segundos para que desaparezca el overlay...")
time.sleep(15)

print("✅ Navegador lanzado y página cargada.")

try:
    print("🌱 Seleccionando 'Ciclo'...")
    Select(driver.find_element(By.ID, "cicloProd")).select_by_visible_text("Año Agrícola (OI + PV)")
    time.sleep(3)
    print("✅ Ciclo seleccionado.")

    print("💧 Seleccionando 'Modalidad'...")
    Select(driver.find_element(By.ID, "modalidad")).select_by_visible_text("Riego + Temporal")
    time.sleep(3)
    print("✅ Modalidad seleccionada.")

    print("📍 Seleccionando 'Entidad Federativa'...")
    Select(driver.find_element(By.ID, "entidad")).select_by_visible_text("Durango")
    time.sleep(3)
    print("✅ Entidad 'Durango' seleccionada.")

    print("🌾 Seleccionando 'Resumen cultivos'...")
    Select(driver.find_element(By.ID, "cultivo")).select_by_visible_text("Resumen cultivos")
    time.sleep(3)
    print("✅ Cultivo 'Resumen cultivos' seleccionado.")
except Exception as e:
    print("❌ Error durante la automatización inicial:", e)
    driver.quit()
    exit()

meses = {
    "1": "enero", "2": "febrero", "3": "marzo", "4": "abril", "5": "mayo", "6": "junio",
    "7": "julio", "8": "agosto", "9": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre"
}

for anio in range(2018, 2024):
    Select(driver.find_element(By.ID, "anioagric")).select_by_visible_text(str(anio))
    print(f"📅 Año {anio} seleccionado.")
    time.sleep(3)

    for num in range(1, 13):
        mes_num = str(num)
        mes_nombre = meses[mes_num]

        print(f"📅 Procesando {mes_nombre} {anio}...")
        try:
            Select(driver.find_element(By.ID, "mesagric")).select_by_visible_text(mes_nombre.capitalize())
            time.sleep(3)
            driver.find_element(By.XPATH, '//button[.//*[contains(text(),"Consultar")]]').click()
            time.sleep(3)
            driver.find_element(By.XPATH, '//button[.//*[contains(text(),"Generar")]]').click()
            print(f"⬇️ Descargando {mes_nombre}_{anio}_durango.xls...")
            time.sleep(5)

            # Renombrar archivo más reciente descargado
            archivos = [f for f in os.listdir(download_dir) if f.endswith(".xls")]
            archivos.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
            if archivos:
                archivo_reciente = archivos[0]
                nuevo_nombre = f"{mes_nombre}_{anio}_durango.xls"
                shutil.move(os.path.join(download_dir, archivo_reciente), os.path.join(download_dir, nuevo_nombre))
                print(f"✅ Archivo renombrado: {nuevo_nombre}")
            else:
                print("⚠️ No se encontró archivo para renombrar.")
        except Exception as e:
            print(f"❌ Error al procesar {mes_nombre}/{anio}:", e)

print("✅ Proceso completado.")
input("Presiona Enter para cerrar...")
driver.quit()
