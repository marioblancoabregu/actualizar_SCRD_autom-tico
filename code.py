import os
import time
import win32com.client as win32

#Rutas archivos Excel:
archivos_excel = [
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL1.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL2.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL3.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL4.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL5.xlsx",   
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL6.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL7.xlsx",
    r"C:\Users\USUARIO\Desktop\ARCHIVO_EXCEL8.xlsx"
  ]

Tiempo entre archivos (asegurando el correcto funcionamiento de actualización y guardado de los archivos):
tiempo_espera = 10  # segundos

#Función principal por archivo:
def actualizar_archivo(ruta_archivo):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    print(f"\n Abriendo: {ruta_archivo}")
    wb = excel.Workbooks.Open(ruta_archivo)

    print("Ejecutando 'Actualizar todo'...")
    wb.RefreshAll()

    #Caso especial: "ARCHIVO_EXCEL1.xlsx" 
    #Este archivo en específico tarda en guardar las actualizaciones caussando demoras y errores, se procede a alargar el tiempo de espera:
    if "Fill Rate.xlsx" in ruta_archivo:
        print("Archivo 'ARCHIVO_EXCEL1.xlsx' detectado. Esperando solo 30 segundos...")
        time.sleep(30)
        print("Continuando con el flujo normal.")
    else:
        #Espera hasta 20 minutos para otros archivos
        print("Esperando que Excel termine de actualizar (máx. 20 minutos)...")
        max_espera = 1200
        esperado = 0
        intervalo = 5

        while excel.CalculationState != 0 and esperado < max_espera:
            time.sleep(intervalo)
            esperado += intervalo

        if esperado >= max_espera:
            print("Tiempo máximo alcanzado. Continuando de todos modos.")
        else:
            print("Actualización finalizada.")

    #Guardar y cerrar:
    wb.Save()
    wb.Close(SaveChanges=True)
    print("Archivo guardado y cerrado.")

    #Cerrar Excel si no quedan libros abiertos:
    if len(excel.Workbooks) == 0:
        excel.Quit()
        print("Excel cerrado.\n")

#Bucle principal:
for archivo in archivos_excel:
    if os.path.exists(archivo):
        actualizar_archivo(archivo)
        print("Esperando para el siguiente archivo...\n")
        time.sleep(tiempo_espera)
    else:
        print(f"No se encontró el archivo: {archivo}")
