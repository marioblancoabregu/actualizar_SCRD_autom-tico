# actualizar_SCRD_automatico
Descripción:
Este proyecto automatiza la actualización de múltiples archivos Excel mediante Python reduciendo la intervención manual en procesos repetitivos y mejorando la eficiencia operativa.

Problema:
En entornos empresariales la actualización de múltiples reportes en Excel puede ser un proceso manual, repetitivo y propenso a errores especialmente cuando se trabaja con múltiples archivos que requieren actualización periódica.

Solución:
Se desarrolló un script en Python que:
-Abre múltiples archivos Excel automáticamente
-Ejecuta la función "Actualizar todo" (Refresh All)
-Gestiona tiempos de espera según el tipo de archivo
-Verifica el estado de cálculo de Excel
-Guarda y cierra los archivos automáticamente

Tecnologías utilizadas:
-Python
-win32com.client (automatización de Excel)
-Manejo de procesos y tiempos (time, os)

Cómo usar:
-Configurar la lista de archivos en el script:
archivos_excel = [
    r"C:\ruta\reportes\archivo_1.xlsx",
    r"C:\ruta\reportes\archivo_2.xlsx"
]
-Ejecutar el script:
python main.py

Resultados:
-Automatización de tareas manuales repetitivas
-Reducción significativa de tiempos operativos
-Mejora en la consistencia de los reportes
-Disminución de errores humanos

🔒 Nota:
Las rutas, nombres de archivos y datos han sido anonimizados para proteger la confidencialidad de la información.

🎯 Próximas mejoras:
-Integración con programación de tareas (Task Scheduler)
-Logging de ejecuciones
-Manejo de errores más robusto
-Integración con bases de datos

👨‍💻 Autor:
Mario Blanco
