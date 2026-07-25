# logging_utils.py

"""
Módulo de Utilidades de Logging
===============================

Este script proporciona una configuración centralizada para el manejo de logs
en aplicaciones Python. Configura automáticamente un doble flujo de salida:
1.  **Consola:** Mensajes simplificados para monitoreo en tiempo real.
2.  **Archivo:** Registro detallado en la carpeta `/logs` con marca de tiempo.

Configuración rápida:
---------------------
1. Asegúrate de que este archivo (`logging_utils.py`) esté en el mismo directorio
   que tu script principal o en el PYTHONPATH.
2. Importa la función `setup_logging`.
3. Establece el nivel de logging en {level} (INFO por defecto): setup_logging({app_name},{target_dir},{level})
    Por defecto, este módulo tiene predefinidos 5 niveles para los logs, ordenados de menor a mayor criticidad:
    DEBUG, INFO, WARNING, ERROR, CRITICAL
    Una vez establecido un determinado nivel, solo se muestran logs de esa criticidad o superior.

Ejemplo de uso:
---------------
    from logging_utils import setup_logging
    import logging

    # 1. Inicializar al principio del script principal
    setup_logging(app_name="mi_proceso_pdf", target_dir="ruta/a/mi/carpeta", level="DEBUG")

    # 2. Usar los comandos estándar de logging
    logging.info("Iniciando el procesamiento de documentos...")
    logging.warning("El archivo X no tiene el formato esperado.")
    try:
        # Tu código aquí
        pass
    except Exception as e:
        logging.exception(f"Error crítico: {e}")
    # si necesitas todo el traceback en el log como hace exception pero fuera de uno:
    logging.error("Error crítico", exc_info=True)

Estructura de archivos generada:
--------------------------------
El módulo creará automáticamente una carpeta `logs/` y archivos con el formato:
`logs/nombre_app_YYYYMMDD_HHMMSS.log`
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path


def setup_logging(
    app_name: str = "app",
    target_dir: str | Path | None = None,
    level: int = logging.INFO,
):
    """
    Configura logging en consola + archivo logs/<app_name>_YYYYMMDD_HHMMSS.log

    app_name: nombre base del archivo de log (ej. "actualizar_bd", "preprocesador")
    target_dir: ruta base opcional donde se guardará la carpeta 'logs'. Si es None, se usa './logs'
    level: nivel de detalle del log
    """

    # Determinar el directorio de destino de forma dinámica y segura
    if target_dir is not None:
        log_dir = Path(target_dir) / "logs"
    else:
        log_dir = Path("logs")

    # Crear directorio logs si no existe
    os.makedirs(log_dir, exist_ok=True)

    # Archivo nuevo para cada ejecución
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # noqa: DTZ005
    log_path = log_dir / f"{app_name}_{timestamp}.log"

    # Root logger
    logger = logging.getLogger()
    # Usamos el nivel pasado por parámetro
    logger.setLevel(level)

    # Muy importante: limpiar handlers previos
    if logger.hasHandlers():
        logger.handlers.clear()

    # Formatos
    file_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    console_formatter = logging.Formatter("[%(levelname)s] %(message)s")

    # Handler archivo
    file_handler = logging.FileHandler(str(log_path), encoding="utf-8")
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    # Handler consola
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    logger.info(f"*** Logging configurado. Archivo: {log_path} ***")
