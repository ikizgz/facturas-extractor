# facturas-extractor

Scripts para extraer datos de facturas y crear un Excel con ellos

## Parámetros de main.py

"--input", "-i", type=str, required=True, help="Carpeta con los PDFs")  
"--output", "-o", type=str, default=None, help="Excel de salida (por defecto: facturas_datos_extraidos.xlsx)"  
"--ocr", choices=["on", "off"], default="on", help="OCR si el PDF no tiene texto"  
"--dpi", type=int, default=DEFAULT_DPI, help=f"DPI para OCR (por defecto {DEFAULT_DPI})"  
"--poppler", type=str, default="", help="Ruta a binarios de Poppler si no están en PATH"  
"--tesseract", type=str, default="", help="Ruta completa a tesseract si no está en PATH"  
"--log", type=str, default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"]  
"--sleep-ms", type=int, default=DEFAULT_SLEEP_MS, help=f"Pausa (ms) entre páginas OCR (por defecto {DEFAULT_SLEEP_MS})"  
"--throttle-every", type=int, default=DEFAULT_THROTTLE_EVERY, help=f"Cada N PDFs aplicar pausa (por defecto {DEFAULT_THROTTLE_EVERY})"  
"--throttle-ms", type=int, default=DEFAULT_THROTTLE_MS, help=f"Pausa (ms) cuando se cumple throttle-every (por defecto {DEFAULT_THROTTLE_MS})"  
"--child-timeout-s", type=int, default=DEFAULT_CHILD_TIMEOUT_S, help=f"Timeout por PDF (por defecto {DEFAULT_CHILD_TIMEOUT_S})"  
