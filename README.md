gestor_facturas/
├── .gitignore
├── README.md
├── requirements.txt
├── pyproject.toml             # Configuración moderna del proyecto y dependencias
├── src/                       # Carpeta raíz del código fuente (separa el código de los tests)
│   └── gestor_facturas/
│       ├── __init__.py
│       ├── main.py            # Punto de entrada de la aplicación (CLI o ejecución principal)
│       ├── core/              # Lógica central del orquestrador
│       │   ├── __init__.py
│       │   └── pipeline.py
│       ├── services/          # Servicios externos y procesamiento
│       │   ├── __init__.py
│       │   ├── extraccion.py  # Aquí irá pdfplumber
│       │   ├── ocr.py         # Tratamiento de imágenes / Tesseract
│       │   └── excel.py       # Escritura con openpyxl/pandas
│       └── providers/         # Parsers específicos por proveedor
│           ├── __init__.py
│           ├── base.py
│           └── amazon.py
└── tests/                     # Pruebas automáticas separadas del código fuente
    ├── __init__.py
    ├── conftest.py            # Configuración compartida de pytest (fixtures, datos de prueba)
    ├── unit/                  # Pruebas unitarias por módulo
    │   └── test_providers.py
    └── integration/           # Pruebas de integración con PDFs reales de muestra
        └── test_extraction.py