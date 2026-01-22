# app_config.py - VERSIÓN ANONIMIZADA PARA GITHUB

# --- CONFIGURACIÓN GLOBAL ---

# Nombres de columnas esperados en el Excel (Generic SCM Terms)
NOMBRES_BUSCADOS = {
    "inv": "Inventory", "dmd": "Total Demand", "wcd": "Confirmed Supply", 
    "wsd": "Planned Supply", "dos_act": "DoS Actual", 
    "tgt_inv": "Target Inv", "dos_tgt": "DoS Target"
}

# Tiempos de entrega (Lead Times) por Planta/Origen
LEAD_TIMES = {"PLANT_01": 8, "DEFAULT": 8}

# Parámetros de simulación
SEMANAS_A_EVALUAR = 20
MAX_SEMANAS_RECORRER = 4

# Colores Excel (Hex ARGB)
COLOR_VERDE = ("8ED973", "0D3512")
COLOR_ROJO = ("FFC7CE", "9C0006")
COLOR_AMARILLO = ("FFEB9C", "9C5700") 
COLOR_MORADO = ("E49EDD", "782170")
COLOR_NAVY = ("A3C2C7", "FFFFFF") 
COLOR_WK01_STATIC = ("65B675", "FFFFFF")

# Colores PDF (Hex Standard)
COLOR_PDF_ROJO = "E60000"      # OOS (Out of Stock)
COLOR_PDF_AMARILLO = "FFCC00"  # USTN (Under Stock Target)
COLOR_PDF_AZUL_CLARO = "99CCFF"# OSTN Med (Over Stock)
COLOR_PDF_AZUL_OSCURO = "003399"# OSTN High
COLOR_PDF_VERDE = "339933"     # OK

# Colores Condicionales Excel (Hex simple)
COLOR_R_MORADO = "E49EDD"
COLOR_R_VERDE = "B5E6A2"
COLOR_R_AMARILLO = "FFFF99"
COLOR_R_ROJO = "FF9999"

# Categorías de Producto (Mapeo Genérico)
CATEGORIAS_MAP = {
    "DH": "Electronics Category A", 
    "ME": "Main Kits", 
    "FA": "Consumables Type 1",
    "DC": "Chargers", 
    "DA": "Holders", 
    "DK": "Device Parts",
    "MA": "Consumables Type 2", 
    "NP": "New Products"
}

# Mapa de Mercados (Códigos ISO Genéricos)
ODM_MAP = {
    "MKT_01": "Colombia", "MKT_02": "Perú", "MKT_03": "Ecuador"
}

# Formatos Excel
FORMATO_CONTABILIDAD_DECIMAL = '_-* #,##0.0_-;-* #,##0.0_-;_-* "-"??_-;_-@_-'
FORMATO_CONTABILIDAD_ENTERO = '_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-'
FORMATO_FECHA = 'YYYY-MM-DD'