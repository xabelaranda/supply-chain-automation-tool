import pandas as pd

def limpiar_numero(valor):
    try:
        if pd.isna(valor): 
            return 0.0
        
        # Convertimos a string, quitamos comas y espacios
        s = str(valor).replace(",", "").strip()
        
        # Si está vacío o es un guion, es 0
        if not s or s == "-": 
            return 0.0
            
        return float(s)
    except:
        return 0.0

def normalizar(texto):
    return str(texto).strip().lower()