"""
Utilidades para leer la hoja "Inventario" de un Excel mensual y separar
cada producto en (nombre base, talla, es del Fondo) para poblar el
catálogo de la base de datos. Independiente de paginas/reporte_ventas.py
para no arriesgar esa lógica ya probada.
"""

import re
import unicodedata

import pandas as pd

PATRON_TALLA = re.compile(r"\s+(XXXL|XXL|XL|XS|S|M|L)$", flags=re.IGNORECASE)

ALIAS_ENCABEZADOS_INVENTARIO = {
    "cantidadvendidafsica": "cantidadvendidafisica",
    "cantdeduccinnmina": "cantdeduccionnomina",
}


def normalizar_texto(texto):
    texto = str(texto).strip().lower()
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9]", "", texto)
    return texto


def construir_lookup(columnas, alias_dict):
    lookup = {}
    for col in columnas:
        norm = normalizar_texto(col)
        norm = alias_dict.get(norm, norm)
        lookup[norm] = col
    return lookup


def obtener_columna(lookup, *candidatos):
    for c in candidatos:
        if c in lookup:
            return lookup[c]
    return None


def categoria_retribucion(nombre_producto):
    n = str(nombre_producto).upper()
    if "CAMISETA" in n:
        return "CAMISETAS"
    if "CHAQUETAC" in n or "CHAQUETA" in n:
        return "CHAQUETAS CORTAVIENTOS"
    if "CHOMPA" in n:
        return "CHAQUETAS ABULLONADAS"
    if "GORRA" in n:
        return "GORRAS"
    return "OTRA"


def parsear_producto_inventario(nombre_crudo):
    """'Camiseta Blanca Mujer S F' -> ('Camiseta Blanca Mujer', 'S', True)"""
    texto = re.sub(r"\s+", " ", str(nombre_crudo).strip())
    es_fondo_por_nombre = bool(re.search(r"\s+F$", texto, flags=re.IGNORECASE))
    if es_fondo_por_nombre:
        texto = re.sub(r"\s+F$", "", texto, flags=re.IGNORECASE).strip()
    match = PATRON_TALLA.search(texto)
    if match:
        talla = match.group(1).upper()
        nombre_base = texto[: match.start()].strip()
    else:
        talla = "Única"
        nombre_base = texto
    return nombre_base, talla, es_fondo_por_nombre


def limpiar_valor_numerico(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = re.sub(r"[^\d,.\-]", "", str(valor))
    if texto == "":
        return 0.0
    tiene_coma, tiene_punto = "," in texto, "." in texto
    if tiene_coma and tiene_punto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif tiene_coma:
        partes = texto.split(",")
        texto = texto.replace(",", "") if len(partes) > 1 and len(partes[-1]) == 3 else texto.replace(",", ".")
    elif tiene_punto:
        partes = texto.split(".")
        if len(partes) > 1 and len(partes[-1]) == 3:
            texto = texto.replace(".", "")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def leer_catalogo_desde_inventario(archivo):
    """Lee la hoja 'Inventario' de un Excel y devuelve una lista de dicts listos para
    previsualizar/importar: nombre, talla, categoria, compra_origen, costo_unitario,
    precio_venta_unitario, stock_actual (= Inventario Final)."""
    xls = pd.ExcelFile(archivo)
    hoja = next((n for n in xls.sheet_names if "inventario" in n.strip().lower()), None)
    if hoja is None:
        return None, "No se encontró una hoja llamada 'Inventario' en este archivo."

    df = pd.read_excel(xls, sheet_name=hoja, header=0)
    lookup = construir_lookup(df.columns, ALIAS_ENCABEZADOS_INVENTARIO)
    col_producto = obtener_columna(lookup, "producto")
    col_origen = obtener_columna(lookup, "compraorigen")
    col_costo = obtener_columna(lookup, "costounitario")
    col_precio = obtener_columna(lookup, "precioventaunitario")
    col_inv_final = obtener_columna(lookup, "inventariofinal")

    faltantes = [n for n, c in [("Producto", col_producto), ("Compra origen", col_origen)] if c is None]
    if faltantes:
        return None, f"Faltan columnas esperadas en 'Inventario': {', '.join(faltantes)}."

    filas = []
    for _, r in df.iterrows():
        nombre_crudo = r[col_producto]
        if pd.isna(nombre_crudo) or not str(nombre_crudo).strip():
            continue
        nombre_base, talla, es_fondo_por_nombre = parsear_producto_inventario(nombre_crudo)
        origen_excel = str(r[col_origen]).strip() if col_origen and pd.notna(r[col_origen]) else ""

        if es_fondo_por_nombre or "fondo" in origen_excel.lower():
            origen_final = "FONDO DE EMPLEADOS"
        elif origen_excel:
            origen_final = origen_excel.upper()
        else:
            origen_final = "OTRO"

        filas.append({
            "nombre_original": str(nombre_crudo).strip(),
            "nombre": nombre_base,
            "talla": talla,
            "categoria": categoria_retribucion(nombre_base),
            "compra_origen": origen_final,
            "costo_unitario": limpiar_valor_numerico(r[col_costo]) if col_costo else 0.0,
            "precio_venta_unitario": limpiar_valor_numerico(r[col_precio]) if col_precio else 0.0,
            "stock_actual": int(limpiar_valor_numerico(r[col_inv_final])) if col_inv_final else 0,
        })

    return filas, None
