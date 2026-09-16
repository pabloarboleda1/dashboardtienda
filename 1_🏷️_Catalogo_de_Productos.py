"""
====================================================================
 CATÁLOGO DE PRODUCTOS - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Alta de productos nuevos (prenda + talla), edición de costo/precio/
stock, y vista general del catálogo con su stock actual.
====================================================================
"""

import streamlit as st
import pandas as pd

from common import inyectar_estilos, encabezado, render_kpi, formato_pesos, ACENTO_INGRESOS, ACENTO_UNIDADES, ACENTO_UTILIDAD_NEG
from db import obtener_sesion, solicitar_conexion_si_falta
from models import Producto, CATEGORIAS_RETRIBUCION, ORIGENES_COMPRA

st.set_page_config(page_title="Catálogo de Productos | Tienda Virtual UCN", page_icon="🏷️", layout="wide")
tema = inyectar_estilos()
encabezado("Catálogo de Productos", "Crea prendas y tallas nuevas, y ajusta costo, precio o stock existente.")

if not solicitar_conexion_si_falta():
    st.stop()

sesion = obtener_sesion()

# --------------------------------------------------------------
# KPIs
# --------------------------------------------------------------
productos = sesion.query(Producto).filter(Producto.activo == 1).all()
total_productos = len(productos)
valor_inventario = sum((p.costo_unitario or 0) * (p.stock_actual or 0) for p in productos)
stock_total = sum(p.stock_actual or 0 for p in productos)
sin_stock = sum(1 for p in productos if (p.stock_actual or 0) <= 0)

col1, col2, col3, col4 = st.columns(4)
render_kpi(col1, "Productos activos", str(total_productos), ACENTO_INGRESOS)
render_kpi(col2, "Unidades en stock", str(stock_total), ACENTO_UNIDADES)
render_kpi(col3, "Valor del inventario (costo)", formato_pesos(valor_inventario), ACENTO_INGRESOS)
render_kpi(col4, "Productos agotados", str(sin_stock), ACENTO_UTILIDAD_NEG if sin_stock else ACENTO_UNIDADES)

st.write("")

# --------------------------------------------------------------
# AGREGAR PRODUCTO NUEVO
# --------------------------------------------------------------
st.markdown('<div class="section-title">Agregar producto nuevo</div>', unsafe_allow_html=True)
with st.form("form_nuevo_producto", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    nombre = c1.text_input("Nombre de la prenda (sin talla)", placeholder="Ej: Camiseta Azul Mujer")
    talla = c2.text_input("Talla", placeholder="Ej: S, M, L, XL, Única")
    categoria = c3.selectbox("Categoría (para retribución al Fondo)", CATEGORIAS_RETRIBUCION)

    c4, c5, c6 = st.columns(3)
    origen = c4.selectbox("Compra Origen", ORIGENES_COMPRA)
    costo_unitario = c5.number_input("Costo unitario ($)", min_value=0.0, step=1000.0, format="%.0f")
    precio_venta = c6.number_input("Precio de venta unitario ($)", min_value=0.0, step=1000.0, format="%.0f")

    stock_inicial = st.number_input("Stock inicial (unidades)", min_value=0, step=1)
    guardar = st.form_submit_button("Agregar producto", use_container_width=True)

    if guardar:
        if not nombre.strip() or not talla.strip():
            st.warning("El nombre y la talla son obligatorios.")
        else:
            existente = (
                sesion.query(Producto)
                .filter(Producto.nombre == nombre.strip(), Producto.talla == talla.strip())
                .first()
            )
            if existente:
                st.warning(f"Ya existe '{nombre.strip()} - {talla.strip()}' en el catálogo. Edítalo abajo en vez de duplicarlo.")
            else:
                nuevo = Producto(
                    nombre=nombre.strip(), talla=talla.strip(), categoria=categoria, compra_origen=origen,
                    costo_unitario=costo_unitario, precio_venta_unitario=precio_venta, stock_actual=int(stock_inicial),
                )
                sesion.add(nuevo)
                sesion.commit()
                st.success(f"'{nuevo.nombre_completo}' agregado al catálogo.")
                st.rerun()

st.write("")

# --------------------------------------------------------------
# CATÁLOGO ACTUAL + EDICIÓN
# --------------------------------------------------------------
st.markdown('<div class="section-title">Catálogo actual</div>', unsafe_allow_html=True)

if not productos:
    st.caption("Todavía no hay productos registrados — agrega el primero arriba.")
else:
    filtro = st.text_input("Buscar producto", placeholder="Escribe para filtrar por nombre...")
    productos_filtrados = [p for p in productos if filtro.lower() in p.nombre.lower()] if filtro else productos

    tabla = pd.DataFrame([{
        "ID": p.id, "Producto": p.nombre, "Talla": p.talla, "Categoría": p.categoria,
        "Origen": p.compra_origen, "Costo Unitario": p.costo_unitario, "Precio Venta": p.precio_venta_unitario,
        "Stock Actual": p.stock_actual,
    } for p in productos_filtrados])
    st.dataframe(
        tabla.style.format({"Costo Unitario": formato_pesos, "Precio Venta": formato_pesos}),
        use_container_width=True, hide_index=True,
    )

    st.markdown("**Editar un producto**")
    opciones = {f"{p.nombre_completo} (ID {p.id})": p.id for p in productos_filtrados}
    seleccion = st.selectbox("Selecciona el producto a editar", ["—"] + list(opciones.keys()))

    if seleccion != "—":
        producto_id = opciones[seleccion]
        producto = sesion.query(Producto).get(producto_id)
        with st.form("form_editar_producto"):
            c1, c2, c3 = st.columns(3)
            nuevo_costo = c1.number_input("Costo unitario ($)", min_value=0.0, step=1000.0, format="%.0f", value=float(producto.costo_unitario or 0))
            nuevo_precio = c2.number_input("Precio de venta ($)", min_value=0.0, step=1000.0, format="%.0f", value=float(producto.precio_venta_unitario or 0))
            nuevo_stock = c3.number_input("Stock actual", min_value=0, step=1, value=int(producto.stock_actual or 0))
            nueva_categoria = st.selectbox("Categoría", CATEGORIAS_RETRIBUCION, index=CATEGORIAS_RETRIBUCION.index(producto.categoria) if producto.categoria in CATEGORIAS_RETRIBUCION else 0)

            col_g, col_d = st.columns(2)
            actualizar = col_g.form_submit_button("Guardar cambios", use_container_width=True)
            descontinuar = col_d.form_submit_button("Descontinuar producto", use_container_width=True)

            if actualizar:
                producto.costo_unitario = nuevo_costo
                producto.precio_venta_unitario = nuevo_precio
                producto.stock_actual = int(nuevo_stock)
                producto.categoria = nueva_categoria
                sesion.commit()
                st.success("Producto actualizado.")
                st.rerun()
            if descontinuar:
                producto.activo = 0
                sesion.commit()
                st.success(f"'{producto.nombre_completo}' se marcó como descontinuado (sigue en el historial de ventas).")
                st.rerun()

sesion.close()
