"""
====================================================================
 CATÁLOGO DE PRODUCTOS - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Alta de productos (con imagen opcional), edición de costo/precio/
stock, importación masiva desde un Excel de Inventario, y una vista
del catálogo organizada por Compra Origen → Categoría.
====================================================================
"""

import pandas as pd
import streamlit as st

from common import (
    inyectar_estilos, encabezado, render_kpi, formato_pesos, sidebar_brand, selector_tema,
    render_producto_card, imagen_a_base64, ACENTO_INGRESOS, ACENTO_UNIDADES, ACENTO_UTILIDAD_NEG,
)
from db import obtener_sesion, solicitar_conexion_si_falta
from models import Producto, CATEGORIAS_RETRIBUCION, ORIGENES_COMPRA
from excel_utils import leer_catalogo_desde_inventario

sidebar_brand("Catálogo de Productos")
selector_tema()
tema = inyectar_estilos()
encabezado("Catálogo de Productos", "Crea prendas nuevas, súbeles una foto, y ajusta costo, precio o stock.")

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

# --------------------------------------------------------------
# IMPORTAR DESDE UN EXCEL DE INVENTARIO
# --------------------------------------------------------------
st.markdown('<div class="section-title">Importar catálogo desde un Excel de Inventario</div>', unsafe_allow_html=True)
with st.expander("Subir un archivo mensual y traer sus productos automáticamente"):
    st.caption(
        "Lee la hoja 'Inventario', separa cada nombre en prenda + talla, y detecta si es del Fondo "
        "(termina en ' F' o el origen lo indica). Revisa la vista previa antes de confirmar."
    )
    archivo_import = st.file_uploader("Archivo .xlsx", type=["xlsx"], key="importar_inventario")

    if archivo_import is not None:
        filas, error = leer_catalogo_desde_inventario(archivo_import)
        if error:
            st.error(error)
        else:
            existentes = {(p.nombre, p.talla, p.compra_origen): p for p in sesion.query(Producto).all()}
            for f in filas:
                f["ya_existe"] = (f["nombre"], f["talla"], f["compra_origen"]) in existentes

            n_nuevos = sum(1 for f in filas if not f["ya_existe"])
            n_existentes = len(filas) - n_nuevos
            st.info(f"{len(filas)} productos leídos: {n_nuevos} nuevos, {n_existentes} ya están en tu catálogo.")

            tabla_preview = pd.DataFrame(filas)[[
                "nombre_original", "nombre", "talla", "categoria", "compra_origen",
                "costo_unitario", "precio_venta_unitario", "stock_actual", "ya_existe",
            ]].rename(columns={
                "nombre_original": "Nombre en Excel", "nombre": "Prenda (sin talla)", "talla": "Talla",
                "categoria": "Categoría", "compra_origen": "Origen", "costo_unitario": "Costo",
                "precio_venta_unitario": "Precio", "stock_actual": "Stock", "ya_existe": "Ya existe",
            })
            st.dataframe(
                tabla_preview.style.format({"Costo": formato_pesos, "Precio": formato_pesos}),
                use_container_width=True, hide_index=True,
            )

            actualizar_existentes = st.checkbox(
                "Actualizar costo, precio y stock de los productos que ya existen (en vez de dejarlos como están)",
                value=False,
            )
            if st.button("Confirmar importación", use_container_width=True):
                creados, actualizados = 0, 0
                for f in filas:
                    clave = (f["nombre"], f["talla"], f["compra_origen"])
                    if clave in existentes:
                        if actualizar_existentes:
                            p = existentes[clave]
                            p.costo_unitario = f["costo_unitario"]
                            p.precio_venta_unitario = f["precio_venta_unitario"]
                            p.stock_actual = f["stock_actual"]
                            p.categoria = f["categoria"]
                            p.compra_origen = f["compra_origen"]
                            actualizados += 1
                    else:
                        nuevo = Producto(
                            nombre=f["nombre"], talla=f["talla"], categoria=f["categoria"],
                            compra_origen=f["compra_origen"], costo_unitario=f["costo_unitario"],
                            precio_venta_unitario=f["precio_venta_unitario"], stock_actual=f["stock_actual"],
                        )
                        sesion.add(nuevo)
                        creados += 1
                sesion.commit()
                st.success(f"Importación lista: {creados} producto(s) nuevo(s), {actualizados} actualizado(s).")
                st.rerun()

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

    c7, c8 = st.columns([1, 2])
    stock_inicial = c7.number_input("Stock inicial (unidades)", min_value=0, step=1)
    imagen_subida = c8.file_uploader("Foto del producto (opcional)", type=["png", "jpg", "jpeg"])

    guardar = st.form_submit_button("Agregar producto", use_container_width=True)

    if guardar:
        if not nombre.strip() or not talla.strip():
            st.warning("El nombre y la talla son obligatorios.")
        else:
            existente = (
                sesion.query(Producto)
                .filter(Producto.nombre == nombre.strip(), Producto.talla == talla.strip(), Producto.compra_origen == origen)
                .first()
            )
            if existente:
                st.warning(f"Ya existe '{nombre.strip()} - {talla.strip()}' con origen '{origen}' en el catálogo. Edítalo abajo en vez de duplicarlo.")
            else:
                imagen_codificada = imagen_a_base64(imagen_subida)
                if imagen_codificada == "DEMASIADO_GRANDE":
                    st.warning("La imagen pesa más de 800 KB — se guardó el producto sin imagen. Usa una más liviana.")
                    imagen_codificada = None
                nuevo = Producto(
                    nombre=nombre.strip(), talla=talla.strip(), categoria=categoria, compra_origen=origen,
                    costo_unitario=costo_unitario, precio_venta_unitario=precio_venta, stock_actual=int(stock_inicial),
                    imagen_base64=imagen_codificada,
                )
                sesion.add(nuevo)
                sesion.commit()
                st.success(f"'{nuevo.nombre_completo}' agregado al catálogo.")
                st.rerun()

# --------------------------------------------------------------
# EXPLORAR CATÁLOGO: ORIGEN -> CATEGORÍA
# --------------------------------------------------------------
st.markdown('<div class="section-title">Explorar catálogo</div>', unsafe_allow_html=True)

if not productos:
    st.caption("Todavía no hay productos registrados — impórtalos o agrega el primero arriba.")
else:
    filtro = st.text_input("Buscar producto", placeholder="Escribe para filtrar por nombre...")
    productos_filtrados = [p for p in productos if filtro.lower() in p.nombre.lower()] if filtro else productos

    orden_origenes = ["FONDO DE EMPLEADOS", "UCN", "SOUVENIRS", "OTRO"]
    origenes_presentes = [o for o in orden_origenes if any((p.compra_origen or "OTRO") == o for p in productos_filtrados)]
    origenes_presentes += sorted(set((p.compra_origen or "OTRO") for p in productos_filtrados) - set(origenes_presentes))

    if not origenes_presentes:
        st.caption("Ningún producto coincide con la búsqueda.")
    else:
        tabs_origen = st.tabs(origenes_presentes)
        for tab, origen_tab in zip(tabs_origen, origenes_presentes):
            with tab:
                productos_origen = [p for p in productos_filtrados if (p.compra_origen or "OTRO") == origen_tab]
                categorias_presentes = sorted(set((p.categoria or "OTRA") for p in productos_origen))
                for categoria_tab in categorias_presentes:
                    productos_cat = sorted(
                        [p for p in productos_origen if (p.categoria or "OTRA") == categoria_tab],
                        key=lambda p: (p.nombre, p.talla),
                    )
                    with st.expander(f"{categoria_tab} ({len(productos_cat)})", expanded=True):
                        cols = st.columns(4)
                        for i, p in enumerate(productos_cat):
                            render_producto_card(cols[i % 4], p)

    st.write("")
    st.markdown("**Editar un producto**")
    opciones = {f"{p.nombre_completo} (ID {p.id})": p.id for p in productos_filtrados}
    seleccion = st.selectbox("Selecciona el producto a editar", ["—"] + list(opciones.keys()))

    if seleccion != "—":
        producto_id = opciones[seleccion]
        producto = sesion.query(Producto).get(producto_id)
        with st.form("form_editar_producto"):
            if producto.imagen_base64 and producto.imagen_base64 != "DEMASIADO_GRANDE":
                st.image(producto.imagen_base64, width=120)
            c1, c2, c3 = st.columns(3)
            nuevo_costo = c1.number_input("Costo unitario ($)", min_value=0.0, step=1000.0, format="%.0f", value=float(producto.costo_unitario or 0))
            nuevo_precio = c2.number_input("Precio de venta ($)", min_value=0.0, step=1000.0, format="%.0f", value=float(producto.precio_venta_unitario or 0))
            nuevo_stock = c3.number_input("Stock actual", min_value=0, step=1, value=int(producto.stock_actual or 0))
            nueva_categoria = st.selectbox(
                "Categoría", CATEGORIAS_RETRIBUCION,
                index=CATEGORIAS_RETRIBUCION.index(producto.categoria) if producto.categoria in CATEGORIAS_RETRIBUCION else 0,
            )
            nueva_imagen = st.file_uploader("Reemplazar foto (opcional)", type=["png", "jpg", "jpeg"], key="editar_imagen")

            col_g, col_d = st.columns(2)
            actualizar = col_g.form_submit_button("Guardar cambios", use_container_width=True)
            descontinuar = col_d.form_submit_button("Descontinuar producto", use_container_width=True)

            if actualizar:
                producto.costo_unitario = nuevo_costo
                producto.precio_venta_unitario = nuevo_precio
                producto.stock_actual = int(nuevo_stock)
                producto.categoria = nueva_categoria
                if nueva_imagen is not None:
                    codificada = imagen_a_base64(nueva_imagen)
                    if codificada == "DEMASIADO_GRANDE":
                        st.warning("La imagen pesa más de 800 KB — no se reemplazó.")
                    else:
                        producto.imagen_base64 = codificada
                sesion.commit()
                st.success("Producto actualizado.")
                st.rerun()
            if descontinuar:
                producto.activo = 0
                sesion.commit()
                st.success(f"'{producto.nombre_completo}' se marcó como descontinuado (sigue en el historial de ventas).")
                st.rerun()

sesion.close()
