"""
====================================================================
 REGISTRO DE VENTAS - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Formulario rápido: elige la prenda y la talla de un catálogo (nada
de texto libre, así no hay typos), cantidad, canal, fecha, nombre y
documento del cliente. Calcula solo el costo, la utilidad y la
retribución al Fondo, y descuenta el stock automáticamente.
====================================================================
"""

from datetime import date

import pandas as pd
import streamlit as st

from common import inyectar_estilos, encabezado, render_kpi, formato_pesos, sidebar_brand, selector_tema, ACENTO_INGRESOS, ACENTO_UNIDADES
from db import obtener_sesion, solicitar_conexion_si_falta
from models import Producto, Venta, TasaRetribucion, CANALES_VENTA

sidebar_brand("Registro de Ventas")
selector_tema()
tema = inyectar_estilos()
encabezado("Registro de Ventas", "Elige la prenda y la talla — el sistema calcula el resto solo.")

if not solicitar_conexion_si_falta():
    st.stop()

sesion = obtener_sesion()

productos_activos = (
    sesion.query(Producto)
    .filter(Producto.activo == 1, Producto.stock_actual > 0)
    .order_by(Producto.nombre, Producto.talla)
    .all()
)

if not productos_activos:
    st.warning("No hay productos con stock disponible. Ve a 'Catálogo de Productos' para agregar o importar productos.")
    sesion.close()
    st.stop()

nombres_unicos = sorted(set(p.nombre for p in productos_activos))

st.markdown('<div class="section-title">Nueva venta</div>', unsafe_allow_html=True)

col_prenda, col_talla = st.columns(2)
prenda_elegida = col_prenda.selectbox("Prenda", nombres_unicos)
tallas_disponibles = [p for p in productos_activos if p.nombre == prenda_elegida]
opciones_talla = {f"{p.talla} — {p.compra_origen or 'Sin origen'} (stock: {p.stock_actual})": p for p in tallas_disponibles}
talla_elegida_label = col_talla.selectbox("Talla", list(opciones_talla.keys()))
producto = opciones_talla[talla_elegida_label]

col_info, col_img = st.columns([4, 1])
col_info.caption(f"Precio de venta sugerido: {formato_pesos(producto.precio_venta_unitario)} · Stock disponible: {producto.stock_actual}")
if producto.imagen_base64 and producto.imagen_base64 != "DEMASIADO_GRANDE":
    col_img.image(producto.imagen_base64, width=64)

with st.form("form_registrar_venta", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    cantidad = c1.number_input("Cantidad", min_value=1, max_value=int(producto.stock_actual), step=1, value=1)
    canal = c2.selectbox("Canal de venta", CANALES_VENTA)
    fecha_venta = c3.date_input("Fecha", value=date.today())

    valor_venta = st.number_input(
        "Valor de la venta ($)", min_value=0.0, step=1000.0, format="%.0f",
        value=float(producto.precio_venta_unitario * cantidad),
        help="Se sugiere Precio de venta × Cantidad; ajústalo si hubo descuento.",
    )

    st.markdown("**Datos del cliente** (opcional, para conciliar consignaciones)")
    c4, c5 = st.columns(2)
    cliente_nombre = c4.text_input("Nombre del cliente")
    cliente_documento = c5.text_input("Número de documento")

    registrar = st.form_submit_button("Registrar Venta", use_container_width=True)

    if registrar:
        if cantidad > producto.stock_actual:
            st.error(f"No hay suficiente stock: quedan {producto.stock_actual} unidades.")
        else:
            costo_total = (producto.costo_unitario or 0) * cantidad
            utilidad = valor_venta - costo_total

            tasa_row = sesion.query(TasaRetribucion).filter(TasaRetribucion.categoria == producto.categoria).first()
            es_fondo = (producto.compra_origen or "").strip().upper() == "FONDO DE EMPLEADOS" or producto.nombre.strip().upper().endswith(" F")
            retribucion = (tasa_row.tasa_por_unidad * cantidad) if (es_fondo and tasa_row and tasa_row.tasa_por_unidad) else 0.0

            nueva_venta = Venta(
                producto_id=producto.id, fecha=fecha_venta, cantidad=cantidad, canal=canal,
                cliente_nombre=cliente_nombre.strip() or None, cliente_documento=cliente_documento.strip() or None,
                valor_venta=valor_venta, costo_total=costo_total, utilidad=utilidad, retribucion_fondo=retribucion,
            )
            producto.stock_actual -= cantidad
            sesion.add(nueva_venta)
            sesion.commit()

            st.success(
                f"Venta registrada: {cantidad} × {producto.nombre_completo} — "
                f"{formato_pesos(valor_venta)} (utilidad: {formato_pesos(utilidad)})"
                + (f" · Retribución al Fondo: {formato_pesos(retribucion)}" if retribucion else "")
            )
            st.rerun()

st.write("")

# --------------------------------------------------------------
# ÚLTIMAS VENTAS REGISTRADAS
# --------------------------------------------------------------
st.markdown('<div class="section-title">Últimas ventas registradas</div>', unsafe_allow_html=True)
ultimas = sesion.query(Venta).order_by(Venta.registrado_en.desc()).limit(15).all()
if not ultimas:
    st.caption("Aún no hay ventas registradas.")
else:
    tabla = pd.DataFrame([{
        "Fecha": v.fecha, "Producto": v.producto.nombre_completo if v.producto else "(eliminado)",
        "Cantidad": v.cantidad, "Canal": v.canal, "Cliente": v.cliente_nombre or "—",
        "Valor": v.valor_venta, "Utilidad": v.utilidad, "Retribución": v.retribucion_fondo,
    } for v in ultimas])
    st.dataframe(
        tabla.style.format({"Valor": formato_pesos, "Utilidad": formato_pesos, "Retribución": formato_pesos}),
        use_container_width=True, hide_index=True,
    )

sesion.close()
