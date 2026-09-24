"""
====================================================================
 EXPORTAR REPORTE - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Genera un Excel con el mismo formato de siempre (Consignaciones,
Inventario, Retribución FONDO) a partir de los datos reales de la
base de datos, filtrando por un mes o por un rango de fechas.
====================================================================
"""

import calendar
from datetime import date

import streamlit as st

from common import (
    inyectar_estilos, encabezado, render_kpi, formato_pesos, sidebar_brand, selector_tema,
    ACENTO_INGRESOS, ACENTO_UNIDADES,
)
from db import obtener_sesion, solicitar_conexion_si_falta
from models import Venta
from export_utils import generar_excel_reporte

sidebar_brand("Exportar Reporte")
selector_tema()
tema = inyectar_estilos()
encabezado("Exportar Reporte", "Genera el Excel de siempre (Consignaciones, Inventario, Retribución FONDO) con tus datos reales.")

if not solicitar_conexion_si_falta():
    st.stop()

sesion = obtener_sesion()

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

st.markdown('<div class="section-title">Elige el periodo</div>', unsafe_allow_html=True)
modo = st.radio("¿Cómo quieres filtrar?", ["Un mes", "Rango de fechas"], horizontal=True)

hoy = date.today()
if modo == "Un mes":
    c1, c2 = st.columns(2)
    anio = c1.number_input("Año", min_value=2020, max_value=2100, value=hoy.year, step=1)
    mes = c2.selectbox("Mes", list(MESES_ES.keys()), format_func=lambda m: MESES_ES[m], index=hoy.month - 1)
    fecha_inicio = date(int(anio), int(mes), 1)
    ultimo_dia = calendar.monthrange(int(anio), int(mes))[1]
    fecha_fin = date(int(anio), int(mes), ultimo_dia)
else:
    c1, c2 = st.columns(2)
    fecha_inicio = c1.date_input("Desde", value=hoy.replace(day=1))
    fecha_fin = c2.date_input("Hasta", value=hoy)

if fecha_inicio > fecha_fin:
    st.error("La fecha 'Desde' no puede ser posterior a 'Hasta'.")
    st.stop()

ventas_periodo = sesion.query(Venta).filter(Venta.fecha >= fecha_inicio, Venta.fecha <= fecha_fin).all()

st.write("")
col1, col2, col3 = st.columns(3)
render_kpi(col1, "Ventas en el periodo", str(len(ventas_periodo)), ACENTO_INGRESOS)
render_kpi(col2, "Ingresos del periodo", formato_pesos(sum(v.valor_venta or 0 for v in ventas_periodo)), ACENTO_INGRESOS)
render_kpi(col3, "Retribución al Fondo", formato_pesos(sum(v.retribucion_fondo or 0 for v in ventas_periodo)), ACENTO_UNIDADES)

st.write("")
if len(ventas_periodo) == 0:
    st.warning("No hay ventas registradas en ese periodo. El Excel se genera igual (con el catálogo completo), pero la hoja de Consignaciones saldrá vacía.")

if st.button("Generar Excel", use_container_width=True):
    with st.spinner("Generando el archivo..."):
        from models import Producto
        productos_activos = sesion.query(Producto).filter(Producto.activo == 1).all()
        excel_bytes = generar_excel_reporte(ventas_periodo, productos_activos)
    st.session_state["export_excel"] = excel_bytes
    st.session_state["export_excel_nombre"] = f"Reporte_Ventas_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.xlsx"

if "export_excel" in st.session_state:
    st.download_button(
        "Descargar Excel", data=st.session_state["export_excel"],
        file_name=st.session_state["export_excel_nombre"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

sesion.close()
