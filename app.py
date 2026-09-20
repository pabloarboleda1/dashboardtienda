"""
====================================================================
 TIENDA VIRTUAL UCN - PUNTO DE ENTRADA / ENRUTADOR
====================================================================
Este archivo NO contiene lógica de negocio: solo define el menú
lateral (nombre, orden e ícono de cada página) usando st.navigation.
Cada página real vive en la carpeta paginas/.

No edites este archivo para agregar contenido — agrega o edita
páginas en paginas/ y regístralas aquí abajo si son nuevas.
====================================================================
"""

import streamlit as st

st.set_page_config(
    page_title="Tienda Virtual UCN",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

paginas = [
    st.Page("paginas/reporte_ventas.py", title="Reporte de Ventas", icon="📊", default=True),
    st.Page("paginas/catalogo_productos.py", title="Catálogo de Productos", icon="🏷️"),
    st.Page("paginas/registrar_venta.py", title="Registro de Ventas", icon="📝"),
]

navegacion = st.navigation(paginas, position="sidebar")
navegacion.run()
