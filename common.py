"""Estilos y utilidades compartidas por las páginas del Gestor de Inventario."""

import base64

import streamlit as st

TEMAS = {
    "Claro": {
        "bg_app": "#F7F9FC", "bg_sidebar": "#FFFFFF", "bg_card": "#FFFFFF",
        "texto": "#0F172A", "texto_muted": "#64748B", "borde": "#E7EBF3",
    },
    "Oscuro": {
        "bg_app": "#0B1220", "bg_sidebar": "#111A2C", "bg_card": "#161F33",
        "texto": "#F1F5F9", "texto_muted": "#93A2BD", "borde": "#243252",
    },
}
ACENTO_INGRESOS = "#2F6FED"
ACENTO_UTILIDAD_POS = "#15B36F"
ACENTO_UTILIDAD_NEG = "#E5484D"
ACENTO_UNIDADES = "#8B5CF6"
ACENTO_MARGEN = "#F5A524"


def obtener_tema():
    if "tema" not in st.session_state:
        st.session_state["tema"] = "Claro"
    return TEMAS[st.session_state["tema"]]


def selector_tema(ubicacion=None):
    ubicacion = ubicacion or st.sidebar
    ubicacion.markdown('<div class="sidebar-section">Apariencia</div>', unsafe_allow_html=True)
    st.session_state["tema"] = ubicacion.radio(
        "Tema", options=["Claro", "Oscuro"], horizontal=True,
        index=["Claro", "Oscuro"].index(st.session_state.get("tema", "Claro")), label_visibility="collapsed",
    )


def inyectar_estilos():
    tema = obtener_tema()
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        html, body, [class*="css"] {{ font-family: 'Inter', -apple-system, sans-serif; }}

        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        .stDeployButton {{display: none;}}
        header[data-testid="stHeader"] {{background: transparent;}}

        [data-testid="stAppViewContainer"] {{ background-color: {tema['bg_app']}; }}
        [data-testid="stSidebar"] {{ background-color: {tema['bg_sidebar']}; border-right: 1px solid {tema['borde']}; }}
        .block-container {{ padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1100px; }}
        h1, h2, h3, p, span, label, .stMarkdown {{ color: {tema['texto']}; }}

        [data-testid="stSidebarNav"] {{ padding-top: 0.5rem; }}
        [data-testid="stSidebarNav"] li div a span {{ font-weight: 600; font-size: 0.92rem; }}

        .app-header {{ padding-bottom: 1rem; margin-bottom: 1.5rem; border-bottom: 2px solid {tema['borde']}; }}
        .app-header .eyebrow {{ font-size: 0.76rem; font-weight: 700; letter-spacing: 0.14em; text-transform: uppercase; color: {ACENTO_INGRESOS}; }}
        .app-header .title {{ font-size: 1.9rem; font-weight: 800; color: {tema['texto']}; line-height: 1.2; }}
        .app-header .subtitle {{ font-size: 0.92rem; color: {tema['texto_muted']}; margin-top: 0.2rem; }}

        .sidebar-brand {{ padding: 0.1rem 0 1rem 0; margin-bottom: 0.8rem; border-bottom: 1px solid {tema['borde']}; }}
        .sidebar-brand .name {{ font-size: 1.02rem; font-weight: 800; color: {tema['texto']}; }}
        .sidebar-brand .tag {{ font-size: 0.76rem; color: {tema['texto_muted']}; }}
        .sidebar-section {{ font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: {tema['texto_muted']}; margin: 1rem 0 0.4rem 0; }}

        .kpi-card {{ background: {tema['bg_card']}; border-radius: 14px; border: 1px solid {tema['borde']}; padding: 1.1rem 1.3rem; box-shadow: 0 1px 3px rgba(15,23,42,0.06); }}
        .kpi-card .kpi-label {{ font-size: 0.75rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: {tema['texto_muted']}; margin-bottom: 0.4rem; }}
        .kpi-card .kpi-value {{ font-size: 1.6rem; font-weight: 800; }}

        .section-title {{ font-size: 1.08rem; font-weight: 700; color: {tema['texto']}; margin: 1.6rem 0 0.9rem 0; }}
        .section-title:first-of-type {{ margin-top: 0.2rem; }}
        [data-testid="stDataFrame"] {{ border: 1px solid {tema['borde']}; border-radius: 10px; }}

        /* Tarjeta de producto (catálogo) */
        .prod-card {{
            background: {tema['bg_card']}; border: 1px solid {tema['borde']}; border-radius: 14px;
            padding: 0.9rem; height: 100%; display: flex; flex-direction: column; gap: 0.35rem;
            box-shadow: 0 1px 3px rgba(15,23,42,0.05);
        }}
        .prod-card .prod-img {{
            width: 100%; aspect-ratio: 1/1; border-radius: 10px; object-fit: cover;
            background: {tema['bg_app']}; border: 1px solid {tema['borde']};
        }}
        .prod-card .prod-img-placeholder {{
            width: 100%; aspect-ratio: 1/1; border-radius: 10px; display: flex; align-items: center;
            justify-content: center; font-size: 2rem; background: {tema['bg_app']}; border: 1px dashed {tema['borde']};
            color: {tema['texto_muted']};
        }}
        .prod-card .prod-nombre {{ font-weight: 700; font-size: 0.92rem; color: {tema['texto']}; margin-top: 0.3rem; }}
        .prod-card .prod-precio {{ font-weight: 800; font-size: 0.98rem; color: {ACENTO_INGRESOS}; }}
        .stock-badge {{ display: inline-block; padding: 0.1rem 0.55rem; border-radius: 999px; font-size: 0.72rem; font-weight: 700; width: fit-content; }}
        .stock-ok {{ background: rgba(21,179,111,0.14); color: {ACENTO_UTILIDAD_POS}; }}
        .stock-bajo {{ background: rgba(245,165,36,0.16); color: {ACENTO_MARGEN}; }}
        .stock-agotado {{ background: rgba(229,72,77,0.14); color: {ACENTO_UTILIDAD_NEG}; }}

        .qa-ok {{ background: rgba(21,179,111,0.12); border: 1px solid {ACENTO_UTILIDAD_POS}; color: {tema['texto']}; padding: 0.65rem 1rem; border-radius: 10px; font-size: 0.86rem; }}
        .qa-alert {{ background: rgba(229,72,77,0.12); border: 1px solid {ACENTO_UTILIDAD_NEG}; color: {tema['texto']}; padding: 0.65rem 1rem; border-radius: 10px; font-size: 0.86rem; }}

        div[data-testid="stForm"] {{ border: 1px solid {tema['borde']}; border-radius: 14px; padding: 1.2rem; background: {tema['bg_card']}; }}
        .stButton button, .stFormSubmitButton button {{ border-radius: 8px; font-weight: 600; }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    return tema


def sidebar_brand(tag="Gestor de Inventario"):
    st.sidebar.markdown(
        f"""<div class="sidebar-brand"><div class="name">TIENDA VIRTUAL UCN</div>
        <div class="tag">{tag}</div></div>""",
        unsafe_allow_html=True,
    )


def encabezado(titulo, subtitulo, eyebrow="Tienda Virtual UCN"):
    st.markdown(
        f"""<div class="app-header"><div class="eyebrow">{eyebrow}</div>
        <div class="title">{titulo}</div><div class="subtitle">{subtitulo}</div></div>""",
        unsafe_allow_html=True,
    )


def render_kpi(col, etiqueta, valor, color_valor):
    col.markdown(
        f"""<div class="kpi-card"><div class="kpi-label">{etiqueta}</div>
        <div class="kpi-value" style="color:{color_valor};">{valor}</div></div>""",
        unsafe_allow_html=True,
    )


def formato_pesos(valor):
    signo = "-" if valor < 0 else ""
    return f"{signo}$ {abs(valor):,.0f}".replace(",", ".")


def imagen_a_base64(archivo_subido, max_bytes=800_000):
    """Convierte un archivo subido (imagen) a base64 listo para guardar. None si es muy pesado."""
    if archivo_subido is None:
        return None
    datos = archivo_subido.getvalue()
    if len(datos) > max_bytes:
        return "DEMASIADO_GRANDE"
    mime = archivo_subido.type or "image/png"
    return f"data:{mime};base64,{base64.b64encode(datos).decode()}"


def badge_stock(stock, umbral_bajo=3):
    if stock <= 0:
        return '<span class="stock-badge stock-agotado">Agotado</span>'
    if stock <= umbral_bajo:
        return f'<span class="stock-badge stock-bajo">{stock} u. — bajo</span>'
    return f'<span class="stock-badge stock-ok">{stock} u.</span>'


def render_producto_card(col, producto):
    """Tarjeta minimalista de un producto: imagen, nombre+talla, precio, stock."""
    with col:
        img_html = (
            f'<img class="prod-img" src="{producto.imagen_base64}" />'
            if producto.imagen_base64 and producto.imagen_base64 != "DEMASIADO_GRANDE"
            else '<div class="prod-img-placeholder">🏷️</div>'
        )
        st.markdown(
            f"""<div class="prod-card">{img_html}
            <div class="prod-nombre">{producto.nombre_completo}</div>
            <div class="prod-precio">{formato_pesos(producto.precio_venta_unitario)}</div>
            {badge_stock(producto.stock_actual)}
            </div>""",
            unsafe_allow_html=True,
        )
