"""Estilos y utilidades compartidas por las páginas del Gestor de Inventario."""

import streamlit as st

TEMAS = {
    "Claro": {
        "bg_app": "#EEF2F8", "bg_sidebar": "#FFFFFF", "bg_card": "#FFFFFF",
        "texto": "#0F172A", "texto_muted": "#5B6B82", "borde": "#E1E7F0",
    },
    "Oscuro": {
        "bg_app": "#0B1220", "bg_sidebar": "#111A2C", "bg_card": "#161F33",
        "texto": "#F1F5F9", "texto_muted": "#9AA8C0", "borde": "#26324A",
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
        .block-container {{ padding-top: 1.4rem; max-width: 1100px; }}
        h1, h2, h3, p, span, label, .stMarkdown {{ color: {tema['texto']}; }}

        .app-header {{ padding-bottom: 1rem; margin-bottom: 1.4rem; border-bottom: 2px solid {tema['borde']}; }}
        .app-header .eyebrow {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: {ACENTO_INGRESOS}; }}
        .app-header .title {{ font-size: 1.85rem; font-weight: 800; color: {tema['texto']}; line-height: 1.2; }}
        .app-header .subtitle {{ font-size: 0.9rem; color: {tema['texto_muted']}; margin-top: 0.15rem; }}

        .kpi-card {{ background: {tema['bg_card']}; border-radius: 12px; border: 1px solid {tema['borde']}; padding: 1.05rem 1.25rem; box-shadow: 0 2px 6px rgba(0,0,0,0.08); }}
        .kpi-card .kpi-label {{ font-size: 0.76rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: {tema['texto_muted']}; margin-bottom: 0.4rem; }}
        .kpi-card .kpi-value {{ font-size: 1.55rem; font-weight: 800; }}

        .section-title {{ font-size: 1.05rem; font-weight: 700; color: {tema['texto']}; margin: 0.2rem 0 0.8rem 0; }}
        [data-testid="stDataFrame"] {{ border: 1px solid {tema['borde']}; border-radius: 8px; }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    return tema


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
