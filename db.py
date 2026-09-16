"""
====================================================================
 CONEXIÓN A LA BASE DE DATOS - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Busca la URL de conexión primero en los secrets de Streamlit
(DATABASE_URL) para uso permanente, y si no existe, la pide una vez
por sesión (igual que la Groq API Key).

Funciona con:
    PostgreSQL / Supabase:  postgresql+psycopg2://usuario:clave@host:puerto/basededatos
    MySQL:                  mysql+pymysql://usuario:clave@host:puerto/basededatos
====================================================================
"""

import streamlit as st
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from models import Base, TasaRetribucion, CATEGORIAS_RETRIBUCION


def obtener_database_url():
    try:
        url = st.secrets.get("DATABASE_URL", "")
    except Exception:
        url = ""
    if not url:
        url = st.session_state.get("database_url_sesion", "")
    return url


@st.cache_resource(show_spinner="Conectando a la base de datos...")
def obtener_engine(database_url):
    engine = create_engine(database_url, pool_pre_ping=True)
    Base.metadata.create_all(engine)
    _sembrar_tasas_iniciales(engine)
    return engine


def _sembrar_tasas_iniciales(engine):
    """Crea filas en blanco (tasa=0) para las categorías conocidas si la tabla está vacía,
    así el usuario solo tiene que editar el valor en vez de crear las filas desde cero."""
    Session = sessionmaker(bind=engine)
    sesion = Session()
    try:
        if sesion.query(TasaRetribucion).count() == 0:
            for categoria in CATEGORIAS_RETRIBUCION:
                if categoria == "OTRA":
                    continue
                sesion.add(TasaRetribucion(categoria=categoria, tasa_por_unidad=0.0))
            sesion.commit()
    finally:
        sesion.close()


def obtener_sesion():
    """Devuelve una sesión de SQLAlchemy lista para usar, o None si aún no hay conexión configurada."""
    database_url = obtener_database_url()
    if not database_url:
        return None
    engine = obtener_engine(database_url)
    Session = sessionmaker(bind=engine)
    return Session()


def probar_conexion(database_url):
    """Intenta conectar y crear las tablas; devuelve (True, None) o (False, mensaje_error)."""
    try:
        engine = create_engine(database_url, pool_pre_ping=True)
        Base.metadata.create_all(engine)
        _sembrar_tasas_iniciales(engine)
        with engine.connect():
            pass
        return True, None
    except Exception as e:
        return False, str(e)


def solicitar_conexion_si_falta():
    """Widget reutilizable: si no hay DATABASE_URL configurada, pide una y la valida.
    Devuelve True si ya hay una conexión lista para usar."""
    database_url = obtener_database_url()
    if database_url:
        return True

    st.info(
        "Todavía no hay una base de datos conectada. Pega aquí tu cadena de conexión de "
        "Supabase (Project Settings → Database → Connection string → modo 'Session pooler')."
    )
    url_ingresada = st.text_input(
        "Cadena de conexión (DATABASE_URL)", type="password",
        placeholder="postgresql+psycopg2://usuario:clave@host:puerto/basededatos",
        help="Solo se usa en esta sesión. Para que quede permanente, agrégala como Secret en Streamlit Cloud.",
    )
    if st.button("Conectar"):
        if not url_ingresada:
            st.warning("Pega primero la cadena de conexión.")
            return False
        with st.spinner("Probando la conexión..."):
            ok, error = probar_conexion(url_ingresada)
        if ok:
            st.session_state["database_url_sesion"] = url_ingresada
            st.success("¡Conectado! Cargando la página...")
            st.rerun()
        else:
            st.error(f"No se pudo conectar: {error}")
    return False
