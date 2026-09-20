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
from sqlalchemy import create_engine, inspect, text, UniqueConstraint
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
    _migrar_columnas_faltantes(engine)
    _migrar_restricciones_unicas(engine)
    _sembrar_tasas_iniciales(engine)
    return engine


def _migrar_columnas_faltantes(engine):
    """Si el modelo (models.py) tiene columnas que la tabla real todavía no tiene
    (porque la tabla ya existía de antes de agregarlas), las agrega con ALTER TABLE.
    Así no hay que borrar/recrear la base de datos cada vez que evoluciona el esquema."""
    inspector = inspect(engine)
    for tabla in Base.metadata.tables.values():
        if not inspector.has_table(tabla.name):
            continue  # tabla nueva: create_all ya la crea completa
        columnas_existentes = {c["name"] for c in inspector.get_columns(tabla.name)}
        for columna in tabla.columns:
            if columna.name in columnas_existentes:
                continue
            tipo_sql = columna.type.compile(dialect=engine.dialect)
            try:
                with engine.begin() as conn:
                    conn.execute(text(f'ALTER TABLE "{tabla.name}" ADD COLUMN "{columna.name}" {tipo_sql}'))
            except Exception as e:
                st.warning(f"No se pudo agregar la columna '{columna.name}' a '{tabla.name}': {e}")


def _migrar_restricciones_unicas(engine):
    """Si cambió qué columnas forman la restricción 'única' de una tabla (ej. productos
    pasó de (nombre, talla) a (nombre, talla, compra_origen)), tumba la restricción vieja
    y crea la nueva. Best-effort: si algo falla, solo avisa, no rompe la app."""
    inspector = inspect(engine)
    for tabla in Base.metadata.tables.values():
        if not inspector.has_table(tabla.name):
            continue
        esperadas = [c for c in tabla.constraints if isinstance(c, UniqueConstraint)]
        if not esperadas:
            continue
        try:
            actuales = inspector.get_unique_constraints(tabla.name)
        except Exception:
            continue
        for restriccion in esperadas:
            columnas_esperadas = set(col.name for col in restriccion.columns)
            if any(set(r["column_names"]) == columnas_esperadas for r in actuales):
                continue  # ya está tal cual la necesitamos
            for r in actuales:
                if set(r["column_names"]).issubset(columnas_esperadas) and r["column_names"]:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text(f'ALTER TABLE "{tabla.name}" DROP CONSTRAINT "{r["name"]}"'))
                    except Exception as e:
                        st.warning(f"No se pudo quitar la restricción vieja '{r['name']}': {e}")
            nombre_restriccion = restriccion.name or f"uq_{tabla.name}_{'_'.join(sorted(columnas_esperadas))}"
            cols_sql = ", ".join(f'"{c}"' for c in columnas_esperadas)
            try:
                with engine.begin() as conn:
                    conn.execute(text(f'ALTER TABLE "{tabla.name}" ADD CONSTRAINT "{nombre_restriccion}" UNIQUE ({cols_sql})'))
            except Exception as e:
                st.warning(f"No se pudo crear la restricción única en '{tabla.name}': {e}")


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
    """Intenta conectar y crear/migrar las tablas; devuelve (True, None) o (False, mensaje_error)."""
    try:
        engine = create_engine(database_url, pool_pre_ping=True)
        Base.metadata.create_all(engine)
        _migrar_columnas_faltantes(engine)
        _migrar_restricciones_unicas(engine)
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
