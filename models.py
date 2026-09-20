"""
====================================================================
 MODELOS DE BASE DE DATOS - GESTOR DE INVENTARIO TIENDA VIRTUAL UCN
====================================================================
Definidos con SQLAlchemy, funcionan igual con PostgreSQL (Supabase)
o MySQL — solo cambia la URL de conexión en db.py, este archivo no
se toca según cuál elijas.
====================================================================
"""

from datetime import datetime, date

from sqlalchemy import Column, Integer, String, Float, Date, DateTime, ForeignKey, UniqueConstraint, Text
from sqlalchemy.orm import relationship, declarative_base

Base = declarative_base()

# Categorías reconocidas para el cálculo de Retribución al Fondo de Empleados
CATEGORIAS_RETRIBUCION = [
    "CAMISETAS", "CHAQUETAS CORTAVIENTOS", "CHAQUETAS ABULLONADAS", "GORRAS", "OTRA",
]
CANALES_VENTA = ["Física (Consignación)", "Online (En línea)", "Deducción Nómina"]
ORIGENES_COMPRA = ["FONDO DE EMPLEADOS", "UCN", "SOUVENIRS", "OTRO"]


class Producto(Base):
    __tablename__ = "productos"

    id = Column(Integer, primary_key=True)
    nombre = Column(String(150), nullable=False)
    talla = Column(String(20), nullable=False, default="Única")
    categoria = Column(String(60))
    compra_origen = Column(String(60))
    costo_unitario = Column(Float, default=0.0)
    precio_venta_unitario = Column(Float, default=0.0)
    stock_actual = Column(Integer, default=0)
    imagen_base64 = Column(Text, default=None)  # imagen codificada en base64 (JPEG/PNG pequeños)
    activo = Column(Integer, default=1)  # 1 = activo, 0 = descontinuado (no se borra, se oculta)
    creado_en = Column(DateTime, default=datetime.utcnow)

    __table_args__ = (UniqueConstraint("nombre", "talla", "compra_origen", name="uq_producto_talla_origen"),)

    ventas = relationship("Venta", back_populates="producto")
    compras = relationship("Compra", back_populates="producto")

    @property
    def nombre_completo(self):
        if self.talla and self.talla != "Única":
            return f"{self.nombre} - {self.talla}"
        return self.nombre


class Venta(Base):
    __tablename__ = "ventas"

    id = Column(Integer, primary_key=True)
    producto_id = Column(Integer, ForeignKey("productos.id"), nullable=False)
    fecha = Column(Date, nullable=False, default=date.today)
    cantidad = Column(Integer, nullable=False, default=1)
    canal = Column(String(40))
    cliente_nombre = Column(String(150))
    cliente_documento = Column(String(30))
    valor_venta = Column(Float, default=0.0)       # lo que pagó el cliente (ingreso bruto)
    costo_total = Column(Float, default=0.0)        # costo_unitario × cantidad, guardado como foto del momento
    utilidad = Column(Float, default=0.0)
    retribucion_fondo = Column(Float, default=0.0)
    registrado_en = Column(DateTime, default=datetime.utcnow)

    producto = relationship("Producto", back_populates="ventas")


class Compra(Base):
    __tablename__ = "compras"

    id = Column(Integer, primary_key=True)
    producto_id = Column(Integer, ForeignKey("productos.id"), nullable=False)
    fecha = Column(Date, nullable=False, default=date.today)
    cantidad = Column(Integer, nullable=False)
    costo_unitario = Column(Float, nullable=False)
    costo_total = Column(Float, nullable=False)
    registrado_en = Column(DateTime, default=datetime.utcnow)

    producto = relationship("Producto", back_populates="compras")


class TasaRetribucion(Base):
    __tablename__ = "tasas_retribucion"

    id = Column(Integer, primary_key=True)
    categoria = Column(String(60), unique=True, nullable=False)
    tasa_por_unidad = Column(Float, nullable=False, default=0.0)
