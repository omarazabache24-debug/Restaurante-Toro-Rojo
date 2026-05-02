
# -*- coding: utf-8 -*-
"""
SISTEMA DE CONTROL Y GESTION
Archivo único .py

Incluye:
- Login de usuarios
- Panel principal
- Ventas
- Pedidos
- Inventario
- Recetas
- Caja
- Delivery / clientes
- Reportes
- Administrador
- Log
- Ticket TXT imprimible
- Control de stock detallado por receta

Base de datos: SQLite
"""

import csv
import os
import sqlite3
import base64
import io
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except Exception:
    DateEntry = None
    TKCALENDAR_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except Exception:
    openpyxl = None
    OPENPYXL_AVAILABLE = False


APP_TITLE = "SISTEMA DE CONTROL - GESTIÓN DE ALIMENTOS"
APP_SUBTITLE = "• Control • Procesos • Reportes"
APP_BRAND = "AORIX SYSTEMS - Automatizamos tu empresa"
DB_NAME = "restaurante_v3.db"
APP_DIR = os.path.dirname(os.path.abspath(__file__))

COLOR_NAVY = "#0B2D4A"
COLOR_NAVY_DARK = "#082238"
COLOR_GREEN = "#66D17B"
COLOR_BLUE = "#7CB7E8"
COLOR_WHITE = "#FFFFFF"
COLOR_SOFT = "#F3F4F6"
COLOR_LINE = "#CBD5E1"
COLOR_TEXT = "#111827"

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

AORIX_FALLBACK_LOGO_B64 = """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"""



def ahora():
    return datetime.now()


def fecha_hoy():
    return date.today().isoformat()


def dinero(v):
    try:
        return f"S/ {float(v):,.2f}"
    except Exception:
        return "S/ 0.00"


def numero(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return float(default)


class Database:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()
        self.seed_data()
        self.sync_real_menu_products()

    def execute(self, sql, params=(), commit=False):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        if commit:
            self.conn.commit()
        return cur

    def executemany(self, sql, rows, commit=False):
        cur = self.conn.cursor()
        cur.executemany(sql, rows)
        if commit:
            self.conn.commit()
        return cur

    def row(self, sql, params=()):
        return self.execute(sql, params).fetchone()

    def rows(self, sql, params=()):
        return self.execute(sql, params).fetchall()

    def scalar(self, sql, params=()):
        r = self.row(sql, params)
        if not r:
            return 0
        val = r[0]
        return 0 if val is None else val

    def create_tables(self):
        self.execute("""
        CREATE TABLE IF NOT EXISTS configuracion(
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS usuarios(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            clave TEXT NOT NULL,
            nombre TEXT NOT NULL,
            rol TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS categorias(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS productos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            categoria_id INTEGER,
            tipo TEXT NOT NULL DEFAULT 'VENTA',
            unidad TEXT NOT NULL DEFAULT 'UND',
            precio REAL NOT NULL DEFAULT 0,
            costo REAL NOT NULL DEFAULT 0,
            stock REAL NOT NULL DEFAULT 0,
            stock_min REAL NOT NULL DEFAULT 0,
            activo INTEGER NOT NULL DEFAULT 1,
            FOREIGN KEY(categoria_id) REFERENCES categorias(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS recetas(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producto_id INTEGER UNIQUE NOT NULL,
            observacion TEXT DEFAULT '',
            FOREIGN KEY(producto_id) REFERENCES productos(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS receta_detalle(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            receta_id INTEGER NOT NULL,
            insumo_id INTEGER NOT NULL,
            cantidad REAL NOT NULL DEFAULT 0,
            FOREIGN KEY(receta_id) REFERENCES recetas(id),
            FOREIGN KEY(insumo_id) REFERENCES productos(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS mesas(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            estado TEXT NOT NULL DEFAULT 'LIBRE'
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS clientes(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            telefono TEXT DEFAULT '',
            direccion TEXT DEFAULT '',
            referencia TEXT DEFAULT '',
            notas TEXT DEFAULT ''
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS pedidos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL,
            fecha TEXT NOT NULL,
            hora TEXT NOT NULL,
            mesa_id INTEGER,
            cliente_id INTEGER,
            cliente TEXT DEFAULT '',
            telefono TEXT DEFAULT '',
            direccion TEXT DEFAULT '',
            referencia TEXT DEFAULT '',
            tipo_servicio TEXT NOT NULL,
            estado TEXT NOT NULL DEFAULT 'PENDIENTE',
            total REAL NOT NULL DEFAULT 0,
            pagado INTEGER NOT NULL DEFAULT 0,
            metodo_pago TEXT DEFAULT '',
            usuario TEXT DEFAULT '',
            observacion TEXT DEFAULT '',
            FOREIGN KEY(mesa_id) REFERENCES mesas(id),
            FOREIGN KEY(cliente_id) REFERENCES clientes(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS pedido_detalle(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id INTEGER NOT NULL,
            producto_id INTEGER NOT NULL,
            producto TEXT NOT NULL,
            cantidad REAL NOT NULL,
            precio_unit REAL NOT NULL,
            subtotal REAL NOT NULL,
            FOREIGN KEY(pedido_id) REFERENCES pedidos(id),
            FOREIGN KEY(producto_id) REFERENCES productos(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS caja(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            turno TEXT NOT NULL,
            apertura REAL NOT NULL DEFAULT 0,
            cierre REAL,
            efectivo_sistema REAL NOT NULL DEFAULT 0,
            efectivo_real REAL,
            yape REAL NOT NULL DEFAULT 0,
            tarjeta REAL NOT NULL DEFAULT 0,
            gastos REAL NOT NULL DEFAULT 0,
            estado TEXT NOT NULL DEFAULT 'ABIERTA',
            usuario TEXT DEFAULT '',
            hora_apertura TEXT,
            hora_cierre TEXT
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS movimientos_caja(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            caja_id INTEGER NOT NULL,
            fecha_hora TEXT NOT NULL,
            tipo TEXT NOT NULL,
            concepto TEXT NOT NULL,
            monto REAL NOT NULL,
            metodo TEXT NOT NULL DEFAULT 'EFECTIVO',
            referencia TEXT DEFAULT '',
            FOREIGN KEY(caja_id) REFERENCES caja(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS movimientos_inventario(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_hora TEXT NOT NULL,
            producto_id INTEGER NOT NULL,
            producto TEXT NOT NULL,
            tipo TEXT NOT NULL,
            cantidad REAL NOT NULL,
            referencia TEXT DEFAULT '',
            costo_unit REAL NOT NULL DEFAULT 0,
            FOREIGN KEY(producto_id) REFERENCES productos(id)
        )
        """, commit=True)

        self.execute("""
        CREATE TABLE IF NOT EXISTS cierres_dia(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT UNIQUE NOT NULL,
            sucursal TEXT DEFAULT '',
            turno TEXT DEFAULT '',
            usuario TEXT DEFAULT '',
            hora_cierre TEXT NOT NULL,
            total_ventas REAL NOT NULL DEFAULT 0,
            total_pedidos INTEGER NOT NULL DEFAULT 0,
            total_pendientes INTEGER NOT NULL DEFAULT 0,
            efectivo REAL NOT NULL DEFAULT 0,
            yape REAL NOT NULL DEFAULT 0,
            tarjeta REAL NOT NULL DEFAULT 0,
            gastos REAL NOT NULL DEFAULT 0,
            estado TEXT NOT NULL DEFAULT 'CERRADO',
            observacion TEXT DEFAULT ''
        )
        """, commit=True)

    def seed_data(self):
        if self.scalar("SELECT COUNT(*) FROM usuarios") == 0:
            self.executemany("""
                INSERT INTO usuarios(usuario, clave, nombre, rol, activo)
                VALUES (?, ?, ?, ?, 1)
            """, [
                ("admin", "1234", "Administrador", "ADMIN"),
                ("caja", "1234", "Caja Principal", "CAJA"),
                ("mesero1", "1234", "Mesero 1", "MESERO"),
            ], commit=True)

        if self.scalar("SELECT COUNT(*) FROM categorias") == 0:
            self.executemany("INSERT INTO categorias(nombre) VALUES (?)", [
                ("PLATOS",), ("BEBIDAS",), ("INSUMOS",), ("POSTRES",), ("COMPLEMENTOS",)
            ], commit=True)

        if self.scalar("SELECT COUNT(*) FROM mesas") == 0:
            self.executemany("INSERT INTO mesas(nombre, estado) VALUES (?,?)", [
                (f"MESA {i}", "LIBRE") for i in range(1, 11)
            ], commit=True)

        if self.scalar("SELECT COUNT(*) FROM configuracion WHERE clave='sucursal'") == 0:
            self.execute("INSERT INTO configuracion(clave, valor) VALUES ('sucursal','Sucursal Principal')", commit=True)

        if self.scalar("SELECT COUNT(*) FROM productos") == 0:
            cat = {r["nombre"]: r["id"] for r in self.rows("SELECT id, nombre FROM categorias")}
            productos = [
                ("POLLO A LA BRASA", cat["PLATOS"], "VENTA", "PLATO", 45.00, 18.00, 50, 5, 1),
                ("1/2 POLLO", cat["PLATOS"], "VENTA", "PLATO", 26.00, 10.00, 40, 5, 1),
                ("GASEOSA 500ML", cat["BEBIDAS"], "VENTA", "UND", 4.50, 2.20, 120, 12, 1),
                ("INKA KOLA 1.5L", cat["BEBIDAS"], "VENTA", "UND", 9.00, 5.50, 60, 8, 1),
                ("ENSALADA EXTRA", cat["COMPLEMENTOS"], "VENTA", "UND", 5.00, 1.50, 30, 5, 1),
                ("POLLO ENTERO", cat["INSUMOS"], "INSUMO", "UND", 0, 14.00, 120, 15, 1),
                ("PAPA KG", cat["INSUMOS"], "INSUMO", "KG", 0, 3.80, 100, 10, 1),
                ("ENSALADA", cat["INSUMOS"], "INSUMO", "PORCION", 0, 1.20, 80, 10, 1),
                ("MAYONESA", cat["INSUMOS"], "INSUMO", "PORCION", 0, 0.40, 200, 20, 1),
                ("AJI", cat["INSUMOS"], "INSUMO", "PORCION", 0, 0.35, 200, 20, 1),
                ("KETCHUP", cat["INSUMOS"], "INSUMO", "PORCION", 0, 0.30, 200, 20, 1),
                ("MOSTAZA", cat["INSUMOS"], "INSUMO", "PORCION", 0, 0.30, 200, 20, 1),
                ("ENVASE", cat["INSUMOS"], "INSUMO", "UND", 0, 0.45, 200, 20, 1),
                ("BOLSA DELIVERY", cat["INSUMOS"], "INSUMO", "UND", 0, 0.25, 120, 10, 1),
                ("BEBIDA 500ML STOCK", cat["INSUMOS"], "INSUMO", "UND", 0, 2.20, 120, 12, 1),
            ]
            self.executemany("""
                INSERT INTO productos(nombre, categoria_id, tipo, unidad, precio, costo, stock, stock_min, activo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, productos, commit=True)

            self.crear_receta_demo("POLLO A LA BRASA", [
                ("POLLO ENTERO", 1),
                ("PAPA KG", 0.50),
                ("ENSALADA", 1),
                ("MAYONESA", 2),
                ("AJI", 2),
                ("KETCHUP", 1),
                ("MOSTAZA", 1),
                ("ENVASE", 1),
            ])
            self.crear_receta_demo("1/2 POLLO", [
                ("POLLO ENTERO", 0.50),
                ("PAPA KG", 0.25),
                ("ENSALADA", 1),
                ("MAYONESA", 1),
                ("AJI", 1),
                ("ENVASE", 1),
            ])
            self.crear_receta_demo("ENSALADA EXTRA", [
                ("ENSALADA", 1),
                ("ENVASE", 1),
            ])
            self.crear_receta_demo("GASEOSA 500ML", [
                ("BEBIDA 500ML STOCK", 1),
            ])

    def crear_receta_demo(self, producto_nombre, detalles):
        prod = self.row("SELECT id FROM productos WHERE nombre=?", (producto_nombre,))
        if not prod:
            return
        r = self.row("SELECT id FROM recetas WHERE producto_id=?", (prod["id"],))
        if r:
            receta_id = r["id"]
        else:
            receta_id = self.execute(
                "INSERT INTO recetas(producto_id, observacion) VALUES (?, '')",
                (prod["id"],), commit=True
            ).lastrowid
        for insumo_nombre, cantidad in detalles:
            ins = self.row("SELECT id FROM productos WHERE nombre=?", (insumo_nombre,))
            if ins:
                self.execute(
                    "INSERT INTO receta_detalle(receta_id, insumo_id, cantidad) VALUES (?, ?, ?)",
                    (receta_id, ins["id"], cantidad), commit=True
                )


    def sync_real_menu_products(self):
        categorias_base = [
            "PLATOS", "BEBIDAS", "INSUMOS", "POSTRES", "COMPLEMENTOS",
            "PIZZAS", "JUGOS", "TRAGOS", "CERVEZAS", "SELVATICOS",
            "HAMBURGUESAS", "PARRILLAS", "COMBOS", "FASTFOOD", "ADICIONALES"
        ]
        for cat in categorias_base:
            if not self.row("SELECT id FROM categorias WHERE nombre=?", (cat,)):
                self.execute("INSERT INTO categorias(nombre) VALUES (?)", (cat,), commit=True)

        categorias = {r["nombre"]: r["id"] for r in self.rows("SELECT id, nombre FROM categorias")}

        productos_menu = [
            {"nombre": "AGUA SAN LUIS", "categoria": "BEBIDAS", "precio": 3.00, "unidad": "UND"},
            {"nombre": "AGUA SAN CARLOS", "categoria": "BEBIDAS", "precio": 2.00, "unidad": "UND"},
            {"nombre": "CAFE", "categoria": "BEBIDAS", "precio": 4.00, "unidad": "UND"},
            {"nombre": "INFUSIONES", "categoria": "BEBIDAS", "precio": 4.00, "unidad": "UND"},
            {"nombre": "CUSQUEÑA", "categoria": "CERVEZAS", "precio": 10.00, "unidad": "UND"},
            {"nombre": "CUSQUEÑA NEGRA", "categoria": "CERVEZAS", "precio": 10.00, "unidad": "UND"},

            {"nombre": "PIZZA AMERICANA PERSONAL", "categoria": "PIZZAS", "precio": 13.00, "unidad": "UND"},
            {"nombre": "PIZZA AMERICANA GRANDE", "categoria": "PIZZAS", "precio": 32.00, "unidad": "UND"},
            {"nombre": "PIZZA AMERICANA FAMILIAR", "categoria": "PIZZAS", "precio": 42.00, "unidad": "UND"},
            {"nombre": "PIZZA ESPAÑOLA PERSONAL", "categoria": "PIZZAS", "precio": 14.00, "unidad": "UND"},
            {"nombre": "PIZZA ESPAÑOLA GRANDE", "categoria": "PIZZAS", "precio": 32.00, "unidad": "UND"},
            {"nombre": "PIZZA ESPAÑOLA FAMILIAR", "categoria": "PIZZAS", "precio": 40.00, "unidad": "UND"},
            {"nombre": "PIZZA HAWAIANA PERSONAL", "categoria": "PIZZAS", "precio": 14.00, "unidad": "UND"},
            {"nombre": "PIZZA HAWAIANA GRANDE", "categoria": "PIZZAS", "precio": 35.00, "unidad": "UND"},
            {"nombre": "PIZZA HAWAIANA FAMILIAR", "categoria": "PIZZAS", "precio": 43.00, "unidad": "UND"},
            {"nombre": "PIZZA PEPERONI PERSONAL", "categoria": "PIZZAS", "precio": 14.00, "unidad": "UND"},
            {"nombre": "PIZZA PEPERONI GRANDE", "categoria": "PIZZAS", "precio": 35.00, "unidad": "UND"},
            {"nombre": "PIZZA PEPERONI FAMILIAR", "categoria": "PIZZAS", "precio": 43.00, "unidad": "UND"},
            {"nombre": "PIZZA POLLO ITALIANO PERSONAL", "categoria": "PIZZAS", "precio": 15.00, "unidad": "UND"},
            {"nombre": "PIZZA POLLO ITALIANO GRANDE", "categoria": "PIZZAS", "precio": 38.00, "unidad": "UND"},
            {"nombre": "PIZZA POLLO ITALIANO FAMILIAR", "categoria": "PIZZAS", "precio": 45.00, "unidad": "UND"},
            {"nombre": "PIZZA COSA NOSTRA PERSONAL", "categoria": "PIZZAS", "precio": 16.00, "unidad": "UND"},
            {"nombre": "PIZZA COSA NOSTRA GRANDE", "categoria": "PIZZAS", "precio": 40.00, "unidad": "UND"},
            {"nombre": "PIZZA COSA NOSTRA FAMILIAR", "categoria": "PIZZAS", "precio": 49.00, "unidad": "UND"},
            {"nombre": "PIZZA SUPREMA PERSONAL", "categoria": "PIZZAS", "precio": 18.00, "unidad": "UND"},
            {"nombre": "PIZZA SUPREMA GRANDE", "categoria": "PIZZAS", "precio": 45.00, "unidad": "UND"},
            {"nombre": "PIZZA SUPREMA FAMILIAR", "categoria": "PIZZAS", "precio": 55.00, "unidad": "UND"},

            {"nombre": "JUGO MIXTO", "categoria": "JUGOS", "precio": 6.00, "unidad": "VASO"},
            {"nombre": "JUGO PAPAYA", "categoria": "JUGOS", "precio": 7.00, "unidad": "VASO"},
            {"nombre": "LECHE CON FRESA", "categoria": "JUGOS", "precio": 10.00, "unidad": "VASO"},
            {"nombre": "JUGO DE FRESA", "categoria": "JUGOS", "precio": 9.00, "unidad": "VASO"},
            {"nombre": "PONCHE", "categoria": "JUGOS", "precio": 12.00, "unidad": "VASO"},
            {"nombre": "FRESA CON ARANDANOS", "categoria": "JUGOS", "precio": 10.00, "unidad": "VASO"},
            {"nombre": "MOJITO", "categoria": "TRAGOS", "precio": 13.00, "unidad": "VASO"},
            {"nombre": "PISCO SOUR", "categoria": "TRAGOS", "precio": 14.00, "unidad": "VASO"},
            {"nombre": "LAGUNA AZUL", "categoria": "TRAGOS", "precio": 13.00, "unidad": "VASO"},
            {"nombre": "ORANGE PARADISE", "categoria": "TRAGOS", "precio": 12.00, "unidad": "VASO"},
            {"nombre": "COCTEL DE FRESA", "categoria": "TRAGOS", "precio": 14.00, "unidad": "VASO"},
            {"nombre": "NEGRONI", "categoria": "TRAGOS", "precio": 10.00, "unidad": "VASO"},
            {"nombre": "COCTEL DE ARANDANOS", "categoria": "TRAGOS", "precio": 14.00, "unidad": "VASO"},

            {"nombre": "JUANE", "categoria": "SELVATICOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "TACACHO CON CECINA Y CHORIZO", "categoria": "SELVATICOS", "precio": 24.00, "unidad": "PLATO"},
            {"nombre": "CHAUFA AMAZONICO", "categoria": "SELVATICOS", "precio": 24.00, "unidad": "PLATO"},
            {"nombre": "ALITAS CON SALSA DE COCONA", "categoria": "SELVATICOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "PARRILLA SELVATICA", "categoria": "SELVATICOS", "precio": 19.00, "unidad": "PLATO"},
            {"nombre": "CECINA SALTEADA", "categoria": "SELVATICOS", "precio": 24.00, "unidad": "PLATO"},
            {"nombre": "CANASTAS RELLENAS", "categoria": "SELVATICOS", "precio": 15.00, "unidad": "PLATO"},
            {"nombre": "TEQUEÑOS AMAZONICOS", "categoria": "SELVATICOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "BROCHETAS AMAZONICAS", "categoria": "SELVATICOS", "precio": 22.00, "unidad": "PLATO"},

            {"nombre": "LA TRADICIONAL", "categoria": "HAMBURGUESAS", "precio": 9.90, "unidad": "UND"},
            {"nombre": "EL BOCON", "categoria": "HAMBURGUESAS", "precio": 12.90, "unidad": "UND"},
            {"nombre": "LA HAWAIANA", "categoria": "HAMBURGUESAS", "precio": 13.90, "unidad": "UND"},
            {"nombre": "LA PATRIOTA", "categoria": "HAMBURGUESAS", "precio": 14.90, "unidad": "UND"},
            {"nombre": "EL GLADIADOR", "categoria": "HAMBURGUESAS", "precio": 17.90, "unidad": "UND"},
            {"nombre": "DOBLE", "categoria": "HAMBURGUESAS", "precio": 13.90, "unidad": "UND"},
            {"nombre": "BURGUER CHEESE", "categoria": "HAMBURGUESAS", "precio": 12.90, "unidad": "UND"},
            {"nombre": "FILETE DE POLLO", "categoria": "HAMBURGUESAS", "precio": 10.00, "unidad": "UND"},
            {"nombre": "CHIKEN CHEESE", "categoria": "HAMBURGUESAS", "precio": 13.90, "unidad": "UND"},
            {"nombre": "ADICIONAL CARNE", "categoria": "ADICIONALES", "precio": 5.00, "unidad": "UND"},
            {"nombre": "ADICIONAL TOCINO", "categoria": "ADICIONALES", "precio": 5.00, "unidad": "UND"},
            {"nombre": "ADICIONAL PIÑA", "categoria": "ADICIONALES", "precio": 3.00, "unidad": "UND"},
            {"nombre": "ENSALADA PORCION", "categoria": "ADICIONALES", "precio": 10.00, "unidad": "PORCION"},
            {"nombre": "PAPA PORCION", "categoria": "ADICIONALES", "precio": 12.00, "unidad": "PORCION"},
            {"nombre": "PAPA 1/2 PORCION", "categoria": "ADICIONALES", "precio": 8.00, "unidad": "PORCION"},
            {"nombre": "CHAUFA PORCION", "categoria": "ADICIONALES", "precio": 10.00, "unidad": "PORCION"},
            {"nombre": "ARROZ BLANCO PORCION", "categoria": "ADICIONALES", "precio": 5.00, "unidad": "PORCION"},
            {"nombre": "HUEVO 1 UND", "categoria": "ADICIONALES", "precio": 2.00, "unidad": "UND"},
            {"nombre": "QUESO CHEDAR", "categoria": "ADICIONALES", "precio": 3.00, "unidad": "UND"},

            {"nombre": "COMBO POLLO ENTERO + GASEOSA 1LT", "categoria": "COMBOS", "precio": 69.90, "unidad": "COMBO"},
            {"nombre": "COMBO POLLO ENTERO", "categoria": "COMBOS", "precio": 59.90, "unidad": "COMBO"},
            {"nombre": "COMBO 1/2 POLLO + GASEOSA 1LT", "categoria": "COMBOS", "precio": 44.90, "unidad": "COMBO"},
            {"nombre": "COMBO 1/2 POLLO", "categoria": "COMBOS", "precio": 37.90, "unidad": "COMBO"},
            {"nombre": "TALLARIN SALTADO POLLO", "categoria": "PLATOS", "precio": 16.00, "unidad": "PLATO"},
            {"nombre": "TALLARIN SALTADO CARNE", "categoria": "PLATOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "LOMO SALTADO POLLO", "categoria": "PLATOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "LOMO SALTADO CARNE", "categoria": "PLATOS", "precio": 20.00, "unidad": "PLATO"},
            {"nombre": "FILETE DE POLLO A LA PARRILLA", "categoria": "PLATOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "CHICHARRON DE POLLO", "categoria": "PLATOS", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "CALDO DE GALLINA", "categoria": "PLATOS", "precio": 13.90, "unidad": "PLATO"},

            {"nombre": "POLLO A LA BRASA", "categoria": "PLATOS", "precio": 56.00, "unidad": "PLATO"},
            {"nombre": "1/2 POLLO", "categoria": "PLATOS", "precio": 35.00, "unidad": "PLATO"},
            {"nombre": "1/4 POLLO", "categoria": "PLATOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "1/8 POLLO", "categoria": "PLATOS", "precio": 12.00, "unidad": "PLATO"},
            {"nombre": "MOSTRO", "categoria": "COMBOS", "precio": 22.00, "unidad": "PLATO"},
            {"nombre": "MOSTRITO", "categoria": "COMBOS", "precio": 16.00, "unidad": "PLATO"},
            {"nombre": "CHIWUAN", "categoria": "COMBOS", "precio": 17.00, "unidad": "PLATO"},
            {"nombre": "CHAUFA DE POLLO", "categoria": "PLATOS", "precio": 16.00, "unidad": "PLATO"},
            {"nombre": "CHAUFA DE CARNE/CERDO", "categoria": "PLATOS", "precio": 18.00, "unidad": "PLATO"},
            {"nombre": "AEROPUERTO", "categoria": "PLATOS", "precio": 19.90, "unidad": "PLATO"},
            {"nombre": "AEROPUERTO ESPECIAL", "categoria": "PLATOS", "precio": 23.90, "unidad": "PLATO"},

            {"nombre": "POLLO ENTERO BROSTER", "categoria": "PLATOS", "precio": 75.00, "unidad": "PLATO"},
            {"nombre": "1/2 POLLO BROSTER", "categoria": "PLATOS", "precio": 45.00, "unidad": "PLATO"},
            {"nombre": "1/4 POLLO BROSTER", "categoria": "PLATOS", "precio": 22.00, "unidad": "PLATO"},
            {"nombre": "1/8 POLLO BROSTER", "categoria": "PLATOS", "precio": 15.00, "unidad": "PLATO"},
            {"nombre": "COMBO POLLO ENTERO DE BROSTER", "categoria": "COMBOS", "precio": 82.00, "unidad": "COMBO"},
            {"nombre": "COMBO 1/2 POLLO BROSTER", "categoria": "COMBOS", "precio": 50.00, "unidad": "COMBO"},
            {"nombre": "COMBO 1/4 POLLO BROSTER", "categoria": "COMBOS", "precio": 27.00, "unidad": "COMBO"},
            {"nombre": "COMBO 1/8 POLLO BROSTER", "categoria": "COMBOS", "precio": 17.00, "unidad": "COMBO"},

            {"nombre": "1/4 DE POLLO PARRILLERO", "categoria": "PARRILLAS", "precio": 18.90, "unidad": "PLATO"},
            {"nombre": "CHULETA 300 GR", "categoria": "PARRILLAS", "precio": 24.00, "unidad": "PLATO"},
            {"nombre": "CHURRASCO 300 GR", "categoria": "PARRILLAS", "precio": 24.00, "unidad": "PLATO"},
            {"nombre": "TIRA DE CERDO 450 GR", "categoria": "PARRILLAS", "precio": 45.00, "unidad": "PLATO"},
            {"nombre": "TIRA DE CARNE 450 GR", "categoria": "PARRILLAS", "precio": 45.00, "unidad": "PLATO"},
            {"nombre": "CHORIZO DE FINAS HIERBAS (2 UND)", "categoria": "PARRILLAS", "precio": 12.00, "unidad": "PLATO"},
            {"nombre": "RACHI", "categoria": "PARRILLAS", "precio": 12.00, "unidad": "PLATO"},
            {"nombre": "ANTICUCHOS DE CORAZON (2 PALITOS)", "categoria": "PARRILLAS", "precio": 17.00, "unidad": "PLATO"},
            {"nombre": "MOLLEJAS A LA PARRILLA", "categoria": "PARRILLAS", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "TOMAHAWK", "categoria": "PARRILLAS", "precio": 45.00, "unidad": "PLATO"},
            {"nombre": "COMBO PARRILLERO MIXTO", "categoria": "COMBOS", "precio": 38.00, "unidad": "COMBO"},
            {"nombre": "COMBO PARRILLERO LA ANTOJADA", "categoria": "COMBOS", "precio": 50.00, "unidad": "COMBO"},
            {"nombre": "COMBO PARRILLERO LA RANCHERA", "categoria": "COMBOS", "precio": 85.00, "unidad": "COMBO"},
            {"nombre": "ESPECIAL EL ROJO", "categoria": "COMBOS", "precio": 199.90, "unidad": "COMBO"},
            {"nombre": "TRIO ANTICUCHERO", "categoria": "COMBOS", "precio": 24.00, "unidad": "COMBO"},
            {"nombre": "RACHI ANTICUCHERO", "categoria": "COMBOS", "precio": 16.00, "unidad": "COMBO"},

            {"nombre": "SALCHIPAPA CLASICA", "categoria": "FASTFOOD", "precio": 6.90, "unidad": "PLATO"},
            {"nombre": "SALCHIPAPA MIXTA", "categoria": "FASTFOOD", "precio": 9.90, "unidad": "PLATO"},
            {"nombre": "SALCHIPAPA MIXTURA", "categoria": "FASTFOOD", "precio": 16.90, "unidad": "PLATO"},
            {"nombre": "SALCHIPAPA CHEESE", "categoria": "FASTFOOD", "precio": 12.00, "unidad": "PLATO"},
            {"nombre": "SALCHIPOLLO", "categoria": "FASTFOOD", "precio": 15.00, "unidad": "PLATO"},
            {"nombre": "SALCHIPOLLO COMBO", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "ALITAS A LA BBQ", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "ALITAS BROSTERAS", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "ALITAS PICANTES", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "ALITAS ACEVICHADAS", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},
            {"nombre": "ALITAS ANTICUCHERAS", "categoria": "FASTFOOD", "precio": 17.90, "unidad": "PLATO"},

            {"nombre": "CHICHA MORADA 1 JARRA", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "JARRA"},
            {"nombre": "CHICHA MORADA 1/2 JARRA", "categoria": "BEBIDAS", "precio": 8.00, "unidad": "JARRA"},
            {"nombre": "LIMONADA FROZEN 1 JARRA", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "JARRA"},
            {"nombre": "LIMONADA FROZEN 1/2 JARRA", "categoria": "BEBIDAS", "precio": 8.00, "unidad": "JARRA"},
            {"nombre": "SANGRIA 1 JARRA", "categoria": "BEBIDAS", "precio": 35.00, "unidad": "JARRA"},
            {"nombre": "SANGRIA 1/2 JARRA", "categoria": "BEBIDAS", "precio": 24.00, "unidad": "JARRA"},
            {"nombre": "PIÑA FROZEN 1 JARRA", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "JARRA"},
            {"nombre": "PIÑA FROZEN 1/2 JARRA", "categoria": "BEBIDAS", "precio": 8.00, "unidad": "JARRA"},
            {"nombre": "MARACUYA FROZEN 1 JARRA", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "JARRA"},
            {"nombre": "MARACUYA FROZEN 1/2 JARRA", "categoria": "BEBIDAS", "precio": 8.00, "unidad": "JARRA"},
            {"nombre": "MOJITO 1 JARRA", "categoria": "BEBIDAS", "precio": 35.00, "unidad": "JARRA"},
            {"nombre": "MOJITO 1/2 JARRA", "categoria": "BEBIDAS", "precio": 24.00, "unidad": "JARRA"},
            {"nombre": "PEPSI 2 L", "categoria": "BEBIDAS", "precio": 10.00, "unidad": "UND"},
            {"nombre": "PEPSI 1 L", "categoria": "BEBIDAS", "precio": 5.00, "unidad": "UND"},
            {"nombre": "COCA-COLA 2 L", "categoria": "BEBIDAS", "precio": 10.00, "unidad": "UND"},
            {"nombre": "COCA-COLA 1.5 L", "categoria": "BEBIDAS", "precio": 9.00, "unidad": "UND"},
            {"nombre": "COCA-COLA 1 L", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "UND"},
            {"nombre": "INKA-COLA 2 L", "categoria": "BEBIDAS", "precio": 10.00, "unidad": "UND"},
            {"nombre": "INKA-COLA 1.5 L", "categoria": "BEBIDAS", "precio": 9.00, "unidad": "UND"},
            {"nombre": "INKA-COLA 1 L", "categoria": "BEBIDAS", "precio": 12.00, "unidad": "UND"},
            {"nombre": "PEPSI 500 ML", "categoria": "BEBIDAS", "precio": 4.00, "unidad": "UND"},
            {"nombre": "COCA-COLA 500 ML", "categoria": "BEBIDAS", "precio": 4.00, "unidad": "UND"},
            {"nombre": "INKA-COLA 500 ML", "categoria": "BEBIDAS", "precio": 4.00, "unidad": "UND"},
            {"nombre": "PEPSI 450 ML", "categoria": "BEBIDAS", "precio": 3.00, "unidad": "UND"},
            {"nombre": "INKA-COLA / COCA-COLA PERSONAL", "categoria": "BEBIDAS", "precio": 5.00, "unidad": "UND"},
        ]

        for item in productos_menu:
            nombre = item["nombre"].strip().upper()
            categoria = item["categoria"].strip().upper()
            categoria_id = categorias[categoria]
            unidad = item.get("unidad", "UND").strip().upper()
            tipo = item.get("tipo", "VENTA").strip().upper()
            precio = float(item.get("precio", 0) or 0)
            stock_min = float(item.get("stock_min", 0) or 0)

            existente = self.row("SELECT * FROM productos WHERE nombre=?", (nombre,))
            if existente:
                costo = float(existente["costo"] or 0)
                stock = float(existente["stock"] or 0)
                stock_min_final = float(existente["stock_min"] or stock_min)
                self.execute("""
                    UPDATE productos
                    SET categoria_id=?, tipo=?, unidad=?, precio=?, costo=?, stock=?, stock_min=?, activo=1
                    WHERE id=?
                """, (categoria_id, tipo, unidad, precio, costo, stock, stock_min_final, existente["id"]), commit=True)
            else:
                self.execute("""
                    INSERT INTO productos(nombre, categoria_id, tipo, unidad, precio, costo, stock, stock_min, activo)
                    VALUES (?, ?, ?, ?, ?, ?, 0, ?, 1)
                """, (nombre, categoria_id, tipo, unidad, precio, 0.0, stock_min), commit=True)


class LoginWindow:
    def __init__(self, db):
        self.db = db
        self.result = None
        self.win = tk.Tk()
        self.win.title("Ingreso al sistema")
        self.win.geometry("420x290")
        self.win.resizable(False, False)
        self.win.configure(bg="#F3F4F6")

        box = tk.Frame(self.win, bg="#FFFFFF", bd=1, relief="solid")
        box.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(box, text="SISTEMA DE CONTROL", bg="#FFFFFF", font=("Arial", 18, "bold")).pack(pady=(20, 6))
        tk.Label(box, text="Ingrese sus credenciales", bg="#FFFFFF", fg="#4B5563", font=("Arial", 11)).pack(pady=(0, 16))

        self.usuario_var = tk.StringVar(value="admin")
        self.clave_var = tk.StringVar(value="1234")

        f1 = tk.Frame(box, bg="#FFFFFF")
        f1.pack(fill="x", padx=30, pady=6)
        tk.Label(f1, text="Usuario", bg="#FFFFFF", anchor="w", font=("Arial", 10, "bold")).pack(fill="x")
        tk.Entry(f1, textvariable=self.usuario_var).pack(fill="x", pady=(4, 0))

        f2 = tk.Frame(box, bg="#FFFFFF")
        f2.pack(fill="x", padx=30, pady=6)
        tk.Label(f2, text="Clave", bg="#FFFFFF", anchor="w", font=("Arial", 10, "bold")).pack(fill="x")
        tk.Entry(f2, textvariable=self.clave_var, show="*").pack(fill="x", pady=(4, 0))

        tk.Button(box, text="Ingresar", width=18, command=self.ingresar).pack(pady=18)
        tk.Label(box, text="Usuario demo: admin | Clave: 1234", bg="#FFFFFF", fg="#6B7280", font=("Arial", 9)).pack()

        self.win.bind("<Return>", lambda e: self.ingresar())

    def ingresar(self):
        usuario = self.usuario_var.get().strip()
        clave = self.clave_var.get().strip()
        row = self.db.row("""
            SELECT * FROM usuarios
            WHERE usuario=? AND clave=? AND activo=1
        """, (usuario, clave))
        if not row:
            messagebox.showerror("Ingreso", "Usuario o clave incorrectos.")
            return
        self.result = row
        self.win.destroy()

    def run(self):
        self.win.mainloop()
        return self.result


class RestauranteAppV3:
    def __init__(self, db, user_row):
        self.db = db
        self.user = user_row
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        # Carpetas automáticas de trabajo. No dependen del nombre del .py.
        self.dir_sistema_control = os.path.join(self.base_dir, "01_SISTEMA_CONTROL_RESTAURANTE")
        self.dir_control_contable = os.path.join(self.base_dir, "02_CONTROL_CONTABLE")
        self.tickets_dir = os.path.join(self.dir_sistema_control, "TICKETS")
        for _d in (self.dir_sistema_control, self.dir_control_contable, self.tickets_dir):
            os.makedirs(_d, exist_ok=True)

        self.root = tk.Tk()
        self.root.title("Sistema Cliente - AORIX SYSTEMS")
        self.root.geometry("1710x980")
        self.root.minsize(1400, 820)
        try:
            self.root.state("zoomed")
        except Exception:
            pass
        self.root.configure(bg=COLOR_SOFT)

        self.sucursal_var = tk.StringVar(value=self.get_config("sucursal", "Sucursal Principal"))
        self.fecha_var = tk.StringVar(value=fecha_hoy())
        self.turno_var = tk.StringVar(value="MAÑANA")
        self.usuario_var = tk.StringVar(value=self.user["usuario"])
        self.usuario_nombre_var = tk.StringVar(value=self.user["nombre"])
        self.rol_var = tk.StringVar(value=self.user["rol"])
        self.servicio_var = tk.StringVar(value="SALÓN")
        self.status_var = tk.StringVar(value="Listo")

        self.venta_mesa_var = tk.StringVar()
        self.venta_tipo_var = tk.StringVar(value="SALÓN")
        self.venta_cliente_var = tk.StringVar()
        self.venta_telefono_var = tk.StringVar()
        self.venta_direccion_var = tk.StringVar()
        self.venta_referencia_var = tk.StringVar()
        self.venta_producto_var = tk.StringVar()
        self.venta_busqueda_producto_var = tk.StringVar()
        self.venta_cantidad_var = tk.StringVar(value="1")
        self.venta_metodo_pago_var = tk.StringVar(value="EFECTIVO")
        self.venta_descuento_var = tk.StringVar(value="0.00")
        self.current_order_items = []
        self.selected_pedido_id = None
        self.selected_related_pedido_ids = []
        self._seleccionando_pedidos = False

        self.inv_nombre_var = tk.StringVar()
        self.inv_categoria_var = tk.StringVar()
        self.inv_tipo_var = tk.StringVar(value="VENTA")
        self.inv_unidad_var = tk.StringVar(value="UND")
        self.inv_precio_var = tk.StringVar(value="0.00")
        self.inv_costo_var = tk.StringVar(value="0.00")
        self.inv_stock_var = tk.StringVar(value="0")
        self.inv_min_var = tk.StringVar(value="0")
        self.inv_alerta_var = tk.StringVar(value="Sin alertas de stock.")

        self.receta_producto_var = tk.StringVar()
        self.receta_insumo_var = tk.StringVar()
        self.receta_cantidad_var = tk.StringVar(value="1")
        self.receta_obs_var = tk.StringVar()

        self.caja_apertura_var = tk.StringVar(value="100.00")
        self.caja_egreso_concepto_var = tk.StringVar()
        self.caja_egreso_monto_var = tk.StringVar(value="0.00")

        self.pedido_estado_var = tk.StringVar(value="TODOS")
        self.reporte_fecha_inicio_var = tk.StringVar(value=fecha_hoy())
        self.reporte_fecha_fin_var = tk.StringVar(value=fecha_hoy())

        # Variables para tablero de indicadores con gráficos
        self.ind_fecha_inicio_var = tk.StringVar(value=fecha_hoy())
        self.ind_fecha_fin_var = tk.StringVar(value=fecha_hoy())
        self.ind_agrupacion_var = tk.StringVar(value="DÍA")
        self.ind_tipo_grafica_var = tk.StringVar(value="VENTAS S/")

        self.cli_nombre_var = tk.StringVar()
        self.cli_telefono_var = tk.StringVar()
        self.cli_direccion_var = tk.StringVar()
        self.cli_referencia_var = tk.StringVar()
        self.cli_notas_var = tk.StringVar()

        self.nueva_sucursal_var = tk.StringVar()
        self.nuevo_usuario_var = tk.StringVar()
        self.nuevo_nombre_usuario_var = tk.StringVar()
        self.nueva_clave_usuario_var = tk.StringVar()
        self.nuevo_rol_usuario_var = tk.StringVar(value="MESERO")

        self.selected_related_pedido_ids = []
        self.selected_detail_line = None

        self.configurar_mayusculas_automaticas()
        self.crear_estilos()
        self.crear_ui()
        self.aplicar_permisos()
        self.sincronizar_contexto_usuario()
        self.refrescar_todo(inicial=True)

    def get_config(self, clave, default=""):
        r = self.db.row("SELECT valor FROM configuracion WHERE clave=?", (clave,))
        return r["valor"] if r else default

    def set_config(self, clave, valor):
        self.db.execute("INSERT OR REPLACE INTO configuracion(clave, valor) VALUES (?, ?)", (clave, str(valor)), commit=True)

    def normalizar_texto(self, valor):
        return str(valor or "").strip().upper()

    def usuario_clave_valida(self, clave):
        return self.normalizar_texto(clave) == str(self.user["clave"]).strip().upper()

    def pedir_clave_actual(self, titulo="Seguridad"):
        win = tk.Toplevel(self.root)
        win.title(titulo)
        win.transient(self.root)
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="Ingrese la clave del sistema:", font=("Arial", 10, "bold")).pack(padx=18, pady=(16, 8))
        var = tk.StringVar()
        ent = tk.Entry(win, textvariable=var, show="*", width=26)
        ent.pack(padx=18, pady=(0, 12))
        result = {"ok": False}
        def aceptar():
            if self.usuario_clave_valida(var.get()):
                result["ok"] = True
                win.destroy()
            else:
                messagebox.showerror(titulo, "Clave incorrecta.", parent=win)
        tk.Button(win, text="Aceptar", width=12, command=aceptar).pack(side="left", padx=(18, 8), pady=(0, 16))
        tk.Button(win, text="Cancelar", width=12, command=win.destroy).pack(side="left", padx=8, pady=(0, 16))
        ent.focus_set()
        ent.bind("<Return>", lambda e: aceptar())
        self.root.wait_window(win)
        return result["ok"]

    def actualizar_total_pedido(self, pedido_id):
        total = self.db.scalar("SELECT COALESCE(SUM(subtotal),0) FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
        self.db.execute("UPDATE pedidos SET total=? WHERE id=?", (float(total), pedido_id), commit=True)
        return float(total)

    def buscar_pedido_abierto_relacionado(self, cliente, mesa_id, tipo):
        cliente = self.normalizar_texto(cliente)
        tipo = self.normalizar_texto(tipo)
        if not cliente:
            return None
        sql = """
            SELECT *
            FROM pedidos
            WHERE UPPER(cliente)=?
              AND COALESCE(mesa_id,0)=COALESCE(?,0)
              AND UPPER(tipo_servicio)=?
              AND fecha=?
              AND pagado=0
              AND estado NOT IN ('ANULADO','PAGADO')
            ORDER BY id DESC
            LIMIT 1
        """
        return self.db.row(sql, (cliente, mesa_id, tipo, self.fecha_trabajo()))

    def pedido_ids_relacionados(self, pedido_id, incluir_pagados=False):
        pedido = self.db.row("SELECT * FROM pedidos WHERE id=?", (pedido_id,))
        if not pedido:
            return []
        sql = """
            SELECT id
            FROM pedidos
            WHERE UPPER(cliente)=UPPER(?)
              AND COALESCE(mesa_id,0)=COALESCE(?,0)
              AND UPPER(tipo_servicio)=UPPER(?)
        """
        params = [pedido["cliente"], pedido["mesa_id"], pedido["tipo_servicio"]]
        if not incluir_pagados:
            sql += " AND pagado=0 AND estado NOT IN ('ANULADO','PAGADO')"
        sql += " ORDER BY id"
        return [r["id"] for r in self.db.rows(sql, tuple(params))]


    def forzar_mayusculas(self, var):
        bandera = {"busy": False}

        def convertir(*args):
            if bandera["busy"]:
                return
            valor = var.get()
            valor_mayus = valor.upper()
            if valor != valor_mayus:
                bandera["busy"] = True
                try:
                    var.set(valor_mayus)
                finally:
                    bandera["busy"] = False

        var.trace_add("write", convertir)

    def configurar_mayusculas_automaticas(self):
        vars_mayus = [
            self.sucursal_var,
            self.usuario_var,
            self.usuario_nombre_var,
            self.rol_var,
            self.venta_mesa_var,
            self.venta_tipo_var,
            self.venta_cliente_var,
            self.venta_direccion_var,
            self.venta_referencia_var,
            self.venta_producto_var,
            self.venta_metodo_pago_var,
            self.inv_nombre_var,
            self.inv_categoria_var,
            self.inv_tipo_var,
            self.inv_unidad_var,
            self.receta_producto_var,
            self.receta_insumo_var,
            self.receta_obs_var,
            self.caja_egreso_concepto_var,
            self.pedido_estado_var,
            self.cli_nombre_var,
            self.cli_direccion_var,
            self.cli_referencia_var,
            self.cli_notas_var,
            self.nueva_sucursal_var,
            self.nuevo_usuario_var,
            self.nuevo_nombre_usuario_var,
            self.nuevo_rol_usuario_var,
        ]
        for var in vars_mayus:
            self.forzar_mayusculas(var)

    def crear_estilos(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background=COLOR_SOFT, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 10), font=("Arial", 11, "bold"), background="#E5E7EB", foreground=COLOR_NAVY)
        style.map("TNotebook.Tab", background=[("selected", COLOR_GREEN)], foreground=[("selected", COLOR_NAVY_DARK)])
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background=COLOR_NAVY, foreground=COLOR_WHITE)
        style.configure("Treeview", rowheight=28, font=("Arial", 10), background=COLOR_WHITE, fieldbackground=COLOR_WHITE)
        style.configure("Aorix.TButton", font=("Arial", 10, "bold"), padding=(10, 6), background=COLOR_WHITE, foreground=COLOR_NAVY)
    def log(self, msg):
        linea = f"[{ahora().strftime('%H:%M:%S')}] {msg}"
        if hasattr(self, "txt_log"):
            self.txt_log.insert(tk.END, linea + "\n")
            self.txt_log.see(tk.END)
        print(linea)

    def set_status(self, msg):
        self.status_var.set(msg)
        self.log(msg)

    def fecha_trabajo(self):
        """Devuelve la fecha operativa elegida en el panel principal.
        Esta fecha manda para pedidos, caja, reportes diarios y respaldos.
        """
        valor = (self.fecha_var.get() or "").strip()
        try:
            datetime.strptime(valor, "%Y-%m-%d")
            return valor
        except Exception:
            hoy = fecha_hoy()
            self.fecha_var.set(hoy)
            return hoy

    def ahora_trabajo(self):
        """Combina la fecha operativa con la hora real actual."""
        h = ahora().strftime("%H:%M:%S")
        return datetime.strptime(f"{self.fecha_trabajo()} {h}", "%Y-%m-%d %H:%M:%S")

    def periodo_dir(self, raiz, fecha_txt=None):
        fecha_txt = fecha_txt or self.fecha_trabajo()
        dt = datetime.strptime(fecha_txt, "%Y-%m-%d")
        carpeta = os.path.join(raiz, str(dt.year), f"{dt.month:02d}_{MESES_ES.get(dt.month, str(dt.month))}")
        os.makedirs(carpeta, exist_ok=True)
        return carpeta

    def _rows_as_list(self, sql, params=()):
        filas = self.db.rows(sql, params)
        if not filas:
            return [], []
        cols = list(filas[0].keys())
        data = [[f[c] for c in cols] for f in filas]
        return cols, data

    def _guardar_excel_o_csv(self, ruta_xlsx, hojas):
        """Guarda respaldo en Excel; si openpyxl no está instalado, deja CSV por hoja."""
        try:
            if OPENPYXL_AVAILABLE:
                wb = openpyxl.Workbook()
                wb.remove(wb.active)
                for nombre_hoja, (cols, data) in hojas.items():
                    ws = wb.create_sheet(nombre_hoja[:31])
                    ws.append(cols or ["SIN_DATOS"])
                    for row in data:
                        ws.append(row)
                    for cell in ws[1]:
                        cell.font = openpyxl.styles.Font(bold=True)
                    ws.freeze_panes = "A2"
                    for col in ws.columns:
                        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
                        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 45)
                wb.save(ruta_xlsx)
            else:
                base = os.path.splitext(ruta_xlsx)[0]
                for nombre_hoja, (cols, data) in hojas.items():
                    ruta_csv = f"{base}_{nombre_hoja}.csv"
                    with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as f:
                        w = csv.writer(f)
                        w.writerow(cols or ["SIN_DATOS"])
                        w.writerows(data)
        except Exception as e:
            self.log(f"No se pudo guardar respaldo Excel: {e}")

    def exportar_respaldo_dia(self, fecha_txt=None):
        """Respaldo automático diario. No borra datos: solo actualiza archivos Excel."""
        fecha_txt = fecha_txt or self.fecha_trabajo()
        carpeta_sis = self.periodo_dir(self.dir_sistema_control, fecha_txt)
        carpeta_con = self.periodo_dir(self.dir_control_contable, fecha_txt)

        hojas_sistema = {
            "PEDIDOS": self._rows_as_list("SELECT * FROM pedidos WHERE fecha=? ORDER BY id", (fecha_txt,)),
            "DETALLE_PEDIDOS": self._rows_as_list("""
                SELECT d.* FROM pedido_detalle d
                INNER JOIN pedidos p ON p.id=d.pedido_id
                WHERE p.fecha=? ORDER BY d.id
            """, (fecha_txt,)),
            "INVENTARIO": self._rows_as_list("SELECT * FROM productos ORDER BY nombre"),
            "MOV_INVENTARIO": self._rows_as_list("SELECT * FROM movimientos_inventario WHERE substr(fecha_hora,1,10)=? ORDER BY id", (fecha_txt,)),
            "CIERRE_DIA": self._rows_as_list("SELECT * FROM cierres_dia WHERE fecha=? ORDER BY id", (fecha_txt,)),
        }
        hojas_contable = {
            "VENTAS_COBRADAS": self._rows_as_list("SELECT * FROM pedidos WHERE fecha=? AND pagado=1 ORDER BY id", (fecha_txt,)),
            "CAJA": self._rows_as_list("SELECT * FROM caja WHERE fecha=? ORDER BY id", (fecha_txt,)),
            "MOV_CAJA": self._rows_as_list("""
                SELECT mc.* FROM movimientos_caja mc
                LEFT JOIN caja c ON c.id=mc.caja_id
                WHERE COALESCE(c.fecha, substr(mc.fecha_hora,1,10))=?
                ORDER BY mc.id
            """, (fecha_txt,)),
            "CIERRE_CONTABLE": self._rows_as_list("SELECT * FROM cierres_dia WHERE fecha=? ORDER BY id", (fecha_txt,)),
        }
        self._guardar_excel_o_csv(os.path.join(carpeta_sis, f"Sistema_Control_Restaurante_{fecha_txt}.xlsx"), hojas_sistema)
        self._guardar_excel_o_csv(os.path.join(carpeta_con, f"Control_Contable_{fecha_txt}.xlsx"), hojas_contable)


    # =========================
    # CIERRE DE DÍA PROFESIONAL
    # =========================
    def dia_cerrado(self, fecha_txt=None):
        fecha_txt = fecha_txt or self.fecha_trabajo()
        return self.db.row("SELECT * FROM cierres_dia WHERE fecha=? AND estado='CERRADO'", (fecha_txt,)) is not None

    def validar_dia_abierto(self, accion="continuar"):
        fecha_txt = self.fecha_trabajo()
        if self.dia_cerrado(fecha_txt):
            messagebox.showwarning(
                "Día cerrado",
                f"El día {fecha_txt} ya está CERRADO.\n\n"
                f"No se puede {accion}.\n"
                "Puedes consultar reportes e indicadores, pero no modificar ventas, pedidos, caja ni stock de ese día."
            )
            return False
        return True

    def resumen_cierre_dia(self, fecha_txt=None):
        fecha_txt = fecha_txt or self.fecha_trabajo()
        return {
            "ventas": float(self.db.scalar("SELECT COALESCE(SUM(total),0) FROM pedidos WHERE fecha=? AND pagado=1", (fecha_txt,)) or 0),
            "pedidos": int(self.db.scalar("SELECT COUNT(*) FROM pedidos WHERE fecha=?", (fecha_txt,)) or 0),
            "pendientes": int(self.db.scalar("SELECT COUNT(*) FROM pedidos WHERE fecha=? AND pagado=0", (fecha_txt,)) or 0),
            "efectivo": float(self.db.scalar("SELECT COALESCE(SUM(efectivo_sistema),0) FROM caja WHERE fecha=?", (fecha_txt,)) or 0),
            "yape": float(self.db.scalar("SELECT COALESCE(SUM(yape),0) FROM caja WHERE fecha=?", (fecha_txt,)) or 0),
            "tarjeta": float(self.db.scalar("SELECT COALESCE(SUM(tarjeta),0) FROM caja WHERE fecha=?", (fecha_txt,)) or 0),
            "gastos": float(self.db.scalar("SELECT COALESCE(SUM(gastos),0) FROM caja WHERE fecha=?", (fecha_txt,)) or 0),
            "cajas_abiertas": int(self.db.scalar("SELECT COUNT(*) FROM caja WHERE fecha=? AND estado='ABIERTA'", (fecha_txt,)) or 0),
        }

    def cierre_dia_estado_texto(self):
        fecha_txt = self.fecha_trabajo()
        if self.dia_cerrado(fecha_txt):
            return f"🔒 DÍA CERRADO: {fecha_txt}"
        return f"🟢 DÍA ABIERTO: {fecha_txt}"

    def actualizar_estado_cierre_ui(self):
        if hasattr(self, "lbl_estado_cierre"):
            cerrado = self.dia_cerrado(self.fecha_trabajo())
            self.lbl_estado_cierre.config(
                text=self.cierre_dia_estado_texto(),
                fg="#B91C1C" if cerrado else "#065F46"
            )

    def ver_resumen_cierre_dia(self):
        fecha_txt = self.fecha_trabajo()
        r = self.resumen_cierre_dia(fecha_txt)
        estado = "CERRADO" if self.dia_cerrado(fecha_txt) else "ABIERTO"
        messagebox.showinfo(
            "Resumen del día",
            f"Fecha: {fecha_txt}\nEstado: {estado}\n\n"
            f"Ventas cobradas: {dinero(r['ventas'])}\n"
            f"Pedidos registrados: {r['pedidos']}\n"
            f"Pedidos pendientes/no pagados: {r['pendientes']}\n"
            f"Cajas abiertas: {r['cajas_abiertas']}\n\n"
            f"Efectivo: {dinero(r['efectivo'])}\n"
            f"Yape: {dinero(r['yape'])}\n"
            f"Tarjeta: {dinero(r['tarjeta'])}\n"
            f"Gastos: {dinero(r['gastos'])}"
        )

    def cerrar_dia_pro(self):
        fecha_txt = self.fecha_trabajo()
        if self.dia_cerrado(fecha_txt):
            messagebox.showinfo("Cierre de día", f"El día {fecha_txt} ya está cerrado.")
            return
        r = self.resumen_cierre_dia(fecha_txt)
        if r["pendientes"] > 0:
            messagebox.showwarning(
                "No se puede cerrar",
                f"Hay {r['pendientes']} pedido(s) pendiente(s) o no pagado(s) en {fecha_txt}.\n\n"
                "Primero cobra, entrega o corrige esos pedidos."
            )
            return
        if r["cajas_abiertas"] > 0:
            if not messagebox.askyesno(
                "Caja abierta",
                f"Hay {r['cajas_abiertas']} caja(s) abierta(s).\n\n"
                "¿Deseas cerrar la caja automáticamente y luego cerrar el día?"
            ):
                return
            while self.caja_abierta_actual():
                self.cerrar_caja()
            r = self.resumen_cierre_dia(fecha_txt)

        msg = (
            f"Vas a cerrar definitivamente el día {fecha_txt}.\n\n"
            f"Ventas cobradas: {dinero(r['ventas'])}\n"
            f"Pedidos: {r['pedidos']}\n"
            f"Efectivo: {dinero(r['efectivo'])}\n"
            f"Yape: {dinero(r['yape'])}\n"
            f"Tarjeta: {dinero(r['tarjeta'])}\n"
            f"Gastos: {dinero(r['gastos'])}\n\n"
            "Después del cierre, ese día quedará bloqueado para ventas, pedidos, caja e inventario.\n\n"
            "¿Confirmas el cierre?"
        )
        if not messagebox.askyesno("Confirmar cierre de día", msg):
            return
        self.exportar_respaldo_dia(fecha_txt)
        self.db.execute("""
            INSERT OR REPLACE INTO cierres_dia(
                fecha, sucursal, turno, usuario, hora_cierre, total_ventas, total_pedidos,
                total_pendientes, efectivo, yape, tarjeta, gastos, estado, observacion
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'CERRADO', ?)
        """, (
            fecha_txt,
            self.sucursal_var.get().strip(),
            self.turno_var.get().strip(),
            self.user["usuario"],
            self.ahora_trabajo().strftime("%H:%M:%S"),
            r["ventas"], r["pedidos"], r["pendientes"], r["efectivo"], r["yape"], r["tarjeta"], r["gastos"],
            "Cierre automático PRO con respaldo Excel."
        ), commit=True)
        self.exportar_respaldo_dia(fecha_txt)
        self.limpiar_venta()
        self.refrescar_todo()
        messagebox.showinfo("Cierre completo", f"Día {fecha_txt} cerrado correctamente.\n\nRespaldos Excel generados en las carpetas automáticas.")

    def reabrir_dia_pro(self):
        fecha_txt = self.fecha_trabajo()
        if not self.dia_cerrado(fecha_txt):
            messagebox.showinfo("Reabrir día", f"El día {fecha_txt} ya está abierto.")
            return
        if not self.pedir_clave_actual("Reabrir día cerrado"):
            return
        if not messagebox.askyesno("Reabrir día", f"¿Deseas REABRIR el día {fecha_txt}?\n\nSe volverán a permitir modificaciones."):
            return
        self.db.execute("UPDATE cierres_dia SET estado='REABIERTO', observacion=COALESCE(observacion,'') || ' | Reabierto por usuario.' WHERE fecha=?", (fecha_txt,), commit=True)
        self.exportar_respaldo_dia(fecha_txt)
        self.refrescar_todo()
        messagebox.showinfo("Reabrir día", f"Día {fecha_txt} reabierto correctamente.")


    def crear_tab_scrollable(self, parent):
        """Crea una pestaña con scroll vertical y ajuste automático al ancho."""
        outer = tk.Frame(parent, bg=COLOR_SOFT)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=COLOR_SOFT, highlightthickness=0, bd=0)
        yscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        xscroll = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)

        frame = tk.Frame(canvas, bg=COLOR_SOFT)
        window_id = canvas.create_window((0, 0), window=frame, anchor="nw")

        def actualizar_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            try:
                ancho_canvas = canvas.winfo_width()
                ancho_requerido = frame.winfo_reqwidth()
                canvas.itemconfigure(window_id, width=max(ancho_canvas, ancho_requerido))
            except Exception:
                pass

        frame.bind("<Configure>", actualizar_scroll)
        canvas.bind("<Configure>", actualizar_scroll)
        canvas.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)
        return frame

    def buscar_logo_aorix(self):
        """Detecta automáticamente el logo en la carpeta del archivo .py."""
        extensiones = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico")
        nombres_prioritarios = [
            "logo.png", "logo.jpg", "logo.jpeg", "logo.bmp", "logo.webp", "logo.ico",
            "aorix.png", "aorix.jpg", "aorix.jpeg", "aorix.bmp", "aorix.webp", "aorix.ico",
            "aorix_logo.png", "aorix_logo.jpg", "aorix_logo.jpeg", "aorix_logo.webp",
            "logo_aorix.png", "logo_aorix.jpg", "logo_aorix.jpeg", "logo_aorix.webp",
            "aorix_systems.png", "aorix_systems.jpg", "aorix_systems.jpeg", "aorix_systems.webp",
            "AORIX.png", "AORIX.jpg", "AORIX.jpeg", "AORIX.webp",
        ]
        carpetas = [self.base_dir]
        for sub in ("assets", "img", "images", "logo", "logos", "recursos"):
            ruta_sub = os.path.join(self.base_dir, sub)
            if os.path.isdir(ruta_sub):
                carpetas.append(ruta_sub)

        for carpeta in carpetas:
            for nombre in nombres_prioritarios:
                ruta = os.path.join(carpeta, nombre)
                if os.path.isfile(ruta):
                    return ruta

        candidatos = []
        try:
            for carpeta in carpetas:
                for nombre in os.listdir(carpeta):
                    bajo = nombre.lower()
                    ruta = os.path.join(carpeta, nombre)
                    if not os.path.isfile(ruta) or not bajo.endswith(extensiones):
                        continue
                    score = 0
                    if "aorix" in bajo:
                        score += 300
                    if "logo" in bajo:
                        score += 220
                    if "system" in bajo or "systems" in bajo:
                        score += 140
                    if "marca" in bajo or "brand" in bajo:
                        score += 80
                    if "screenshot" in bajo or "captura" in bajo or "image(" in bajo:
                        score -= 120
                    try:
                        size = os.path.getsize(ruta)
                    except Exception:
                        size = 0
                    candidatos.append((score, size, os.path.getmtime(ruta), ruta))
            if candidatos:
                candidatos.sort(reverse=True)
                return candidatos[0][3]
        except Exception:
            pass
        return None

    def cargar_logo_tk(self, max_w=210, max_h=135):
        """Carga el logo real con buena nitidez; si no existe, usa respaldo integrado."""
        ruta = self.buscar_logo_aorix()
        try:
            from PIL import Image, ImageTk, ImageOps
            if ruta:
                img = Image.open(ruta).convert("RGBA")
            else:
                img = Image.open(io.BytesIO(base64.b64decode(AORIX_FALLBACK_LOGO_B64))).convert("RGBA")
            img = ImageOps.contain(img, (max_w - 8, max_h - 8), Image.LANCZOS)
            lienzo = Image.new("RGBA", (max_w, max_h), COLOR_NAVY_DARK)
            x = (max_w - img.width) // 2
            y = (max_h - img.height) // 2
            lienzo.alpha_composite(img, (x, y))
            return ImageTk.PhotoImage(lienzo)
        except Exception:
            if ruta:
                try:
                    img = tk.PhotoImage(file=ruta)
                    while img.width() > max_w or img.height() > max_h:
                        img = img.subsample(2, 2)
                    return img
                except Exception:
                    pass
            try:
                raw = base64.b64decode(AORIX_FALLBACK_LOGO_B64)
                return tk.PhotoImage(data=base64.b64encode(raw).decode("ascii"))
            except Exception:
                return None

    def crear_ui(self):
        header = tk.Frame(self.root, bg=COLOR_NAVY, bd=0, relief="flat", highlightthickness=0)
        header.pack(fill="x", padx=0, pady=0)

        logo_box = tk.Frame(header, bg=COLOR_NAVY_DARK, width=230, height=150)
        logo_box.pack(side="left", padx=(24, 28), pady=18)
        logo_box.pack_propagate(False)

        self.logo_img = self.cargar_logo_tk(215, 135)
        if self.logo_img:
            tk.Label(logo_box, image=self.logo_img, bg=COLOR_NAVY_DARK).pack(expand=True)
        else:
            tk.Label(logo_box, text="AORIX", bg=COLOR_NAVY_DARK, fg=COLOR_WHITE,
                     font=("Arial", 26, "bold")).pack(expand=True)
            tk.Label(logo_box, text="SYSTEMS", bg=COLOR_NAVY_DARK, fg=COLOR_BLUE,
                     font=("Arial", 11, "bold")).pack()

        left = tk.Frame(header, bg=COLOR_NAVY)
        left.pack(side="left", fill="both", expand=True, pady=20)
        tk.Label(left, text=APP_TITLE, bg=COLOR_NAVY, fg=COLOR_WHITE,
                 font=("Arial", 30, "bold")).pack(anchor="w")
        tk.Label(left, text=APP_SUBTITLE, bg=COLOR_NAVY, fg="#D9ECFF",
                 font=("Arial", 14)).pack(anchor="w", pady=(14, 0))
        tk.Label(left, text=APP_BRAND, bg=COLOR_NAVY, fg=COLOR_GREEN,
                 font=("Arial", 13, "bold")).pack(anchor="w", pady=(14, 0))

        right = tk.Frame(header, bg=COLOR_NAVY)
        right.pack(side="right", padx=(10, 14), pady=18)
        btn_opts = dict(width=18, bg="#F8FAFC", fg=COLOR_NAVY, activebackground=COLOR_GREEN,
                        activeforeground=COLOR_NAVY_DARK, relief="raised", bd=2,
                        font=("Arial", 10, "bold"), cursor="hand2")
        tk.Button(right, text="Abrir base", command=self.abrir_bd, **btn_opts).pack(pady=5)
        tk.Button(right, text="Ver registro", command=self.abrir_tickets, **btn_opts).pack(pady=5)
        tk.Button(right, text="Config correo", command=self.exportar_ventas_csv, **btn_opts).pack(pady=5)

        nav = tk.Frame(self.root, bg=COLOR_NAVY_DARK)
        nav.pack(fill="x", padx=0, pady=(0, 8))
        nav_items = [
            ("Panel principal", 0),
            ("Ventas", 1),
            ("Pedidos", 2),
            ("Inventario", 3),
            ("Recetas", 4),
            ("Caja", 5),
            ("Delivery", 6),
            ("Indicadores", 7),
            ("Reportes", 8),
            ("Administrador", 9),
            ("Log", 10),
        ]
        for texto, idx in nav_items:
            tk.Button(nav, text=texto, width=16, relief="flat", bg=COLOR_NAVY_DARK, fg=COLOR_WHITE,
                      activebackground=COLOR_GREEN, activeforeground=COLOR_NAVY_DARK,
                      font=("Arial", 10, "bold"), cursor="hand2",
                      command=lambda i=idx: self.ir_tab(i)).pack(side="left", padx=3, pady=7)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        names = ["Panel Principal", "Ventas", "Pedidos", "Inventario", "Recetas", "Caja", "Delivery", "Indicadores", "Reportes", "Administrador", "Log"]
        self.tabs = []
        for name in names:
            wrap = tk.Frame(self.notebook, bg=COLOR_SOFT)
            self.notebook.add(wrap, text=name)
            self.tabs.append(self.crear_tab_scrollable(wrap))

        (self.tab_panel, self.tab_ventas, self.tab_pedidos, self.tab_inventario, self.tab_recetas,
         self.tab_caja, self.tab_delivery, self.tab_indicadores, self.tab_reportes, self.tab_admin, self.tab_log) = self.tabs

        self.crear_panel_principal()
        self.crear_tab_ventas()
        self.crear_tab_pedidos()
        self.crear_tab_inventario()
        self.crear_tab_recetas()
        self.crear_tab_caja()
        self.crear_tab_delivery()
        self.crear_tab_indicadores()
        self.crear_tab_reportes()
        self.crear_tab_admin()
        self.crear_tab_log()

        status = tk.Frame(self.root, bg=COLOR_NAVY)
        status.pack(fill="x", side="bottom")
        tk.Label(status, textvariable=self.status_var, bg=COLOR_NAVY, fg=COLOR_WHITE,
                 anchor="w", padx=12, pady=8, font=("Arial", 10, "bold")).pack(fill="x")

    def crear_campo_grid(self, parent, row, col, titulo, var, valores, state="readonly", width=None):
        box = tk.Frame(parent, bg="#F3F4F6")
        box.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
        tk.Label(box, text=titulo, bg="#F3F4F6", fg="#374151", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 4))
        cbo = ttk.Combobox(box, textvariable=var, values=valores, state=state, width=width)
        cbo.pack(fill="x", ipady=2)
        return cbo
    def crear_entry_grid(self, parent, row, col, titulo, var, width=None, state="normal"):
        box = tk.Frame(parent, bg="#F3F4F6")
        box.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
        tk.Label(box, text=titulo, bg="#F3F4F6", fg="#374151", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 4))
        ent = tk.Entry(box, textvariable=var, width=width, state=state)
        ent.pack(fill="x", ipady=2)
        return ent
    def crear_tarjeta(self, parent, col, titulo, valor, color):
        frame = tk.Frame(parent, bg="#FFFFFF", bd=1, relief="solid")
        frame.grid(row=0, column=col, padx=8, pady=6, sticky="nsew")
        parent.grid_columnconfigure(col, weight=1)
        tk.Label(frame, text=titulo, bg="#FFFFFF", fg="#374151", font=("Arial", 12, "bold")).pack(pady=(16, 8))
        lbl = tk.Label(frame, text=valor, bg="#FFFFFF", fg=color, font=("Arial", 22, "bold"))
        lbl.pack(pady=(0, 16))
        return lbl

    def ir_tab(self, idx):
        self.notebook.select(idx)

    def abrir_bd(self):
        ruta = os.path.join(self.base_dir, DB_NAME)
        try:
            if os.name == "nt":
                os.startfile(ruta)
            else:
                messagebox.showinfo("Base de datos", ruta)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def abrir_tickets(self):
        try:
            if os.name == "nt":
                os.startfile(self.tickets_dir)
            else:
                messagebox.showinfo("Tickets", self.tickets_dir)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def aplicar_permisos(self):
        rol = self.user["rol"].upper()
        if rol == "MESERO":
            self.notebook.tab(3, state="disabled")
            self.notebook.tab(4, state="disabled")
            self.notebook.tab(8, state="disabled")
        elif rol == "CAJA":
            self.notebook.tab(8, state="disabled")

    # Data helpers
    def categorias(self):
        return [r["nombre"] for r in self.db.rows("SELECT nombre FROM categorias ORDER BY nombre")]

    def tipos_producto(self):
        return ["VENTA", "INSUMO"]

    def unidades(self):
        return ["UND", "KG", "LITRO", "PORCION", "PLATO", "BOTELLA", "CAJA"]

    def productos_venta(self):
        return [r["nombre"] for r in self.db.rows("SELECT nombre FROM productos WHERE activo=1 AND tipo='VENTA' ORDER BY nombre")]

    def productos_insumo(self):
        return [r["nombre"] for r in self.db.rows("SELECT nombre FROM productos WHERE activo=1 AND tipo='INSUMO' ORDER BY nombre")]

    def clientes_nombres(self):
        return [r["nombre"] for r in self.db.rows("SELECT nombre FROM clientes ORDER BY nombre")]

    def mesas(self):
        nombres = [r["nombre"] for r in self.db.rows("SELECT nombre FROM mesas")]

        def clave(nombre):
            try:
                return int(str(nombre).upper().replace("MESA", "").strip())
            except Exception:
                return 999999

        return sorted(nombres, key=clave)

    def producto_por_nombre(self, nombre):
        return self.db.row("""
            SELECT p.*, c.nombre categoria
            FROM productos p
            LEFT JOIN categorias c ON c.id=p.categoria_id
            WHERE p.nombre=?
        """, (nombre,))

    def cliente_por_nombre(self, nombre):
        return self.db.row("SELECT * FROM clientes WHERE nombre=?", (nombre,))

    def mesa_id_por_nombre(self, nombre):
        r = self.db.row("SELECT id FROM mesas WHERE nombre=?", (nombre,))
        return r["id"] if r else None

    def receta_id_de_producto(self, producto_id):
        r = self.db.row("SELECT id FROM recetas WHERE producto_id=?", (producto_id,))
        return r["id"] if r else None

    def sucursales_disponibles(self):
        base = [self.get_config("sucursal", "Sucursal Principal"), "Sucursal Principal", "Sucursal 2", "Sucursal 3"]
        extras = [r["valor"] for r in self.db.rows("SELECT valor FROM configuracion WHERE clave LIKE 'sucursal_extra_%' ORDER BY valor")]
        vistos = []
        for item in base + extras:
            item = str(item).strip()
            if item and item not in vistos:
                vistos.append(item)
        return vistos

    def usuarios_activos_contexto(self):
        return self.db.rows("SELECT usuario, nombre, rol FROM usuarios WHERE activo=1 ORDER BY usuario")

    def nombres_contexto(self):
        return [r["nombre"] for r in self.usuarios_activos_contexto()]

    def roles_contexto(self):
        return [r["rol"] for r in self.usuarios_activos_contexto()]

    def sincronizar_contexto_usuario(self, event=None):
        usuario = self.usuario_var.get().strip()
        row = self.db.row("SELECT nombre, rol FROM usuarios WHERE usuario=? AND activo=1", (usuario,))
        if row:
            self.usuario_nombre_var.set(row["nombre"])
            self.rol_var.set(row["rol"])

    def crear_fecha_grid(self, parent, row, col, titulo):
        box = tk.Frame(parent, bg="#F3F4F6")
        box.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
        tk.Label(box, text=titulo, bg="#F3F4F6", fg="#374151", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 4))
        if TKCALENDAR_AVAILABLE:
            try:
                self.fecha_ctx_widget = DateEntry(
                    box,
                    textvariable=self.fecha_var,
                    date_pattern="yyyy-mm-dd",
                    locale="es_ES",
                    width=18
                )
                self.fecha_ctx_widget.pack(fill="x", ipady=1)
                return self.fecha_ctx_widget
            except Exception:
                pass
        self.fecha_ctx_widget = tk.Entry(box, textvariable=self.fecha_var)
        self.fecha_ctx_widget.pack(fill="x", ipady=2)
        return self.fecha_ctx_widget
    def crear_panel_principal(self):
        frm_ctx = tk.LabelFrame(self.tab_panel, text="Contexto de trabajo", bg="#F3F4F6", padx=12, pady=12)
        frm_ctx.pack(fill="x", padx=12, pady=8)
        for i in range(6):
            frm_ctx.grid_columnconfigure(i, weight=1)

        self.cbo_sucursal_ctx = self.crear_campo_grid(frm_ctx, 0, 0, "Sucursal", self.sucursal_var, self.sucursales_disponibles(), state="normal")
        self.crear_fecha_grid(frm_ctx, 0, 1, "Fecha")
        self.cbo_turno_ctx = self.crear_campo_grid(frm_ctx, 0, 2, "Turno", self.turno_var, ["MAÑANA", "TARDE", "NOCHE"])
        self.cbo_usuario_ctx = self.crear_campo_grid(frm_ctx, 0, 3, "Usuario", self.usuario_var, [r["usuario"] for r in self.usuarios_activos_contexto()])
        self.cbo_nombre_ctx = self.crear_campo_grid(frm_ctx, 0, 4, "Nombre", self.usuario_nombre_var, self.nombres_contexto())
        self.cbo_rol_ctx = self.crear_campo_grid(frm_ctx, 0, 5, "Rol", self.rol_var, self.roles_contexto())
        self.cbo_usuario_ctx.bind("<<ComboboxSelected>>", self.sincronizar_contexto_usuario)

        tk.Button(frm_ctx, text="Guardar contexto", width=16, command=self.guardar_contexto).grid(row=0, column=6, padx=8, pady=20)
        tk.Button(frm_ctx, text="Recargar panel", width=16, command=self.refrescar_todo).grid(row=0, column=7, padx=8, pady=20)

        cierre = tk.Frame(self.tab_panel, bg="#F3F4F6")
        cierre.pack(fill="x", padx=12, pady=(0, 8))
        self.lbl_estado_cierre = tk.Label(cierre, text=self.cierre_dia_estado_texto(), bg="#F3F4F6", fg="#065F46", font=("Arial", 12, "bold"))
        self.lbl_estado_cierre.pack(side="left", padx=8)
        tk.Button(cierre, text="🔒 Cerrar día", width=18, command=self.cerrar_dia_pro, bg="#FEE2E2", fg="#991B1B", font=("Arial", 10, "bold")).pack(side="left", padx=8)
        tk.Button(cierre, text="Ver resumen", width=14, command=self.ver_resumen_cierre_dia).pack(side="left", padx=8)
        tk.Button(cierre, text="Reabrir día", width=14, command=self.reabrir_dia_pro).pack(side="left", padx=8)

        ind = tk.LabelFrame(self.tab_panel, text="Indicadores", bg="#F3F4F6", padx=10, pady=10)
        ind.pack(fill="x", padx=12, pady=8)
        self.lbl_total_ventas = self.crear_tarjeta(ind, 0, "Ventas hoy", "S/ 0.00", "#111827")
        self.lbl_pedidos = self.crear_tarjeta(ind, 1, "Pedidos activos", "0", "#065F46")
        self.lbl_mesas = self.crear_tarjeta(ind, 2, "Mesas ocupadas", "0", "#B91C1C")
        self.lbl_stock = self.crear_tarjeta(ind, 3, "Stock bajo", "0", "#C2410C")
    def crear_tab_ventas(self):
        top = tk.LabelFrame(self.tab_ventas, text="Nueva venta / pedido", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        for i in range(8):
            top.grid_columnconfigure(i, weight=1)

        self.cbo_venta_mesa = self.crear_campo_grid(top, 0, 0, "Mesa", self.venta_mesa_var, self.mesas())
        self.cbo_venta_tipo = self.crear_campo_grid(top, 0, 1, "Tipo servicio", self.venta_tipo_var, ["SALÓN", "DELIVERY", "PARA LLEVAR"])
        self.crear_entry_grid(top, 0, 2, "Cliente", self.venta_cliente_var)
        self.crear_entry_grid(top, 0, 3, "Teléfono", self.venta_telefono_var)
        self.crear_entry_grid(top, 0, 4, "Dirección", self.venta_direccion_var)
        self.crear_entry_grid(top, 0, 5, "Referencia", self.venta_referencia_var)
        self.cbo_venta_producto = self.crear_campo_grid(top, 0, 6, "Producto", self.venta_producto_var, self.productos_venta(), state="normal")
        try:
            self.cbo_venta_producto.configure(height=25)
        except Exception:
            pass
        self.crear_entry_grid(top, 0, 7, "Cantidad", self.venta_cantidad_var)

        self.cbo_pago = self.crear_campo_grid(top, 1, 0, "Método pago", self.venta_metodo_pago_var, ["EFECTIVO", "YAPE", "TARJETA"])
        self.crear_entry_grid(top, 1, 1, "Descuento", self.venta_descuento_var)

        tk.Button(top, text="Buscar cliente", width=16, command=self.buscar_cliente_en_form).grid(row=1, column=2, padx=6, pady=10)
        tk.Button(top, text="Agregar ítem", width=16, command=self.agregar_item_actual).grid(row=1, column=3, padx=6, pady=10)
        tk.Button(top, text="Quitar ítem", width=16, command=self.quitar_item_actual).grid(row=1, column=4, padx=6, pady=10)
        tk.Button(top, text="Limpiar venta", width=16, command=self.limpiar_venta).grid(row=1, column=5, padx=6, pady=10)
        tk.Button(top, text="Guardar pedido", width=16, command=self.guardar_pedido).grid(row=1, column=6, padx=6, pady=10)
        tk.Button(top, text="Cobrar y ticket", width=16, command=self.cobrar_pedido_directo).grid(row=1, column=7, padx=6, pady=10)
        tk.Button(top, text="Carga inicio día", width=16, command=self.carga_masiva_inicio_dia).grid(row=2, column=4, padx=6, pady=(0, 10))
        tk.Button(top, text="Plantilla inicio día", width=16, command=self.descargar_plantilla_inicio_dia).grid(row=2, column=5, padx=6, pady=(0, 10))
        tk.Button(top, text="Nuevo prod. venta +", width=16, command=self.abrir_dialogo_producto_venta_desde_ventas).grid(row=2, column=6, padx=6, pady=(0, 10))
        tk.Button(top, text="Refrescar productos", width=16, command=self.refrescar_catalogo_ventas).grid(row=2, column=7, padx=6, pady=(0, 10))

        ayuda = tk.Label(
            top,
            text="Inicio del día: usa la plantilla Excel, carga stock inicial y el sistema crea o actualiza productos automáticamente.",
            bg="#F3F4F6",
            fg="#1F4E78",
            font=("Arial", 9, "bold"),
            anchor="w",
            justify="left"
        )
        ayuda.grid(row=3, column=0, columnspan=8, padx=8, pady=(0, 4), sticky="w")

        catalogo = tk.LabelFrame(self.tab_ventas, text="Catálogo dinámico de productos para venta", bg="#F3F4F6", padx=10, pady=10)
        catalogo.pack(fill="both", expand=False, padx=12, pady=8)
        barra = tk.Frame(catalogo, bg="#F3F4F6")
        barra.pack(fill="x", pady=(0, 8))
        tk.Label(barra, text="Buscar producto:", bg="#F3F4F6", font=("Arial", 10, "bold")).pack(side="left")
        ent_buscar = tk.Entry(barra, textvariable=self.venta_busqueda_producto_var)
        ent_buscar.pack(side="left", fill="x", expand=True, padx=8)
        ent_buscar.bind("<KeyRelease>", lambda e: self.filtrar_catalogo_ventas())
        tk.Button(barra, text="Limpiar búsqueda", width=16, command=self.limpiar_busqueda_catalogo_ventas).pack(side="left", padx=6)

        frame_catalogo = tk.Frame(catalogo, bg="#F3F4F6")
        frame_catalogo.pack(fill="both", expand=True)
        cols_catalogo = ("id", "producto", "categoria", "precio", "stock", "estado")
        self.tree_catalogo_ventas = ttk.Treeview(frame_catalogo, columns=cols_catalogo, show="headings", height=8)
        for c, t, w in [
            ("id", "ID", 60), ("producto", "Producto", 320), ("categoria", "Categoría", 130),
            ("precio", "Precio", 110), ("stock", "Stock", 90), ("estado", "Estado", 120),
        ]:
            self.tree_catalogo_ventas.heading(c, text=t)
            self.tree_catalogo_ventas.column(c, width=w, anchor="center")
        ys = ttk.Scrollbar(frame_catalogo, orient="vertical", command=self.tree_catalogo_ventas.yview)
        self.tree_catalogo_ventas.configure(yscrollcommand=ys.set)
        self.tree_catalogo_ventas.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree_catalogo_ventas.bind("<Double-1>", self.seleccionar_producto_catalogo_ventas)
        self.tree_catalogo_ventas.bind("<<TreeviewSelect>>", self.seleccionar_producto_catalogo_ventas)

        frame = tk.LabelFrame(self.tab_ventas, text="Detalle actual", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("producto", "cantidad", "precio", "subtotal")
        self.tree_venta = ttk.Treeview(frame, columns=cols, show="headings", height=14)
        for c, t, w in [("producto", "Producto", 360), ("cantidad", "Cantidad", 100), ("precio", "Precio", 120), ("subtotal", "Subtotal", 120)]:
            self.tree_venta.heading(c, text=t)
            self.tree_venta.column(c, width=w, anchor="center")
        self.tree_venta.pack(fill="both", expand=True)

        bottom = tk.Frame(self.tab_ventas, bg="#F3F4F6")
        bottom.pack(fill="x", padx=12, pady=8)
        self.lbl_subtotal = tk.Label(bottom, text="Subtotal: S/ 0.00", bg="#F3F4F6", font=("Arial", 12, "bold"))
        self.lbl_subtotal.pack(side="right", padx=8)
        self.lbl_total = tk.Label(bottom, text="Total: S/ 0.00", bg="#F3F4F6", fg="#065F46", font=("Arial", 16, "bold"))
        self.lbl_total.pack(side="right", padx=8)


    def productos_venta_rows(self, filtro=""):
        filtro = self.normalizar_texto(filtro)
        sql = """
            SELECT p.id, p.nombre, COALESCE(c.nombre,'-') AS categoria, p.precio, p.stock, p.stock_min
            FROM productos p
            LEFT JOIN categorias c ON c.id = p.categoria_id
            WHERE p.activo=1 AND p.tipo='VENTA'
        """
        params = []
        if filtro:
            sql += " AND UPPER(p.nombre) LIKE ?"
            params.append(f"%{filtro}%")
        sql += " ORDER BY p.nombre"
        return self.db.rows(sql, tuple(params))

    def refrescar_catalogo_ventas(self):
        self.actualizar_combos()
        self.set_status("Catálogo de ventas actualizado.")

    def filtrar_catalogo_ventas(self):
        if not hasattr(self, "tree_catalogo_ventas"):
            return
        for item in self.tree_catalogo_ventas.get_children():
            self.tree_catalogo_ventas.delete(item)
        rows = self.productos_venta_rows(self.venta_busqueda_producto_var.get())
        for r in rows:
            stock = float(r["stock"])
            minimo = float(r["stock_min"])
            if stock <= 0:
                estado = "SIN STOCK"
                tag = "agotado"
            elif stock <= minimo:
                estado = "STOCK BAJO"
                tag = "bajo"
            else:
                estado = "DISPONIBLE"
                tag = "ok"
            self.tree_catalogo_ventas.insert("", "end", values=(
                r["id"], r["nombre"], r["categoria"], dinero(r["precio"]), r["stock"], estado
            ), tags=(tag,))
        try:
            self.tree_catalogo_ventas.tag_configure("agotado", background="#FEE2E2", foreground="#991B1B")
            self.tree_catalogo_ventas.tag_configure("bajo", background="#FEF3C7", foreground="#92400E")
            self.tree_catalogo_ventas.tag_configure("ok", background="#ECFDF5", foreground="#065F46")
        except Exception:
            pass

    def limpiar_busqueda_catalogo_ventas(self):
        self.venta_busqueda_producto_var.set("")
        self.filtrar_catalogo_ventas()

    def seleccionar_producto_catalogo_ventas(self, event=None):
        if not hasattr(self, "tree_catalogo_ventas"):
            return
        sel = self.tree_catalogo_ventas.selection()
        if not sel:
            return
        vals = self.tree_catalogo_ventas.item(sel[0], "values")
        if not vals:
            return
        self.venta_producto_var.set(str(vals[1]))
        try:
            self.cbo_venta_producto.focus_set()
        except Exception:
            pass
        self.set_status(f"Producto seleccionado para venta: {vals[1]}")

    def abrir_dialogo_producto_venta_desde_ventas(self):
        self.abrir_dialogo_nuevo_producto(tipo_predefinido="VENTA", categoria_predefinida="PLATOS", seleccionar_en_venta=True)

    def crear_tab_pedidos(self):
        top = tk.LabelFrame(self.tab_pedidos, text="Control de pedidos", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        self.crear_campo_grid(top, 0, 0, "Estado", self.pedido_estado_var, ["TODOS", "PENDIENTE", "EN PREPARACIÓN", "LISTO", "ENTREGADO", "PAGADO", "ANULADO"])
        tk.Button(top, text="Refrescar", width=14, command=self.cargar_pedidos).grid(row=0, column=1, padx=6, pady=8)
        tk.Button(top, text="A preparación", width=16, command=lambda: self.cambiar_estado_pedido("EN PREPARACIÓN")).grid(row=0, column=2, padx=6, pady=8)
        tk.Button(top, text="A listo", width=16, command=lambda: self.cambiar_estado_pedido("LISTO")).grid(row=0, column=3, padx=6, pady=8)
        tk.Button(top, text="A entregado", width=16, command=lambda: self.cambiar_estado_pedido("ENTREGADO")).grid(row=0, column=4, padx=6, pady=8)
        tk.Button(top, text="Marcar pagado", width=16, command=lambda: self.cambiar_estado_pedido("PAGADO")).grid(row=0, column=5, padx=6, pady=8)
        tk.Button(top, text="Imprimir ticket", width=16, command=self.ticket_desde_pedido_seleccionado).grid(row=0, column=6, padx=6, pady=8)
        tk.Button(top, text="Quitar item", width=16, command=self.quitar_item_pedido_con_clave).grid(row=0, column=7, padx=6, pady=8)
        tk.Button(top, text="Limpiar pedido", width=16, command=self.limpiar_pedido_con_clave).grid(row=0, column=8, padx=6, pady=8)

        frame = tk.LabelFrame(self.tab_pedidos, text="Listado de pedidos", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("id", "codigo", "fecha", "hora", "mesa", "cliente", "tipo", "estado", "total", "pagado")
        self.tree_pedidos = ttk.Treeview(frame, columns=cols, show="headings", height=10, selectmode="extended")
        for c, t, w in [
            ("id", "ID", 60), ("codigo", "Código", 150), ("fecha", "Fecha", 110), ("hora", "Hora", 90),
            ("mesa", "Mesa", 110), ("cliente", "Cliente", 180), ("tipo", "Servicio", 120),
            ("estado", "Estado", 140), ("total", "Total", 100), ("pagado", "Pagado", 90),
        ]:
            self.tree_pedidos.heading(c, text=t)
            self.tree_pedidos.column(c, width=w, anchor="center")
        self.tree_pedidos.pack(fill="both", expand=True)
        self.tree_pedidos.bind("<<TreeviewSelect>>", self.seleccionar_pedido)

        frame_det = tk.LabelFrame(self.tab_pedidos, text="Detalle del pedido seleccionado", bg="#F3F4F6", padx=10, pady=10)
        frame_det.pack(fill="both", expand=True, padx=12, pady=8)
        cols_det = ("pedido_id", "codigo", "item", "producto", "cantidad", "precio", "subtotal")
        self.tree_pedido_detalle = ttk.Treeview(frame_det, columns=cols_det, show="headings", height=8)
        for c, t, w in [
            ("pedido_id", "Pedido ID", 90), ("codigo", "Código", 150), ("item", "Item", 70),
            ("producto", "Producto", 280), ("cantidad", "Cantidad", 110), ("precio", "Precio", 100), ("subtotal", "Subtotal", 120),
        ]:
            self.tree_pedido_detalle.heading(c, text=t)
            self.tree_pedido_detalle.column(c, width=w, anchor="center")
        self.tree_pedido_detalle.pack(fill="both", expand=True)
        self.tree_pedido_detalle.bind("<<TreeviewSelect>>", self.seleccionar_detalle_pedido)

    def crear_tab_inventario(self):
        top = tk.LabelFrame(self.tab_inventario, text="Productos e inventario", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        for i in range(8):
            top.grid_columnconfigure(i, weight=1)

        self.cbo_inv_nombre = self.crear_campo_grid(top, 0, 0, "Nombre producto", self.inv_nombre_var, self.productos_venta() + self.productos_insumo(), state="normal")
        self.cbo_inv_cat = self.crear_campo_grid(top, 0, 1, "Categoría", self.inv_categoria_var, self.categorias(), state="normal")
        self.cbo_inv_tipo = self.crear_campo_grid(top, 0, 2, "Tipo", self.inv_tipo_var, ["VENTA", "INSUMO"])
        self.cbo_inv_unidad = self.crear_campo_grid(top, 0, 3, "Unidad", self.inv_unidad_var, ["UND", "KG", "LITRO", "PORCION", "PLATO", "BOTELLA", "CAJA"], state="normal")
        self.crear_entry_grid(top, 0, 4, "Precio venta", self.inv_precio_var)
        self.crear_entry_grid(top, 0, 5, "Costo", self.inv_costo_var)
        self.crear_entry_grid(top, 0, 6, "Stock", self.inv_stock_var)
        self.crear_entry_grid(top, 0, 7, "Stock mínimo", self.inv_min_var)

        tk.Button(top, text="Nuevo producto +", width=16, command=self.abrir_dialogo_nuevo_producto).grid(row=1, column=0, padx=6, pady=10)
        tk.Button(top, text="Guardar producto", width=16, command=self.guardar_producto).grid(row=1, column=1, padx=6, pady=10)
        tk.Button(top, text="Entrada stock", width=16, command=lambda: self.movimiento_stock_manual("ENTRADA")).grid(row=1, column=2, padx=6, pady=10)
        tk.Button(top, text="Salida stock", width=16, command=lambda: self.movimiento_stock_manual("SALIDA")).grid(row=1, column=3, padx=6, pady=10)
        tk.Button(top, text="Llenar datos", width=16, command=self.cargar_producto_seleccionado_en_form).grid(row=1, column=4, padx=6, pady=10)
        tk.Button(top, text="Limpiar", width=16, command=self.limpiar_form_producto).grid(row=1, column=5, padx=6, pady=10)
        tk.Button(top, text="Alertas stock", width=16, command=self.mostrar_alertas_stock).grid(row=1, column=6, padx=6, pady=10)

        carga = tk.LabelFrame(self.tab_inventario, text="Carga masiva y utilidades", bg="#F3F4F6", padx=10, pady=10)
        carga.pack(fill="x", padx=12, pady=6)
        tk.Button(carga, text="Importar Excel", width=18, command=self.importar_productos_excel).pack(side="left", padx=6, pady=6)
        tk.Button(carga, text="Carga inicio día", width=18, command=self.carga_masiva_inicio_dia).pack(side="left", padx=6, pady=6)
        tk.Button(carga, text="Plantilla inicio día", width=18, command=self.descargar_plantilla_inicio_dia).pack(side="left", padx=6, pady=6)
        tk.Button(carga, text="Descargar plantilla", width=18, command=self.descargar_plantilla_productos).pack(side="left", padx=6, pady=6)
        tk.Button(carga, text="Exportar inventario", width=18, command=self.exportar_inventario_excel).pack(side="left", padx=6, pady=6)
        tk.Label(carga, textvariable=self.inv_alerta_var, bg="#F3F4F6", fg="#B45309", font=("Arial", 10, "bold")).pack(side="left", padx=18)

        frame = tk.LabelFrame(self.tab_inventario, text="Stock y productos", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("id", "nombre", "categoria", "tipo", "unidad", "precio", "costo", "stock", "minimo", "estado")
        self.tree_inv = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c, t, w in [
            ("id", "ID", 60), ("nombre", "Producto", 250), ("categoria", "Categoría", 120), ("tipo", "Tipo", 100),
            ("unidad", "Unidad", 100), ("precio", "Precio", 90), ("costo", "Costo", 90), ("stock", "Stock", 90),
            ("minimo", "Mínimo", 90), ("estado", "Estado", 110),
        ]:
            self.tree_inv.heading(c, text=t)
            self.tree_inv.column(c, width=w, anchor="center")
        self.tree_inv.pack(fill="both", expand=True)
        self.tree_inv.tag_configure("ok", foreground="#065F46")
        self.tree_inv.tag_configure("bajo", foreground="#B45309")
        self.tree_inv.tag_configure("agotado", foreground="#B91C1C")
    def crear_tab_recetas(self):
        top = tk.LabelFrame(self.tab_recetas, text="Recetas detalladas", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        for i in range(4):
            top.grid_columnconfigure(i, weight=1)
        self.cbo_receta_producto = self.crear_campo_grid(top, 0, 0, "Producto de venta", self.receta_producto_var, self.productos_venta())
        self.cbo_receta_insumo = self.crear_campo_grid(top, 0, 1, "Insumo", self.receta_insumo_var, self.productos_insumo())
        self.crear_entry_grid(top, 0, 2, "Cantidad del insumo", self.receta_cantidad_var)
        self.crear_entry_grid(top, 0, 3, "Observación", self.receta_obs_var)
        tk.Button(top, text="Guardar observación", width=18, command=self.guardar_observacion_receta).grid(row=1, column=0, padx=6, pady=10)
        tk.Button(top, text="Agregar insumo", width=16, command=self.agregar_insumo_a_receta).grid(row=1, column=1, padx=6, pady=10)
        tk.Button(top, text="Quitar insumo", width=16, command=self.quitar_insumo_receta).grid(row=1, column=2, padx=6, pady=10)
        tk.Button(top, text="Ver receta", width=16, command=self.cargar_receta_producto).grid(row=1, column=3, padx=6, pady=10)

        frame = tk.LabelFrame(self.tab_recetas, text="Detalle de receta", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("id", "insumo", "unidad", "cantidad")
        self.tree_receta = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c, t, w in [("id", "ID", 60), ("insumo", "Insumo", 260), ("unidad", "Unidad", 120), ("cantidad", "Cantidad", 120)]:
            self.tree_receta.heading(c, text=t)
            self.tree_receta.column(c, width=w, anchor="center")
        self.tree_receta.pack(fill="both", expand=True)

    def crear_tab_caja(self):
        top = tk.LabelFrame(self.tab_caja, text="Caja", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        tk.Label(top, text="Monto apertura:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=6, pady=8, sticky="w")
        tk.Entry(top, textvariable=self.caja_apertura_var, width=14).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        tk.Button(top, text="Abrir caja", width=16, command=self.abrir_caja).grid(row=0, column=2, padx=6, pady=8)
        tk.Button(top, text="Cerrar caja", width=16, command=self.cerrar_caja).grid(row=0, column=3, padx=6, pady=8)

        tk.Label(top, text="Egreso:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=1, column=0, padx=6, pady=8, sticky="w")
        tk.Entry(top, textvariable=self.caja_egreso_concepto_var, width=24).grid(row=1, column=1, padx=6, pady=8, sticky="w")
        tk.Entry(top, textvariable=self.caja_egreso_monto_var, width=14).grid(row=1, column=2, padx=6, pady=8, sticky="w")
        tk.Button(top, text="Registrar gasto", width=16, command=self.registrar_gasto).grid(row=1, column=3, padx=6, pady=8)

        frame = tk.LabelFrame(self.tab_caja, text="Resumen de caja", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="x", padx=12, pady=8)
        self.lbl_caja_estado = tk.Label(frame, text="Caja: SIN ABRIR", bg="#F3F4F6", font=("Arial", 14, "bold"))
        self.lbl_caja_estado.grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.lbl_caja_apertura = tk.Label(frame, text="Apertura: S/ 0.00", bg="#F3F4F6", font=("Arial", 12))
        self.lbl_caja_apertura.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.lbl_caja_efectivo = tk.Label(frame, text="Efectivo sistema: S/ 0.00", bg="#F3F4F6", font=("Arial", 12))
        self.lbl_caja_efectivo.grid(row=1, column=1, padx=6, pady=4, sticky="w")
        self.lbl_caja_yape = tk.Label(frame, text="Yape: S/ 0.00", bg="#F3F4F6", font=("Arial", 12))
        self.lbl_caja_yape.grid(row=2, column=0, padx=6, pady=4, sticky="w")
        self.lbl_caja_tarjeta = tk.Label(frame, text="Tarjeta: S/ 0.00", bg="#F3F4F6", font=("Arial", 12))
        self.lbl_caja_tarjeta.grid(row=2, column=1, padx=6, pady=4, sticky="w")
        self.lbl_caja_gastos = tk.Label(frame, text="Gastos: S/ 0.00", bg="#F3F4F6", font=("Arial", 12))
        self.lbl_caja_gastos.grid(row=3, column=0, padx=6, pady=4, sticky="w")

    def crear_tab_delivery(self):
        top = tk.LabelFrame(self.tab_delivery, text="Clientes delivery / frecuentes", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        for i in range(5):
            top.grid_columnconfigure(i, weight=1)

        self.crear_entry_grid(top, 0, 0, "Nombre", self.cli_nombre_var)
        self.crear_entry_grid(top, 0, 1, "Teléfono", self.cli_telefono_var)
        self.crear_entry_grid(top, 0, 2, "Dirección", self.cli_direccion_var)
        self.crear_entry_grid(top, 0, 3, "Referencia", self.cli_referencia_var)
        self.crear_entry_grid(top, 0, 4, "Notas", self.cli_notas_var)

        tk.Button(top, text="Guardar cliente", width=16, command=self.guardar_cliente).grid(row=1, column=0, padx=6, pady=10)
        tk.Button(top, text="Cargar en venta", width=16, command=self.cargar_cliente_en_venta).grid(row=1, column=1, padx=6, pady=10)
        tk.Button(top, text="Llenar datos", width=16, command=self.cargar_cliente_en_form).grid(row=1, column=2, padx=6, pady=10)
        tk.Button(top, text="Refrescar", width=16, command=self.cargar_clientes).grid(row=1, column=3, padx=6, pady=10)

        frame = tk.LabelFrame(self.tab_delivery, text="Base de clientes", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("id", "nombre", "telefono", "direccion", "referencia", "notas")
        self.tree_clientes = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c, t, w in [
            ("id", "ID", 60), ("nombre", "Nombre", 180), ("telefono", "Teléfono", 120),
            ("direccion", "Dirección", 240), ("referencia", "Referencia", 180), ("notas", "Notas", 200)
        ]:
            self.tree_clientes.heading(c, text=t)
            self.tree_clientes.column(c, width=w, anchor="center")
        self.tree_clientes.pack(fill="both", expand=True)


    def crear_tab_indicadores(self):
        top = tk.LabelFrame(self.tab_indicadores, text="Filtros de indicadores", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)

        tk.Label(top, text="Agrupar por:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=6, pady=8, sticky="w")
        self.cbo_ind_agrupacion = ttk.Combobox(
            top, textvariable=self.ind_agrupacion_var,
            values=["DÍA", "MES", "AÑO"], state="readonly", width=12
        )
        self.cbo_ind_agrupacion.grid(row=0, column=1, padx=6, pady=8, sticky="w")

        tk.Label(top, text="Fecha inicio:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=6, pady=8, sticky="w")
        tk.Entry(top, textvariable=self.ind_fecha_inicio_var, width=14).grid(row=0, column=3, padx=6, pady=8, sticky="w")

        tk.Label(top, text="Fecha fin:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=4, padx=6, pady=8, sticky="w")
        tk.Entry(top, textvariable=self.ind_fecha_fin_var, width=14).grid(row=0, column=5, padx=6, pady=8, sticky="w")

        tk.Label(top, text="Gráfica:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=6, padx=6, pady=8, sticky="w")
        self.cbo_ind_tipo = ttk.Combobox(
            top, textvariable=self.ind_tipo_grafica_var,
            values=["VENTAS S/", "PEDIDOS", "TICKET PROMEDIO"], state="readonly", width=18
        )
        self.cbo_ind_tipo.grid(row=0, column=7, padx=6, pady=8, sticky="w")

        tk.Button(top, text="Actualizar indicadores", width=22, command=self.actualizar_indicadores).grid(row=0, column=8, padx=6, pady=8)
        tk.Button(top, text="Hoy", width=10, command=self.ind_filtro_hoy).grid(row=0, column=9, padx=6, pady=8)

        resumen = tk.Frame(self.tab_indicadores, bg="#F3F4F6")
        resumen.pack(fill="x", padx=12, pady=(4, 8))
        self.ind_cards = {}
        cards = [
            ("ventas", "Ventas netas", "S/ 0.00"),
            ("pedidos", "Pedidos pagados", "0"),
            ("ticket", "Ticket promedio", "S/ 0.00"),
            ("pendientes", "Pedidos pendientes", "0"),
            ("stock", "Productos stock bajo", "0"),
        ]
        for key, titulo, valor in cards:
            card = tk.Frame(resumen, bg=COLOR_WHITE, bd=1, relief="solid", padx=12, pady=10)
            card.pack(side="left", fill="x", expand=True, padx=5)
            tk.Label(card, text=titulo, bg=COLOR_WHITE, fg="#4B5563", font=("Arial", 10, "bold")).pack(anchor="w")
            lbl = tk.Label(card, text=valor, bg=COLOR_WHITE, fg=COLOR_NAVY, font=("Arial", 18, "bold"))
            lbl.pack(anchor="w", pady=(4, 0))
            self.ind_cards[key] = lbl

        cuerpo = tk.Frame(self.tab_indicadores, bg="#F3F4F6")
        cuerpo.pack(fill="both", expand=True, padx=12, pady=8)

        graf_frame = tk.LabelFrame(cuerpo, text="Gráfica comparativa", bg="#F3F4F6", padx=10, pady=10)
        graf_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))
        self.ind_canvas = tk.Canvas(graf_frame, bg=COLOR_WHITE, height=420, highlightthickness=1, highlightbackground=COLOR_LINE)
        self.ind_canvas.pack(fill="both", expand=True)
        self.ind_canvas.bind("<Configure>", lambda e: self.dibujar_grafica_indicadores(getattr(self, "_ind_chart_data", [])))

        tabla_frame = tk.LabelFrame(cuerpo, text="Detalle por periodo", bg="#F3F4F6", padx=10, pady=10)
        tabla_frame.pack(side="right", fill="both", expand=True, padx=(8, 0))
        cols = ("periodo", "ventas", "pedidos", "ticket")
        self.tree_indicadores = ttk.Treeview(tabla_frame, columns=cols, show="headings", height=18)
        for col, txt, ancho in [
            ("periodo", "Periodo", 140),
            ("ventas", "Ventas S/", 130),
            ("pedidos", "Pedidos", 100),
            ("ticket", "Ticket promedio", 140),
        ]:
            self.tree_indicadores.heading(col, text=txt)
            self.tree_indicadores.column(col, width=ancho, anchor="center")
        self.tree_indicadores.pack(fill="both", expand=True)

        self.actualizar_indicadores()

    def ind_filtro_hoy(self):
        hoy = fecha_hoy()
        self.ind_fecha_inicio_var.set(hoy)
        self.ind_fecha_fin_var.set(hoy)
        self.ind_agrupacion_var.set("DÍA")
        self.actualizar_indicadores()

    def _indicador_periodo_sql(self):
        agrupacion = self.ind_agrupacion_var.get().strip().upper()
        if agrupacion == "MES":
            return "substr(fecha,1,7)", "Mes"
        if agrupacion == "AÑO":
            return "substr(fecha,1,4)", "Año"
        return "fecha", "Día"

    def actualizar_indicadores(self):
        fi = self.ind_fecha_inicio_var.get().strip()
        ff = self.ind_fecha_fin_var.get().strip()
        if not fi or not ff:
            messagebox.showwarning("Indicadores", "Ingrese fecha inicio y fecha fin en formato YYYY-MM-DD.")
            return
        if fi > ff:
            messagebox.showwarning("Indicadores", "La fecha inicio no puede ser mayor que la fecha fin.")
            return

        periodo_sql, _ = self._indicador_periodo_sql()
        estado_ok = ("PAGADO", "ENTREGADO")
        total = float(self.db.scalar(
            "SELECT COALESCE(SUM(total),0) FROM pedidos WHERE fecha BETWEEN ? AND ? AND estado IN (?,?)",
            (fi, ff, estado_ok[0], estado_ok[1])
        ) or 0)
        pedidos = int(self.db.scalar(
            "SELECT COUNT(*) FROM pedidos WHERE fecha BETWEEN ? AND ? AND estado IN (?,?)",
            (fi, ff, estado_ok[0], estado_ok[1])
        ) or 0)
        pendientes = int(self.db.scalar(
            "SELECT COUNT(*) FROM pedidos WHERE fecha BETWEEN ? AND ? AND estado IN ('PENDIENTE','EN PREPARACIÓN','LISTO')",
            (fi, ff)
        ) or 0)
        stock_bajo = int(self.db.scalar(
            "SELECT COUNT(*) FROM productos WHERE activo=1 AND stock <= stock_min"
        ) or 0)
        ticket = total / pedidos if pedidos else 0

        self.ind_cards["ventas"].config(text=dinero(total))
        self.ind_cards["pedidos"].config(text=str(pedidos))
        self.ind_cards["ticket"].config(text=dinero(ticket))
        self.ind_cards["pendientes"].config(text=str(pendientes))
        self.ind_cards["stock"].config(text=str(stock_bajo))

        rows = self.db.rows(f"""
            SELECT {periodo_sql} AS periodo,
                   COALESCE(SUM(total),0) AS ventas,
                   COUNT(*) AS pedidos,
                   CASE WHEN COUNT(*) > 0 THEN COALESCE(SUM(total),0) / COUNT(*) ELSE 0 END AS ticket
            FROM pedidos
            WHERE fecha BETWEEN ? AND ? AND estado IN (?,?)
            GROUP BY {periodo_sql}
            ORDER BY periodo
        """, (fi, ff, estado_ok[0], estado_ok[1]))

        for item in self.tree_indicadores.get_children():
            self.tree_indicadores.delete(item)
        datos = []
        for r in rows:
            periodo = str(r["periodo"])
            ventas = float(r["ventas"] or 0)
            cant = int(r["pedidos"] or 0)
            prom = float(r["ticket"] or 0)
            self.tree_indicadores.insert("", "end", values=(periodo, dinero(ventas), cant, dinero(prom)))
            tipo = self.ind_tipo_grafica_var.get().strip().upper()
            if tipo == "PEDIDOS":
                valor = cant
            elif tipo == "TICKET PROMEDIO":
                valor = prom
            else:
                valor = ventas
            datos.append((periodo, float(valor)))

        self._ind_chart_data = datos
        self.dibujar_grafica_indicadores(datos)
        self.set_status(f"Indicadores actualizados: {fi} a {ff}.")

    def dibujar_grafica_indicadores(self, datos):
        if not hasattr(self, "ind_canvas"):
            return
        canvas = self.ind_canvas
        canvas.delete("all")
        w = max(canvas.winfo_width(), 600)
        h = max(canvas.winfo_height(), 360)
        margen_izq, margen_der, margen_sup, margen_inf = 70, 25, 35, 70
        canvas.create_text(w // 2, 18, text=self.ind_tipo_grafica_var.get(), fill=COLOR_NAVY, font=("Arial", 14, "bold"))

        if not datos:
            canvas.create_text(w // 2, h // 2, text="Sin datos para el rango seleccionado", fill="#6B7280", font=("Arial", 14, "bold"))
            return

        max_val = max(v for _, v in datos) or 1
        plot_w = w - margen_izq - margen_der
        plot_h = h - margen_sup - margen_inf
        x0, y0 = margen_izq, h - margen_inf

        canvas.create_line(x0, margen_sup, x0, y0, fill="#94A3B8", width=2)
        canvas.create_line(x0, y0, w - margen_der, y0, fill="#94A3B8", width=2)

        pasos = 4
        for i in range(pasos + 1):
            val = max_val * i / pasos
            y = y0 - (plot_h * i / pasos)
            canvas.create_line(x0 - 5, y, w - margen_der, y, fill="#E5E7EB")
            etiqueta = dinero(val) if "S/" in self.ind_tipo_grafica_var.get() or "TICKET" in self.ind_tipo_grafica_var.get() else str(int(val))
            canvas.create_text(x0 - 10, y, text=etiqueta, anchor="e", fill="#475569", font=("Arial", 8))

        n = len(datos)
        gap = 10
        bar_w = max(14, min(60, (plot_w - gap * (n + 1)) / max(n, 1)))
        for i, (periodo, valor) in enumerate(datos):
            x1 = x0 + gap + i * (bar_w + gap)
            x2 = x1 + bar_w
            bar_h = 0 if max_val == 0 else (valor / max_val) * (plot_h - 10)
            y1 = y0 - bar_h
            canvas.create_rectangle(x1, y1, x2, y0, fill=COLOR_GREEN, outline=COLOR_NAVY)
            etiqueta_valor = dinero(valor) if "S/" in self.ind_tipo_grafica_var.get() or "TICKET" in self.ind_tipo_grafica_var.get() else str(int(valor))
            canvas.create_text((x1 + x2) / 2, y1 - 10, text=etiqueta_valor, fill=COLOR_NAVY, font=("Arial", 8, "bold"))
            txt = periodo if len(periodo) <= 10 else periodo[:10]
            canvas.create_text((x1 + x2) / 2, y0 + 16, text=txt, fill="#334155", font=("Arial", 8), angle=0)

    def crear_tab_reportes(self):
        top = tk.LabelFrame(self.tab_reportes, text="Filtros", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        tk.Label(top, text="Fecha inicio:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=6, pady=8)
        tk.Entry(top, textvariable=self.reporte_fecha_inicio_var, width=14).grid(row=0, column=1, padx=6, pady=8)
        tk.Label(top, text="Fecha fin:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=6, pady=8)
        tk.Entry(top, textvariable=self.reporte_fecha_fin_var, width=14).grid(row=0, column=3, padx=6, pady=8)
        tk.Button(top, text="Generar", width=16, command=self.generar_reporte).grid(row=0, column=4, padx=6, pady=8)
        tk.Button(top, text="Exportar CSV", width=16, command=self.exportar_ventas_csv).grid(row=0, column=5, padx=6, pady=8)

        frame = tk.LabelFrame(self.tab_reportes, text="Reporte", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.txt_reporte = tk.Text(frame, height=24, font=("Consolas", 10), wrap="word")
        self.txt_reporte.pack(fill="both", expand=True)

    def crear_tab_admin(self):
        top = tk.LabelFrame(self.tab_admin, text="Administrador", bg="#F3F4F6", padx=10, pady=10)
        top.pack(fill="x", padx=12, pady=8)
        tk.Label(top, text="Nombre negocio / sucursal:", bg="#F3F4F6", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=6, pady=8, sticky="w")
        self.cbo_sucursal_admin = ttk.Combobox(top, textvariable=self.sucursal_var, values=self.sucursales_disponibles(), state="normal", width=34)
        self.cbo_sucursal_admin.grid(row=0, column=1, padx=6, pady=8, sticky="w")
        tk.Button(top, text="Guardar negocio", width=18, command=self.guardar_contexto).grid(row=0, column=2, padx=6, pady=8)
        tk.Button(top, text="Sembrar demo", width=18, command=self.sembrar_demo).grid(row=0, column=3, padx=6, pady=8)
        tk.Button(top, text="Crear usuario demo", width=18, command=self.crear_usuario_demo).grid(row=0, column=4, padx=6, pady=8)

        suc = tk.LabelFrame(self.tab_admin, text="Registrar sucursal", bg="#F3F4F6", padx=10, pady=10)
        suc.pack(fill="x", padx=12, pady=8)
        tk.Label(suc, text="Nueva sucursal:", bg="#F3F4F6", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, pady=8, sticky="w")
        tk.Entry(suc, textvariable=self.nueva_sucursal_var, width=28).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        tk.Button(suc, text="Registrar sucursal", width=18, command=self.registrar_sucursal_admin).grid(row=0, column=2, padx=6, pady=8)

        usr = tk.LabelFrame(self.tab_admin, text="Registrar usuario", bg="#F3F4F6", padx=10, pady=10)
        usr.pack(fill="x", padx=12, pady=8)
        tk.Label(usr, text="Usuario:", bg="#F3F4F6", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, pady=8, sticky="w")
        tk.Entry(usr, textvariable=self.nuevo_usuario_var, width=18).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        tk.Label(usr, text="Nombre:", bg="#F3F4F6", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=6, pady=8, sticky="w")
        tk.Entry(usr, textvariable=self.nuevo_nombre_usuario_var, width=24).grid(row=0, column=3, padx=6, pady=8, sticky="w")
        tk.Label(usr, text="Clave:", bg="#F3F4F6", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=6, pady=8, sticky="w")
        tk.Entry(usr, textvariable=self.nueva_clave_usuario_var, width=16, show="*").grid(row=0, column=5, padx=6, pady=8, sticky="w")
        tk.Label(usr, text="Rol:", bg="#F3F4F6", font=("Arial", 10, "bold")).grid(row=0, column=6, padx=6, pady=8, sticky="w")
        ttk.Combobox(usr, textvariable=self.nuevo_rol_usuario_var, values=["ADMIN", "CAJA", "MESERO"], state="readonly", width=14).grid(row=0, column=7, padx=6, pady=8, sticky="w")
        tk.Button(usr, text="Guardar usuario", width=18, command=self.guardar_usuario_admin).grid(row=0, column=8, padx=6, pady=8)

        helpf = tk.LabelFrame(self.tab_admin, text="Resumen funcional", bg="#F3F4F6", padx=10, pady=10)
        helpf.pack(fill="both", expand=True, padx=12, pady=8)
        self.txt_admin = tk.Text(helpf, height=18, font=("Consolas", 10), wrap="word")
        self.txt_admin.pack(fill="both", expand=True)
        self.txt_admin.insert(
            "1.0",
            "VERSION 3 AJUSTADA\n\n"
            "MEJORAS ACTIVAS:\n"
            "- Sucursales registrables\n"
            "- Usuarios registrables\n"
            "- Pedido unificado para mismo cliente/mesa/servicio\n"
            "- Detalle de items en pestaña Pedidos\n"
            "- Quitar item y limpiar pedido con clave\n\n"
            f"Base actual: {os.path.join(self.base_dir, DB_NAME)}\n"
            f"Tickets: {self.tickets_dir}\n"
        )
        self.txt_admin.config(state="disabled")

    def crear_tab_log(self):
        frame = tk.LabelFrame(self.tab_log, text="Bitácora", bg="#F3F4F6", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.txt_log = tk.Text(frame, height=26, font=("Consolas", 10), wrap="word")
        self.txt_log.pack(fill="both", expand=True)

    # Refresh
    def actualizar_combos(self):
        if hasattr(self, "cbo_venta_mesa"):
            self.cbo_venta_mesa["values"] = self.mesas()
        if hasattr(self, "cbo_venta_producto"):
            self.cbo_venta_producto["values"] = self.productos_venta()
            try:
                self.cbo_venta_producto.configure(height=25)
            except Exception:
                pass
        if hasattr(self, "tree_catalogo_ventas"):
            self.filtrar_catalogo_ventas()
        if hasattr(self, "cbo_inv_cat"):
            self.cbo_inv_cat["values"] = self.categorias()
        if hasattr(self, "cbo_inv_nombre"):
            self.cbo_inv_nombre["values"] = self.productos_venta() + self.productos_insumo()
        if hasattr(self, "cbo_receta_producto"):
            self.cbo_receta_producto["values"] = self.productos_venta()
        if hasattr(self, "cbo_receta_insumo"):
            self.cbo_receta_insumo["values"] = self.productos_insumo()
        if hasattr(self, "cbo_sucursal_ctx"):
            self.cbo_sucursal_ctx["values"] = self.sucursales_disponibles()
        if hasattr(self, "cbo_sucursal_admin"):
            self.cbo_sucursal_admin["values"] = self.sucursales_disponibles()
        if hasattr(self, "cbo_usuario_ctx"):
            filas = self.usuarios_activos_contexto()
            self.cbo_usuario_ctx["values"] = [r["usuario"] for r in filas]
        if hasattr(self, "cbo_nombre_ctx"):
            self.cbo_nombre_ctx["values"] = self.nombres_contexto()
        if hasattr(self, "cbo_rol_ctx"):
            self.cbo_rol_ctx["values"] = self.roles_contexto()
    def refrescar_todo(self, inicial=False):
        self.actualizar_combos()
        self.cargar_indicadores()
        self.cargar_pedidos()
        self.cargar_inventario()
        self.filtrar_catalogo_ventas()
        self.cargar_caja_resumen()
        self.cargar_receta_producto()
        self.cargar_clientes()
        self.generar_reporte(silencioso=True)
        self.actualizar_estado_cierre_ui()
        if inicial:
            self.log("Sistema cargado correctamente.")
        self.set_status("Datos actualizados.")

    def guardar_contexto(self):
        self.set_config("sucursal", self.sucursal_var.get().strip() or "Sucursal Principal")
        self.sincronizar_contexto_usuario()
        self.actualizar_combos()
        self.set_status("Contexto guardado correctamente.")
    def cargar_indicadores(self):
        hoy = self.fecha_trabajo()
        ventas = self.db.scalar("SELECT COALESCE(SUM(total),0) FROM pedidos WHERE fecha=? AND estado IN ('PAGADO','ENTREGADO')", (hoy,))
        pedidos = self.db.scalar("SELECT COUNT(*) FROM pedidos WHERE estado IN ('PENDIENTE','EN PREPARACIÓN','LISTO')")
        mesas = self.db.scalar("SELECT COUNT(*) FROM mesas WHERE estado='OCUPADA'")
        stock = self.db.scalar("SELECT COUNT(*) FROM productos WHERE stock <= stock_min AND activo=1")
        self.lbl_total_ventas.config(text=dinero(ventas))
        self.lbl_pedidos.config(text=str(int(pedidos)))
        self.lbl_mesas.config(text=str(int(mesas)))
        self.lbl_stock.config(text=str(int(stock)))

    # Clientes / delivery
    def guardar_cliente(self):
        try:
            nombre = self.cli_nombre_var.get().strip()
            if not nombre:
                raise ValueError("Ingresa nombre del cliente.")
            tel = self.cli_telefono_var.get().strip()
            dirr = self.cli_direccion_var.get().strip()
            ref = self.cli_referencia_var.get().strip()
            notas = self.cli_notas_var.get().strip()

            ex = self.db.row("SELECT id FROM clientes WHERE nombre=?", (nombre,))
            if ex:
                self.db.execute("""
                    UPDATE clientes
                    SET telefono=?, direccion=?, referencia=?, notas=?
                    WHERE id=?
                """, (tel, dirr, ref, notas, ex["id"]), commit=True)
                self.set_status(f"Cliente actualizado: {nombre}")
            else:
                self.db.execute("""
                    INSERT INTO clientes(nombre, telefono, direccion, referencia, notas)
                    VALUES (?, ?, ?, ?, ?)
                """, (nombre, tel, dirr, ref, notas), commit=True)
                self.set_status(f"Cliente guardado: {nombre}")
            self.cargar_clientes()
        except Exception as e:
            messagebox.showerror("Clientes", str(e))

    def cargar_clientes(self):
        if not hasattr(self, "tree_clientes"):
            return
        for i in self.tree_clientes.get_children():
            self.tree_clientes.delete(i)
        for r in self.db.rows("SELECT * FROM clientes ORDER BY nombre"):
            self.tree_clientes.insert("", "end", values=(r["id"], r["nombre"], r["telefono"], r["direccion"], r["referencia"], r["notas"]))

    def cargar_cliente_en_form(self):
        sel = self.tree_clientes.selection()
        if not sel:
            messagebox.showwarning("Clientes", "Selecciona un cliente.")
            return
        vals = self.tree_clientes.item(sel[0], "values")
        self.cli_nombre_var.set(vals[1])
        self.cli_telefono_var.set(vals[2])
        self.cli_direccion_var.set(vals[3])
        self.cli_referencia_var.set(vals[4])
        self.cli_notas_var.set(vals[5])
        self.set_status(f"Cliente cargado en formulario: {vals[1]}")

    def cargar_cliente_en_venta(self):
        sel = self.tree_clientes.selection()
        if not sel:
            messagebox.showwarning("Clientes", "Selecciona un cliente.")
            return
        vals = self.tree_clientes.item(sel[0], "values")
        self.venta_cliente_var.set(vals[1])
        self.venta_telefono_var.set(vals[2])
        self.venta_direccion_var.set(vals[3])
        self.venta_referencia_var.set(vals[4])
        self.ir_tab(1)
        self.set_status(f"Cliente cargado en venta: {vals[1]}")

    def buscar_cliente_en_form(self):
        nombre = self.venta_cliente_var.get().strip()
        if not nombre:
            messagebox.showwarning("Venta", "Escribe el nombre del cliente.")
            return
        cli = self.cliente_por_nombre(nombre)
        if not cli:
            messagebox.showinfo("Venta", "Cliente no encontrado en la base.")
            return
        self.venta_telefono_var.set(cli["telefono"])
        self.venta_direccion_var.set(cli["direccion"])
        self.venta_referencia_var.set(cli["referencia"])
        self.set_status(f"Cliente encontrado: {cli['nombre']}")

    # Ventas

    def consolidar_items_con_nuevo(self, producto_id_nuevo=None, producto_nuevo=None, cantidad_nueva=0, precio_nuevo=0):
        """Consolida el detalle actual de venta/pedido para validar stock antes de guardar."""
        items = []
        for it in getattr(self, "current_order_items", []):
            items.append({
                "producto_id": it["producto_id"],
                "producto": it["producto"],
                "cantidad": float(it["cantidad"] or 0),
                "precio": float(it["precio"] or 0),
                "subtotal": float(it["subtotal"] or 0),
            })
        if producto_id_nuevo is not None and cantidad_nueva and float(cantidad_nueva) > 0:
            items.append({
                "producto_id": producto_id_nuevo,
                "producto": producto_nuevo or "",
                "cantidad": float(cantidad_nueva),
                "precio": float(precio_nuevo or 0),
                "subtotal": float(cantidad_nueva) * float(precio_nuevo or 0),
            })
        return items

    def validar_stock_para_items(self, items):
        """
        Valida stock antes de agregar/guardar/cobrar.
        Si el producto vendido tiene receta, se valida por insumos.
        Si no tiene receta, se valida por stock directo del producto.
        Agrupa cantidades repetidas para evitar sobreventa.
        """
        requerimientos = {}
        nombres = {}
        unidades = {}
        for it in items:
            producto_id = it.get("producto_id")
            cantidad_det = float(it.get("cantidad") or 0)
            if not producto_id or cantidad_det <= 0:
                continue
            prod = self.db.row("SELECT id, nombre, stock, unidad FROM productos WHERE id=?", (producto_id,))
            receta_id = self.receta_id_de_producto(producto_id)
            if receta_id:
                insumos = self.db.rows("""
                    SELECT p.id, p.nombre, p.stock, p.unidad, rd.cantidad
                    FROM receta_detalle rd
                    INNER JOIN productos p ON p.id=rd.insumo_id
                    WHERE rd.receta_id=?
                """, (receta_id,))
                if insumos:
                    for ins in insumos:
                        req = float(ins["cantidad"] or 0) * cantidad_det
                        requerimientos[ins["id"]] = requerimientos.get(ins["id"], 0) + req
                        nombres[ins["id"]] = ins["nombre"]
                        unidades[ins["id"]] = ins["unidad"]
                elif prod:
                    requerimientos[prod["id"]] = requerimientos.get(prod["id"], 0) + cantidad_det
                    nombres[prod["id"]] = prod["nombre"]
                    unidades[prod["id"]] = prod["unidad"]
            elif prod:
                requerimientos[prod["id"]] = requerimientos.get(prod["id"], 0) + cantidad_det
                nombres[prod["id"]] = prod["nombre"]
                unidades[prod["id"]] = prod["unidad"]
        faltantes = []
        advertencias = []
        for prod_id, requerido in requerimientos.items():
            row = self.db.row("SELECT stock, stock_min FROM productos WHERE id=?", (prod_id,))
            disponible = float(row["stock"] or 0) if row else 0.0
            minimo = float(row["stock_min"] or 0) if row else 0.0
            nombre = nombres.get(prod_id, str(prod_id))
            unidad = unidades.get(prod_id, "UND")
            if requerido > disponible:
                faltantes.append(f"{nombre} | requiere {requerido:g} {unidad} | stock {disponible:g}")
            elif (disponible - requerido) <= minimo:
                advertencias.append(f"{nombre} quedará en stock bajo: {disponible - requerido:g} {unidad}")
        return faltantes, advertencias

    def mostrar_faltantes_stock(self, titulo, faltantes, advertencias=None):
        lineas = []
        if faltantes:
            lineas.append("No hay stock suficiente para continuar:")
            lineas.append("")
            lineas.extend(faltantes[:20])
        if advertencias:
            if lineas:
                lineas.append("")
            lineas.append("Advertencias de stock bajo:")
            lineas.append("")
            lineas.extend(advertencias[:10])
        messagebox.showwarning(titulo, "\n".join(lineas) if lineas else "Revisa el stock.")

    def agregar_item_actual(self):
        try:
            if not self.validar_dia_abierto("agregar productos"):
                return
            producto = self.venta_producto_var.get().strip()
            cantidad = numero(self.venta_cantidad_var.get(), 0)
            if not producto:
                raise ValueError("Selecciona un producto.")
            if cantidad <= 0:
                raise ValueError("La cantidad debe ser mayor a cero.")
            prod = self.producto_por_nombre(producto)
            if not prod:
                if messagebox.askyesno("Venta", f"El producto '{producto}' no existe. ¿Deseas crearlo ahora como producto de venta?"):
                    self.inv_nombre_var.set(producto.upper())
                    self.inv_categoria_var.set("PLATOS")
                    self.inv_tipo_var.set("VENTA")
                    self.inv_unidad_var.set("UND")
                    self.inv_precio_var.set("0.00")
                    self.inv_costo_var.set("0.00")
                    self.inv_stock_var.set("0")
                    self.inv_min_var.set("0")
                    self.abrir_dialogo_producto_venta_desde_ventas()
                return
            items_proyectados = self.consolidar_items_con_nuevo(prod["id"], prod["nombre"], cantidad, float(prod["precio"]))
            faltantes, advertencias = self.validar_stock_para_items(items_proyectados)
            if faltantes:
                self.mostrar_faltantes_stock("Stock insuficiente", faltantes)
                self.set_status("Ítem bloqueado por stock insuficiente.")
                return
            if advertencias:
                if not messagebox.askyesno("Stock bajo", "El ítem se puede agregar, pero dejará productos en stock bajo.\n\n" + "\n".join(advertencias[:10]) + "\n\n¿Deseas continuar?"):
                    return
            subtotal = cantidad * float(prod["precio"])
            self.current_order_items.append({
                "producto_id": prod["id"],
                "producto": prod["nombre"],
                "cantidad": cantidad,
                "precio": float(prod["precio"]),
                "subtotal": subtotal
            })
            self.render_venta_actual()
            self.set_status(f"Ítem agregado: {producto} x {cantidad}")
        except Exception as e:
            messagebox.showwarning("Venta", str(e))

    def quitar_item_actual(self):
        sel = self.tree_venta.selection()
        if not sel:
            messagebox.showwarning("Venta", "Selecciona un ítem.")
            return
        idx = self.tree_venta.index(sel[0])
        if 0 <= idx < len(self.current_order_items):
            prod = self.current_order_items[idx]["producto"]
            del self.current_order_items[idx]
            self.render_venta_actual()
            self.set_status(f"Ítem quitado: {prod}")

    def limpiar_venta(self):
        self.current_order_items = []
        self.venta_cliente_var.set("")
        self.venta_telefono_var.set("")
        self.venta_direccion_var.set("")
        self.venta_referencia_var.set("")
        self.venta_cantidad_var.set("1")
        self.venta_descuento_var.set("0.00")
        self.render_venta_actual()
        self.set_status("Venta actual limpiada.")

    def render_venta_actual(self):
        for item in self.tree_venta.get_children():
            self.tree_venta.delete(item)
        for it in self.current_order_items:
            self.tree_venta.insert("", "end", values=(it["producto"], it["cantidad"], dinero(it["precio"]), dinero(it["subtotal"])))
        self.actualizar_totales_venta()

    def actualizar_totales_venta(self):
        subtotal = sum(x["subtotal"] for x in self.current_order_items)
        descuento = numero(self.venta_descuento_var.get(), 0)
        total = max(0, subtotal - descuento)
        self.lbl_subtotal.config(text=f"Subtotal: {dinero(subtotal)}")
        self.lbl_total.config(text=f"Total: {dinero(total)}")
        return subtotal, descuento, total

    def generar_codigo_pedido(self):
        return f"PED-{self.ahora_trabajo().strftime('%Y%m%d-%H%M%S')}"

    def guardar_cliente_si_no_existe_desde_venta(self):
        nombre = self.venta_cliente_var.get().strip()
        if not nombre:
            return None
        row = self.cliente_por_nombre(nombre)
        if row:
            return row["id"]
        self.db.execute("""
            INSERT INTO clientes(nombre, telefono, direccion, referencia, notas)
            VALUES (?, ?, ?, ?, '')
        """, (
            nombre,
            self.venta_telefono_var.get().strip(),
            self.venta_direccion_var.get().strip(),
            self.venta_referencia_var.get().strip(),
        ), commit=True)
        row = self.cliente_por_nombre(nombre)
        return row["id"] if row else None

    def guardar_pedido(self):
        try:
            if not self.validar_dia_abierto("guardar pedidos"):
                return None
            if not self.current_order_items:
                raise ValueError("Agrega al menos un producto.")
            faltantes, advertencias = self.validar_stock_para_items(self.consolidar_items_con_nuevo())
            if faltantes:
                self.mostrar_faltantes_stock("No se puede guardar", faltantes)
                return None
            if advertencias:
                if not messagebox.askyesno("Stock bajo", "El pedido se puede guardar, pero dejará productos en stock bajo.\n\n" + "\n".join(advertencias[:10]) + "\n\n¿Deseas guardar de todos modos?"):
                    return None
            mesa = self.normalizar_texto(self.venta_mesa_var.get())
            tipo = self.normalizar_texto(self.venta_tipo_var.get()) or "SALÓN"
            cliente = self.normalizar_texto(self.venta_cliente_var.get())
            if tipo == "SALÓN" and not mesa:
                raise ValueError("Selecciona una mesa para salón.")
            if not cliente:
                raise ValueError("Ingresa el cliente.")
            _, _, total_nuevo = self.actualizar_totales_venta()
            fh = self.ahora_trabajo()
            fecha_operativa = self.fecha_trabajo()
            mesa_id = self.mesa_id_por_nombre(mesa) if mesa else None
            cliente_id = self.guardar_cliente_si_no_existe_desde_venta()
            pedido_existente = self.buscar_pedido_abierto_relacionado(cliente, mesa_id, tipo)

            if pedido_existente:
                pedido_id = pedido_existente["id"]
                codigo = pedido_existente["codigo"]
                self.db.execute("""
                    UPDATE pedidos
                    SET telefono=?, direccion=?, referencia=?, usuario=?, hora=?
                    WHERE id=?
                """, (
                    self.normalizar_texto(self.venta_telefono_var.get()),
                    self.normalizar_texto(self.venta_direccion_var.get()),
                    self.normalizar_texto(self.venta_referencia_var.get()),
                    self.user["usuario"], fh.strftime("%H:%M:%S"), pedido_id
                ), commit=True)
            else:
                codigo = self.generar_codigo_pedido()
                cur = self.db.execute("""
                    INSERT INTO pedidos(
                        codigo, fecha, hora, mesa_id, cliente_id, cliente, telefono, direccion, referencia,
                        tipo_servicio, estado, total, pagado, metodo_pago, usuario, observacion
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDIENTE', ?, 0, '', ?, '')
                """, (
                    codigo, fecha_operativa, fh.strftime("%H:%M:%S"), mesa_id, cliente_id,
                    cliente, self.normalizar_texto(self.venta_telefono_var.get()), self.normalizar_texto(self.venta_direccion_var.get()),
                    self.normalizar_texto(self.venta_referencia_var.get()), tipo, total_nuevo, self.user["usuario"]
                ), commit=True)
                pedido_id = cur.lastrowid

            detalles = [(pedido_id, x["producto_id"], x["producto"], x["cantidad"], x["precio"], x["subtotal"]) for x in self.current_order_items]
            self.db.executemany("""
                INSERT INTO pedido_detalle(pedido_id, producto_id, producto, cantidad, precio_unit, subtotal)
                VALUES (?, ?, ?, ?, ?, ?)
            """, detalles, commit=True)
            total_actual = self.actualizar_total_pedido(pedido_id)

            if mesa_id:
                self.db.execute("UPDATE mesas SET estado='OCUPADA' WHERE id=?", (mesa_id,), commit=True)

            self.exportar_respaldo_dia(fecha_operativa)
            self.set_status(f"Pedido guardado: {codigo} | Fecha operativa: {fecha_operativa} | Total actual: {dinero(total_actual)}")
            self.limpiar_venta()
            self.refrescar_todo()
            return pedido_id
        except Exception as e:
            messagebox.showerror("Guardar pedido", str(e))
            return None

    def cobrar_pedido_directo(self):
        try:
            if not self.validar_dia_abierto("cobrar ventas"):
                return
            if not self.current_order_items:
                raise ValueError("No hay ítems para cobrar.")
            if not self.caja_abierta_actual():
                raise ValueError("Primero abre caja.")
            pedido_id = self.guardar_pedido()
            if pedido_id:
                self.procesar_pago_pedido(pedido_id, self.venta_metodo_pago_var.get())
                self.generar_ticket_txt(pedido_id)
                self.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Cobro", str(e))

    # Pedidos
    def cargar_pedidos(self):
        for i in self.tree_pedidos.get_children():
            self.tree_pedidos.delete(i)
        estado = self.pedido_estado_var.get().strip()
        sql = """
            SELECT p.id, p.codigo, p.fecha, p.hora, COALESCE(m.nombre,'-') mesa,
                   p.cliente, p.tipo_servicio, p.estado, p.total, p.pagado
            FROM pedidos p
            LEFT JOIN mesas m ON m.id=p.mesa_id
        """
        params = ()
        if estado and estado != "TODOS":
            sql += " WHERE p.estado=?"
            params = (estado,)
        sql += " ORDER BY p.pagado ASC, p.id DESC LIMIT 300"
        for r in self.db.rows(sql, params):
            self.tree_pedidos.insert("", "end", iid=str(r["id"]), values=(
                r["id"], r["codigo"], r["fecha"], r["hora"], r["mesa"], r["cliente"], r["tipo_servicio"],
                r["estado"], dinero(r["total"]), "SI" if int(r["pagado"]) else "NO"
            ))
        if hasattr(self, "tree_pedido_detalle"):
            for i in self.tree_pedido_detalle.get_children():
                self.tree_pedido_detalle.delete(i)
        self.selected_pedido_id = None
        self.selected_related_pedido_ids = []

    def seleccionar_pedido(self, event=None):
        if getattr(self, "_seleccionando_pedidos", False):
            return
        sel = self.tree_pedidos.selection()
        self.selected_pedido_id = None
        self.selected_related_pedido_ids = []
        if not sel:
            self.cargar_detalle_pedido_relacionado([])
            return
        vals = self.tree_pedidos.item(sel[0], "values")
        if not vals:
            self.cargar_detalle_pedido_relacionado([])
            return
        self.selected_pedido_id = int(vals[0])
        relacionados = self.pedido_ids_relacionados(self.selected_pedido_id, incluir_pagados=False)
        if not relacionados:
            relacionados = [self.selected_pedido_id]
        self.selected_related_pedido_ids = relacionados
        relacionados_iid = [str(x) for x in relacionados if self.tree_pedidos.exists(str(x))]
        actuales = list(self.tree_pedidos.selection())
        if relacionados_iid and set(actuales) != set(relacionados_iid):
            self._seleccionando_pedidos = True
            try:
                self.tree_pedidos.selection_set(relacionados_iid)
            finally:
                self._seleccionando_pedidos = False
        self.cargar_detalle_pedido_relacionado(relacionados)

    def cambiar_estado_pedido(self, estado):
        pedido_ids = list(self.selected_related_pedido_ids or ([self.selected_pedido_id] if self.selected_pedido_id else []))
        if not pedido_ids:
            messagebox.showwarning("Pedidos", "Selecciona un pedido.")
            return
        try:
            if estado == "PAGADO":
                metodo = self.venta_metodo_pago_var.get() or "EFECTIVO"
                for pid in pedido_ids:
                    self.procesar_pago_pedido(pid, metodo)
                self.generar_ticket_txt(pedido_ids[0], pedido_ids_relacionados=pedido_ids)
            else:
                for pid in pedido_ids:
                    self.db.execute("UPDATE pedidos SET estado=? WHERE id=?", (estado, pid), commit=True)
            self.set_status(f"Pedidos actualizados: {estado}")
            self.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Pedidos", str(e))


    def seleccionar_detalle_pedido(self, event=None):
        self.selected_detail_line = None
        if not hasattr(self, "tree_pedido_detalle"):
            return
        sel = self.tree_pedido_detalle.selection()
        if not sel:
            return
        vals = self.tree_pedido_detalle.item(sel[0], "values")
        if not vals:
            return
        self.selected_detail_line = {
            "pedido_id": int(vals[0]),
            "codigo": vals[1],
            "item": vals[2],
            "producto": vals[3],
            "cantidad": float(vals[4]),
            "precio": vals[5],
            "subtotal": vals[6],
        }

    def cargar_detalle_pedido_relacionado(self, pedido_ids):
        if not hasattr(self, "tree_pedido_detalle"):
            return
        for i in self.tree_pedido_detalle.get_children():
            self.tree_pedido_detalle.delete(i)
        self.selected_detail_line = None
        if not pedido_ids:
            return
        placeholders = ",".join(["?"] * len(pedido_ids))
        rows = self.db.rows(f"""
            SELECT d.id, d.pedido_id, p.codigo, d.producto, d.cantidad, d.precio_unit, d.subtotal
            FROM pedido_detalle d
            INNER JOIN pedidos p ON p.id = d.pedido_id
            WHERE d.pedido_id IN ({placeholders})
            ORDER BY d.pedido_id, d.id
        """, tuple(pedido_ids))
        item_num = 1
        for r in rows:
            self.tree_pedido_detalle.insert(
                "", "end", iid=f"DET_{r['id']}",
                values=(r["pedido_id"], r["codigo"], item_num, r["producto"], float(r["cantidad"]), dinero(r["precio_unit"]), dinero(r["subtotal"]))
            )
            item_num += 1

    def _normalizar_mesas_despues_de_cambios(self, mesa_id):
        if not mesa_id:
            return
        pendientes = self.db.scalar(
            "SELECT COUNT(*) FROM pedidos WHERE COALESCE(mesa_id,0)=COALESCE(?,0) AND pagado=0 AND estado NOT IN ('ANULADO','PAGADO')",
            (mesa_id,)
        )
        nuevo_estado = 'OCUPADA' if int(pendientes or 0) > 0 else 'LIBRE'
        self.db.execute("UPDATE mesas SET estado=? WHERE id=?", (nuevo_estado, mesa_id), commit=True)

    def _eliminar_pedido_si_vacio(self, pedido_id):
        pedido = self.db.row("SELECT * FROM pedidos WHERE id=?", (pedido_id,))
        if not pedido:
            return
        total = self.actualizar_total_pedido(pedido_id)
        cant = self.db.scalar("SELECT COUNT(*) FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
        if int(cant or 0) == 0 or float(total or 0) <= 0:
            mesa_id = pedido["mesa_id"]
            self.db.execute("DELETE FROM pedidos WHERE id=?", (pedido_id,), commit=True)
            self._normalizar_mesas_despues_de_cambios(mesa_id)
        else:
            self._normalizar_mesas_despues_de_cambios(pedido["mesa_id"])

    def quitar_item_pedido_con_clave(self):
        if not self.selected_pedido_id:
            messagebox.showwarning("Pedidos", "Selecciona un pedido.")
            return
        if not getattr(self, "selected_detail_line", None):
            messagebox.showwarning("Pedidos", "Selecciona un item del detalle del pedido.")
            return
        if not self.pedir_clave_actual("Seguridad - Quitar item"):
            return
        try:
            pedido_id = int(self.selected_detail_line["pedido_id"])
            producto = self.selected_detail_line["producto"]
            detalle = self.db.row(
                "SELECT id FROM pedido_detalle WHERE pedido_id=? AND producto=? ORDER BY id DESC LIMIT 1",
                (pedido_id, producto)
            )
            if not detalle:
                raise ValueError("No se encontró el item seleccionado.")
            self.db.execute("DELETE FROM pedido_detalle WHERE id=?", (detalle["id"],), commit=True)
            self._eliminar_pedido_si_vacio(pedido_id)
            self.set_status(f"Item quitado del pedido: {producto}")
            self.refrescar_todo()
            try:
                if self.tree_pedidos.exists(str(pedido_id)):
                    self.tree_pedidos.selection_set(str(pedido_id))
                    self.seleccionar_pedido()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Pedidos", str(e))

    def limpiar_pedido_con_clave(self):
        pedido_ids = list(self.selected_related_pedido_ids or ([self.selected_pedido_id] if self.selected_pedido_id else []))
        if not pedido_ids:
            messagebox.showwarning("Pedidos", "Selecciona un pedido.")
            return
        if not self.pedir_clave_actual("Seguridad - Limpiar pedido"):
            return
        try:
            if not messagebox.askyesno("Pedidos", "Se eliminarán todos los items de los pedidos seleccionados. ¿Deseas continuar?"):
                return
            for pid in pedido_ids:
                pedido = self.db.row("SELECT * FROM pedidos WHERE id=?", (pid,))
                if not pedido:
                    continue
                self.db.execute("DELETE FROM pedido_detalle WHERE pedido_id=?", (pid,), commit=True)
                self.db.execute("DELETE FROM pedidos WHERE id=?", (pid,), commit=True)
                self._normalizar_mesas_despues_de_cambios(pedido["mesa_id"])
            self.selected_pedido_id = None
            self.selected_related_pedido_ids = []
            self.selected_detail_line = None
            self.set_status("Pedido limpiado correctamente.")
            self.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Pedidos", str(e))

    def guardar_producto(self):
        try:
            nombre = self.inv_nombre_var.get().strip().upper()
            categoria = self.inv_categoria_var.get().strip().upper()
            tipo = self.inv_tipo_var.get().strip().upper()
            unidad = self.inv_unidad_var.get().strip().upper()
            precio = numero(self.inv_precio_var.get(), 0)
            costo = numero(self.inv_costo_var.get(), 0)
            stock = numero(self.inv_stock_var.get(), 0)
            stock_min = numero(self.inv_min_var.get(), 0)

            if not nombre:
                raise ValueError("Ingresa nombre.")
            if not categoria:
                raise ValueError("Selecciona categoría.")
            if tipo not in ("VENTA", "INSUMO"):
                raise ValueError("El tipo debe ser VENTA o INSUMO.")
            if not unidad:
                raise ValueError("Ingresa una unidad.")
            if precio < 0 or costo < 0 or stock < 0 or stock_min < 0:
                raise ValueError("No se permiten valores negativos.")

            cat = self.db.row("SELECT id FROM categorias WHERE nombre=?", (categoria,))
            if not cat:
                self.db.execute("INSERT INTO categorias(nombre) VALUES (?)", (categoria,), commit=True)
                cat = self.db.row("SELECT id FROM categorias WHERE nombre=?", (categoria,))
            ex = self.db.row("SELECT id FROM productos WHERE nombre=?", (nombre,))
            if ex:
                self.db.execute("""
                    UPDATE productos
                    SET categoria_id=?, tipo=?, unidad=?, precio=?, costo=?, stock=?, stock_min=?, activo=1
                    WHERE id=?
                """, (cat["id"], tipo, unidad, precio, costo, stock, stock_min, ex["id"]), commit=True)
                self.set_status(f"Producto actualizado: {nombre}")
            else:
                self.db.execute("""
                    INSERT INTO productos(nombre, categoria_id, tipo, unidad, precio, costo, stock, stock_min, activo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                """, (nombre, cat["id"], tipo, unidad, precio, costo, stock, stock_min), commit=True)
                self.set_status(f"Producto creado: {nombre}")
            self.actualizar_combos()
            self.cargar_inventario()
        except Exception as e:
            messagebox.showerror("Inventario", str(e))

    def cargar_inventario(self):
        for i in self.tree_inv.get_children():
            self.tree_inv.delete(i)
        rows = self.db.rows("""
            SELECT p.id, p.nombre, c.nombre categoria, p.tipo, p.unidad, p.precio, p.costo, p.stock, p.stock_min
            FROM productos p
            LEFT JOIN categorias c ON c.id=p.categoria_id
            ORDER BY p.nombre
        """)
        bajos = 0
        agotados = 0
        for r in rows:
            stock = float(r["stock"])
            minimo = float(r["stock_min"])
            if stock <= 0:
                estado = "SIN STOCK"
                tag = "agotado"
                agotados += 1
            elif stock <= minimo:
                estado = "STOCK BAJO"
                tag = "bajo"
                bajos += 1
            else:
                estado = "OK"
                tag = "ok"
            self.tree_inv.insert("", "end", values=(
                r["id"], r["nombre"], r["categoria"] or "-", r["tipo"], r["unidad"], dinero(r["precio"]), dinero(r["costo"]), r["stock"], r["stock_min"], estado
            ), tags=(tag,))
        resumen = f"Alertas: {agotados} sin stock | {bajos} con stock bajo"
        self.inv_alerta_var.set(resumen if (agotados or bajos) else "Sin alertas de stock.")
        self.cargar_indicadores()

    def limpiar_form_producto(self):
        self.inv_nombre_var.set("")
        self.inv_categoria_var.set("")
        self.inv_tipo_var.set("VENTA")
        self.inv_unidad_var.set("UND")
        self.inv_precio_var.set("0.00")
        self.inv_costo_var.set("0.00")
        self.inv_stock_var.set("0")
        self.inv_min_var.set("0")
        self.set_status("Formulario de producto limpio.")

    def abrir_dialogo_nuevo_producto(self, tipo_predefinido="VENTA", categoria_predefinida="PLATOS", seleccionar_en_venta=False):
        win = tk.Toplevel(self.root)
        win.title("Nuevo producto")
        win.geometry("430x430")
        win.transient(self.root)
        win.grab_set()
        win.configure(bg="#F3F4F6")

        nombre_var = tk.StringVar()
        categoria_var = tk.StringVar(value=categoria_predefinida)
        tipo_var = tk.StringVar(value=tipo_predefinido)
        unidad_var = tk.StringVar(value="UND")
        precio_var = tk.StringVar(value="0.00")
        costo_var = tk.StringVar(value="0.00")
        stock_var = tk.StringVar(value="0")
        minimo_var = tk.StringVar(value="0")

        campos = [
            ("Nombre del producto", nombre_var),
            ("Precio de venta", precio_var),
            ("Costo", costo_var),
            ("Stock inicial", stock_var),
            ("Stock mínimo", minimo_var),
        ]
        for i, (lbl, var) in enumerate(campos):
            box = tk.Frame(win, bg="#F3F4F6")
            box.pack(fill="x", padx=16, pady=6)
            tk.Label(box, text=lbl, bg="#F3F4F6", font=("Arial", 10, "bold")).pack(anchor="w")
            tk.Entry(box, textvariable=var).pack(fill="x", pady=(4, 0))

        box1 = tk.Frame(win, bg="#F3F4F6")
        box1.pack(fill="x", padx=16, pady=6)
        tk.Label(box1, text="Categoría", bg="#F3F4F6", font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Combobox(box1, textvariable=categoria_var, values=self.categorias(), state="normal").pack(fill="x", pady=(4, 0))

        box2 = tk.Frame(win, bg="#F3F4F6")
        box2.pack(fill="x", padx=16, pady=6)
        tk.Label(box2, text="Tipo", bg="#F3F4F6", font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Combobox(box2, textvariable=tipo_var, values=["VENTA", "INSUMO"], state="readonly").pack(fill="x", pady=(4, 0))

        box3 = tk.Frame(win, bg="#F3F4F6")
        box3.pack(fill="x", padx=16, pady=6)
        tk.Label(box3, text="Unidad", bg="#F3F4F6", font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Combobox(box3, textvariable=unidad_var, values=self.unidades(), state="normal").pack(fill="x", pady=(4, 0))

        def guardar_nuevo():
            self.inv_nombre_var.set(nombre_var.get().strip().upper())
            self.inv_categoria_var.set(categoria_var.get().strip().upper())
            self.inv_tipo_var.set(tipo_var.get().strip().upper())
            self.inv_unidad_var.set(unidad_var.get().strip().upper())
            self.inv_precio_var.set(precio_var.get().strip())
            self.inv_costo_var.set(costo_var.get().strip())
            self.inv_stock_var.set(stock_var.get().strip())
            self.inv_min_var.set(minimo_var.get().strip())
            self.guardar_producto()
            if seleccionar_en_venta and self.inv_tipo_var.get().strip().upper() == "VENTA":
                self.venta_producto_var.set(self.inv_nombre_var.get().strip().upper())
                self.venta_busqueda_producto_var.set("")
                self.filtrar_catalogo_ventas()
            win.destroy()

        botones = tk.Frame(win, bg="#F3F4F6")
        botones.pack(fill="x", padx=16, pady=16)
        tk.Button(botones, text="Guardar", width=16, command=guardar_nuevo).pack(side="left", padx=4)
        tk.Button(botones, text="Cancelar", width=16, command=win.destroy).pack(side="left", padx=4)

    def productos_stock_bajo(self):
        return self.db.rows("""
            SELECT p.nombre, p.stock, p.stock_min, p.unidad
            FROM productos p
            WHERE p.activo=1 AND p.stock <= p.stock_min
            ORDER BY p.stock ASC, p.nombre ASC
        """)

    def mostrar_alertas_stock(self):
        bajos = self.productos_stock_bajo()
        if not bajos:
            messagebox.showinfo("Stock", "No hay productos con stock bajo.")
            return
        lineas = ["Productos con alerta de stock:", ""]
        for r in bajos[:40]:
            lineas.append(f"- {r['nombre']}: {r['stock']} {r['unidad']} | mínimo {r['stock_min']}")
        if len(bajos) > 40:
            lineas.append("")
            lineas.append(f"... y {len(bajos)-40} más")
        messagebox.showwarning("Alerta de stock mínimo", "\n".join(lineas))


    def mostrar_alerta_stock_en_pantalla(self):
        bajos = self.productos_stock_bajo()
        if not bajos:
            self.inv_alerta_var.set("Sin alertas de stock.")
            return
        agotados = sum(1 for r in bajos if float(r["stock"]) <= 0)
        bajos_no_agotados = sum(1 for r in bajos if float(r["stock"]) > 0)
        self.inv_alerta_var.set(f"Alertas: {agotados} sin stock | {bajos_no_agotados} con stock bajo")



    def _obtener_o_crear_categoria(self, categoria):
        categoria = self.normalizar_texto(categoria) or "GENERAL"
        cat = self.db.row("SELECT id FROM categorias WHERE nombre=?", (categoria,))
        if not cat:
            self.db.execute("INSERT INTO categorias(nombre) VALUES (?)", (categoria,), commit=True)
            cat = self.db.row("SELECT id FROM categorias WHERE nombre=?", (categoria,))
        return cat["id"], categoria

    def _upsert_producto_excel(self, nombre, categoria, tipo, unidad, precio, costo, stock, stock_min, referencia_movimiento=None, stock_es_absoluto=True):
        nombre = self.normalizar_texto(nombre)
        categoria = self.normalizar_texto(categoria) or ("PLATOS" if self.normalizar_texto(tipo) == "VENTA" else "INSUMOS")
        tipo = self.normalizar_texto(tipo) or "VENTA"
        unidad = self.normalizar_texto(unidad) or "UND"
        if not nombre:
            return "omitido", None

        categoria_id, categoria = self._obtener_o_crear_categoria(categoria)
        precio = numero(precio, 0)
        costo = numero(costo, 0)
        stock = numero(stock, 0)
        stock_min = numero(stock_min, 0)

        ex = self.db.row("SELECT * FROM productos WHERE nombre=?", (nombre,))
        if ex:
            stock_anterior = float(ex["stock"])
            nuevo_stock = stock if stock_es_absoluto else stock_anterior + stock
            self.db.execute("""
                UPDATE productos
                SET categoria_id=?, tipo=?, unidad=?, precio=?, costo=?, stock=?, stock_min=?, activo=1
                WHERE id=?
            """, (categoria_id, tipo, unidad, precio, costo, nuevo_stock, stock_min, ex["id"]), commit=True)
            prod_id = ex["id"]
            accion = "actualizado"
        else:
            nuevo_stock = stock
            prod_id = self.db.execute("""
                INSERT INTO productos(nombre, categoria_id, tipo, unidad, precio, costo, stock, stock_min, activo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
            """, (nombre, categoria_id, tipo, unidad, precio, costo, nuevo_stock, stock_min), commit=True).lastrowid
            stock_anterior = 0.0
            accion = "insertado"

        if referencia_movimiento is not None:
            diferencia = float(nuevo_stock) - float(stock_anterior)
            self.db.execute("""
                INSERT INTO movimientos_inventario(fecha_hora, producto_id, producto, tipo, cantidad, referencia, costo_unit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                ahora().strftime("%Y-%m-%d %H:%M:%S"),
                prod_id,
                nombre,
                "CARGA_INICIAL" if stock_es_absoluto else "ENTRADA",
                abs(diferencia),
                referencia_movimiento,
                float(costo),
            ), commit=True)
        return accion, prod_id

    def descargar_plantilla_inicio_dia(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel", "Falta openpyxl. Instala con: pip install openpyxl")
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar plantilla de carga inicial",
            defaultextension=".xlsx",
            initialfile="plantilla_inicio_dia_restaurante.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not ruta:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "INICIO_DIA"
        ws.append(["NOMBRE", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK_INICIAL", "STOCK_MINIMO"])
        rows = self.db.rows("""
            SELECT p.nombre, COALESCE(c.nombre,'') categoria, p.tipo, p.unidad, p.precio, p.costo, p.stock, p.stock_min
            FROM productos p
            LEFT JOIN categorias c ON c.id=p.categoria_id
            WHERE p.activo=1
            ORDER BY p.nombre
        """)
        if rows:
            for r in rows:
                ws.append([r["nombre"], r["categoria"], r["tipo"], r["unidad"], float(r["precio"]), float(r["costo"]), float(r["stock"]), float(r["stock_min"])])
        else:
            ws.append(["POLLO A LA BRASA", "PLATOS", "VENTA", "PLATO", 45.00, 18.00, 30, 5])
            ws.append(["GASEOSA 500ML", "BEBIDAS", "VENTA", "UND", 4.50, 2.20, 60, 12])
            ws.append(["PAPA KG", "INSUMOS", "INSUMO", "KG", 0, 3.80, 25, 10])
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col) + 4
            ws.column_dimensions[col[0].column_letter].width = max(14, largo)
        wb.save(ruta)
        self.set_status(f"Plantilla de inicio del día guardada: {ruta}")

    def carga_masiva_inicio_dia(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel", "Falta openpyxl. Instala con: pip install openpyxl")
            return
        ruta = filedialog.askopenfilename(
            title="Carga masiva de inicio del día",
            filetypes=[("Excel", "*.xlsx *.xlsm")],
        )
        if not ruta:
            return
        if not messagebox.askyesno(
            "Carga inicio día",
            "Esta carga colocará el stock del día según el Excel y actualizará productos automáticamente.\n\n¿Deseas continuar?"
        ):
            return
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb.active
        headers = [str(c.value or "").strip().upper() for c in ws[1]]
        requeridas = ["NOMBRE", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK_INICIAL", "STOCK_MINIMO"]
        faltantes = [h for h in requeridas if h not in headers]
        if faltantes:
            wb.close()
            raise ValueError("Faltan columnas en el Excel de inicio del día: " + ", ".join(faltantes))
        idx = {h: headers.index(h) for h in requeridas}
        procesados = 0
        insertados = 0
        actualizados = 0
        errores = []
        referencia = f"INICIO_DIA {fecha_hoy()}"
        for nro, fila in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            try:
                nombre = str(fila[idx["NOMBRE"]] or "").strip().upper()
                if not nombre:
                    continue
                categoria = str(fila[idx["CATEGORIA"]] or "").strip().upper()
                tipo = str(fila[idx["TIPO"]] or "VENTA").strip().upper()
                unidad = str(fila[idx["UNIDAD"]] or "UND").strip().upper()
                precio = numero(fila[idx["PRECIO"]], 0)
                costo = numero(fila[idx["COSTO"]], 0)
                stock = numero(fila[idx["STOCK_INICIAL"]], 0)
                stock_min = numero(fila[idx["STOCK_MINIMO"]], 0)
                accion, _ = self._upsert_producto_excel(
                    nombre, categoria, tipo, unidad, precio, costo, stock, stock_min,
                    referencia_movimiento=referencia,
                    stock_es_absoluto=True
                )
                if accion == "insertado":
                    insertados += 1
                elif accion == "actualizado":
                    actualizados += 1
                procesados += 1
            except Exception as e:
                errores.append(f"Fila {nro}: {e}")
        wb.close()
        self.actualizar_combos()
        self.cargar_inventario()
        self.filtrar_catalogo_ventas()
        self.mostrar_alerta_stock_en_pantalla()
        mensaje = f"Carga inicial aplicada.\n\nProcesados: {procesados}\nNuevos: {insertados}\nActualizados: {actualizados}"
        if errores:
            mensaje += f"\n\nErrores: {len(errores)}\n" + "\n".join(errores[:10])
        self.set_status(f"Carga de inicio del día completada. Procesados: {procesados}")
        messagebox.showinfo("Carga inicio día", mensaje)


    def descargar_plantilla_productos(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel", "Falta openpyxl. Instala con: pip install openpyxl")
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar plantilla de productos",
            defaultextension=".xlsx",
            initialfile="plantilla_productos_restaurante.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not ruta:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PRODUCTOS"
        ws.append(["NOMBRE", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "STOCK_MINIMO"])
        ejemplos = [
            ["ARROZ CHAUFA", "PLATOS", "VENTA", "PLATO", 18.00, 7.50, 20, 3],
            ["GASEOSA 350ML", "BEBIDAS", "VENTA", "UND", 3.50, 1.80, 40, 8],
            ["ARROZ KG", "INSUMOS", "INSUMO", "KG", 0, 4.20, 50, 10],
        ]
        for fila in ejemplos:
            ws.append(fila)
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col) + 4
            ws.column_dimensions[col[0].column_letter].width = max(14, largo)
        wb.save(ruta)
        self.set_status(f"Plantilla guardada: {ruta}")
    def importar_productos_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel", "Falta openpyxl. Instala con: pip install openpyxl")
            return
        ruta = filedialog.askopenfilename(
            title="Importar productos desde Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm")],
        )
        if not ruta:
            return
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb.active
        headers = [str(c.value or "").strip().upper() for c in ws[1]]
        requeridas = ["NOMBRE", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "STOCK_MINIMO"]
        faltantes = [h for h in requeridas if h not in headers]
        if faltantes:
            wb.close()
            raise ValueError("Faltan columnas en el Excel: " + ", ".join(faltantes))
        idx = {h: headers.index(h) for h in requeridas}
        insertados = 0
        actualizados = 0
        errores = []
        for nro, fila in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            try:
                nombre = str(fila[idx["NOMBRE"]] or "").strip().upper()
                if not nombre:
                    continue
                categoria = str(fila[idx["CATEGORIA"]] or "").strip().upper()
                tipo = str(fila[idx["TIPO"]] or "VENTA").strip().upper()
                unidad = str(fila[idx["UNIDAD"]] or "UND").strip().upper()
                precio = numero(fila[idx["PRECIO"]], 0)
                costo = numero(fila[idx["COSTO"]], 0)
                stock = numero(fila[idx["STOCK"]], 0)
                stock_min = numero(fila[idx["STOCK_MINIMO"]], 0)
                accion, _ = self._upsert_producto_excel(
                    nombre, categoria, tipo, unidad, precio, costo, stock, stock_min,
                    referencia_movimiento=None,
                    stock_es_absoluto=True
                )
                if accion == "insertado":
                    insertados += 1
                elif accion == "actualizado":
                    actualizados += 1
            except Exception as e:
                errores.append(f"Fila {nro}: {e}")
        wb.close()
        self.actualizar_combos()
        self.cargar_inventario()
        self.filtrar_catalogo_ventas()
        self.mostrar_alerta_stock_en_pantalla()
        resumen = f"Importación completada. Nuevos: {insertados} | Actualizados: {actualizados}"
        if errores:
            resumen += f" | Errores: {len(errores)}"
        self.set_status(resumen)
        messagebox.showinfo("Excel", resumen + ("" if not errores else "\n\n" + "\n".join(errores[:10])))


    def exportar_inventario_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel", "Falta openpyxl. Instala con: pip install openpyxl")
            return
        ruta = filedialog.asksaveasfilename(
            title="Exportar inventario",
            defaultextension=".xlsx",
            initialfile=f"inventario_{fecha_hoy()}.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not ruta:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "INVENTARIO"
        ws.append(["ID", "NOMBRE", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "STOCK_MINIMO", "ESTADO"])
        rows = self.db.rows("""
            SELECT p.id, p.nombre, COALESCE(c.nombre,'-') categoria, p.tipo, p.unidad, p.precio, p.costo, p.stock, p.stock_min
            FROM productos p
            LEFT JOIN categorias c ON c.id=p.categoria_id
            ORDER BY p.nombre
        """)
        for r in rows:
            stock = float(r['stock'])
            minimo = float(r['stock_min'])
            estado = 'SIN STOCK' if stock <= 0 else ('STOCK BAJO' if stock <= minimo else 'OK')
            ws.append([r['id'], r['nombre'], r['categoria'], r['tipo'], r['unidad'], float(r['precio']), float(r['costo']), stock, minimo, estado])
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col) + 4
            ws.column_dimensions[col[0].column_letter].width = max(12, largo)
        wb.save(ruta)
        self.set_status(f"Inventario exportado: {ruta}")

    def cargar_producto_seleccionado_en_form(self):
        sel = self.tree_inv.selection()
        if not sel:
            messagebox.showwarning("Inventario", "Selecciona un producto.")
            return
        vals = self.tree_inv.item(sel[0], "values")
        prod = self.producto_por_nombre(vals[1])
        if not prod:
            return
        self.inv_nombre_var.set(prod["nombre"])
        self.inv_categoria_var.set(prod["categoria"] or "")
        self.inv_tipo_var.set(prod["tipo"])
        self.inv_unidad_var.set(prod["unidad"])
        self.inv_precio_var.set(str(prod["precio"]))
        self.inv_costo_var.set(str(prod["costo"]))
        self.inv_stock_var.set(str(prod["stock"]))
        self.inv_min_var.set(str(prod["stock_min"]))
        self.set_status(f"Producto cargado: {prod['nombre']}")

    def movimiento_stock_manual(self, tipo):
        sel = self.tree_inv.selection()
        if not sel:
            messagebox.showwarning("Inventario", "Selecciona un producto.")
            return
        vals = self.tree_inv.item(sel[0], "values")
        prod_id = int(vals[0])
        nombre = vals[1]
        win = tk.Toplevel(self.root)
        win.title(f"{tipo} de stock")
        win.geometry("360x180")
        tk.Label(win, text=nombre, font=("Arial", 12, "bold")).pack(pady=12)
        cantidad_var = tk.StringVar(value="1")
        tk.Label(win, text="Cantidad").pack()
        tk.Entry(win, textvariable=cantidad_var).pack(pady=6)

        def guardar():
            try:
                cant = numero(cantidad_var.get(), 0)
                if cant <= 0:
                    raise ValueError("Cantidad inválida.")
                signo = 1 if tipo == "ENTRADA" else -1
                actual = float(self.db.scalar("SELECT stock FROM productos WHERE id=?", (prod_id,)))
                nuevo_stock = actual + (signo * cant)
                if nuevo_stock < 0:
                    raise ValueError("La salida supera el stock disponible.")
                self.db.execute("UPDATE productos SET stock = ? WHERE id=?", (nuevo_stock, prod_id), commit=True)
                self.db.execute("""
                    INSERT INTO movimientos_inventario(fecha_hora, producto_id, producto, tipo, cantidad, referencia, costo_unit)
                    VALUES (?, ?, ?, ?, ?, 'MANUAL', 0)
                """, (ahora().strftime("%Y-%m-%d %H:%M:%S"), prod_id, nombre, tipo, cant), commit=True)
                win.destroy()
                self.cargar_inventario()
                self.set_status(f"Movimiento {tipo}: {nombre} {cant}")
            except Exception as e:
                messagebox.showerror("Inventario", str(e))

        tk.Button(win, text="Guardar", width=14, command=guardar).pack(pady=12)

    # Recetas
    def obtener_o_crear_receta(self, producto_id):
        receta_id = self.receta_id_de_producto(producto_id)
        if receta_id:
            return receta_id
        return self.db.execute("INSERT INTO recetas(producto_id, observacion) VALUES (?, '')", (producto_id,), commit=True).lastrowid

    def guardar_observacion_receta(self):
        try:
            prod = self.producto_por_nombre(self.receta_producto_var.get().strip())
            if not prod:
                raise ValueError("Selecciona un producto.")
            receta_id = self.obtener_o_crear_receta(prod["id"])
            self.db.execute("UPDATE recetas SET observacion=? WHERE id=?", (self.receta_obs_var.get().strip(), receta_id), commit=True)
            self.set_status(f"Observación guardada para {prod['nombre']}")
        except Exception as e:
            messagebox.showerror("Recetas", str(e))

    def agregar_insumo_a_receta(self):
        try:
            prod = self.producto_por_nombre(self.receta_producto_var.get().strip())
            ins = self.producto_por_nombre(self.receta_insumo_var.get().strip())
            cant = numero(self.receta_cantidad_var.get(), 0)
            if not prod:
                raise ValueError("Selecciona producto de venta.")
            if not ins:
                raise ValueError("Selecciona insumo.")
            if cant <= 0:
                raise ValueError("Cantidad inválida.")
            receta_id = self.obtener_o_crear_receta(prod["id"])
            ex = self.db.row("SELECT id FROM receta_detalle WHERE receta_id=? AND insumo_id=?", (receta_id, ins["id"]))
            if ex:
                self.db.execute("UPDATE receta_detalle SET cantidad=? WHERE id=?", (cant, ex["id"]), commit=True)
            else:
                self.db.execute("INSERT INTO receta_detalle(receta_id, insumo_id, cantidad) VALUES (?, ?, ?)", (receta_id, ins["id"], cant), commit=True)
            self.cargar_receta_producto()
            self.set_status(f"Insumo agregado: {ins['nombre']}")
        except Exception as e:
            messagebox.showerror("Recetas", str(e))

    def quitar_insumo_receta(self):
        sel = self.tree_receta.selection()
        if not sel:
            messagebox.showwarning("Recetas", "Selecciona un insumo.")
            return
        vals = self.tree_receta.item(sel[0], "values")
        self.db.execute("DELETE FROM receta_detalle WHERE id=?", (int(vals[0]),), commit=True)
        self.cargar_receta_producto()
        self.set_status(f"Insumo quitado: {vals[1]}")

    def cargar_receta_producto(self):
        if not hasattr(self, "tree_receta"):
            return
        for i in self.tree_receta.get_children():
            self.tree_receta.delete(i)
        nombre = self.receta_producto_var.get().strip()
        if not nombre:
            return
        prod = self.producto_por_nombre(nombre)
        if not prod:
            return
        receta = self.db.row("SELECT * FROM recetas WHERE producto_id=?", (prod["id"],))
        if receta:
            self.receta_obs_var.set(receta["observacion"] or "")
            dets = self.db.rows("""
                SELECT rd.id, p.nombre insumo, p.unidad, rd.cantidad
                FROM receta_detalle rd
                INNER JOIN productos p ON p.id=rd.insumo_id
                WHERE rd.receta_id=?
                ORDER BY p.nombre
            """, (receta["id"],))
            for d in dets:
                self.tree_receta.insert("", "end", values=(d["id"], d["insumo"], d["unidad"], d["cantidad"]))

    # Caja
    def caja_abierta_actual(self):
        return self.db.row("SELECT * FROM caja WHERE fecha=? AND estado='ABIERTA' ORDER BY id DESC LIMIT 1", (self.fecha_trabajo(),))

    def abrir_caja(self):
        try:
            if not self.validar_dia_abierto("abrir caja"):
                return
            if self.caja_abierta_actual():
                messagebox.showinfo("Caja", "Ya existe una caja abierta hoy.")
                return
            apertura = numero(self.caja_apertura_var.get(), 0)
            self.db.execute("""
                INSERT INTO caja(fecha, turno, apertura, estado, usuario, hora_apertura)
                VALUES (?, ?, ?, 'ABIERTA', ?, ?)
            """, (self.fecha_trabajo(), self.turno_var.get(), apertura, self.user["usuario"], self.ahora_trabajo().strftime("%H:%M:%S")), commit=True)
            self.exportar_respaldo_dia(self.fecha_trabajo())
            self.set_status(f"Caja abierta con {dinero(apertura)} | Fecha operativa: {self.fecha_trabajo()}")
            self.cargar_caja_resumen()
        except Exception as e:
            messagebox.showerror("Caja", str(e))

    def cerrar_caja(self):
        if not self.validar_dia_abierto("cerrar caja"):
            return
        caja = self.caja_abierta_actual()
        if not caja:
            messagebox.showwarning("Caja", "No hay caja abierta.")
            return
        efectivo = float(caja["apertura"]) + float(caja["efectivo_sistema"]) - float(caja["gastos"])
        self.db.execute("""
            UPDATE caja
            SET cierre=?, efectivo_real=?, estado='CERRADA', hora_cierre=?
            WHERE id=?
        """, (efectivo, efectivo, self.ahora_trabajo().strftime("%H:%M:%S"), caja["id"]), commit=True)
        self.exportar_respaldo_dia(self.fecha_trabajo())
        self.set_status("Caja cerrada correctamente.")
        self.cargar_caja_resumen()

    def registrar_movimiento_caja(self, tipo, concepto, monto, metodo="EFECTIVO", referencia=""):
        caja = self.caja_abierta_actual()
        if not caja:
            return
        self.db.execute("""
            INSERT INTO movimientos_caja(caja_id, fecha_hora, tipo, concepto, monto, metodo, referencia)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (caja["id"], self.ahora_trabajo().strftime("%Y-%m-%d %H:%M:%S"), tipo, concepto, monto, metodo, referencia), commit=True)
        if tipo == "INGRESO":
            if metodo == "EFECTIVO":
                self.db.execute("UPDATE caja SET efectivo_sistema = efectivo_sistema + ? WHERE id=?", (monto, caja["id"]), commit=True)
            elif metodo == "YAPE":
                self.db.execute("UPDATE caja SET yape = yape + ? WHERE id=?", (monto, caja["id"]), commit=True)
            elif metodo == "TARJETA":
                self.db.execute("UPDATE caja SET tarjeta = tarjeta + ? WHERE id=?", (monto, caja["id"]), commit=True)
        elif tipo == "EGRESO":
            self.db.execute("UPDATE caja SET gastos = gastos + ? WHERE id=?", (monto, caja["id"]), commit=True)

    def registrar_gasto(self):
        try:
            caja = self.caja_abierta_actual()
            if not caja:
                raise ValueError("Primero abre caja.")
            concepto = self.caja_egreso_concepto_var.get().strip()
            monto = numero(self.caja_egreso_monto_var.get(), 0)
            if not concepto:
                raise ValueError("Ingresa concepto.")
            if monto <= 0:
                raise ValueError("Monto inválido.")
            self.registrar_movimiento_caja("EGRESO", concepto, monto, "EFECTIVO")
            self.caja_egreso_concepto_var.set("")
            self.caja_egreso_monto_var.set("0.00")
            self.cargar_caja_resumen()
            self.set_status(f"Gasto registrado: {concepto}")
        except Exception as e:
            messagebox.showerror("Caja", str(e))

    def cargar_caja_resumen(self):
        caja = self.caja_abierta_actual()
        if not caja:
            self.lbl_caja_estado.config(text="Caja: SIN ABRIR", fg="#B91C1C")
            self.lbl_caja_apertura.config(text="Apertura: S/ 0.00")
            self.lbl_caja_efectivo.config(text="Efectivo sistema: S/ 0.00")
            self.lbl_caja_yape.config(text="Yape: S/ 0.00")
            self.lbl_caja_tarjeta.config(text="Tarjeta: S/ 0.00")
            self.lbl_caja_gastos.config(text="Gastos: S/ 0.00")
            return
        efectivo = float(caja["apertura"]) + float(caja["efectivo_sistema"]) - float(caja["gastos"])
        self.lbl_caja_estado.config(text=f"Caja: {caja['estado']} | Turno {caja['turno']}", fg="#065F46")
        self.lbl_caja_apertura.config(text=f"Apertura: {dinero(caja['apertura'])}")
        self.lbl_caja_efectivo.config(text=f"Efectivo sistema: {dinero(efectivo)}")
        self.lbl_caja_yape.config(text=f"Yape: {dinero(caja['yape'])}")
        self.lbl_caja_tarjeta.config(text=f"Tarjeta: {dinero(caja['tarjeta'])}")
        self.lbl_caja_gastos.config(text=f"Gastos: {dinero(caja['gastos'])}")

    # Stock by recipe

    def validar_stock_para_pago(self, pedido_id):
        detalles = self.db.rows("SELECT * FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
        items = []
        for det in detalles:
            items.append({
                "producto_id": det["producto_id"],
                "producto": det["producto"],
                "cantidad": float(det["cantidad"] or 0),
                "precio": float(det["precio_unit"] or 0),
                "subtotal": float(det["subtotal"] or 0),
            })
        faltantes, _advertencias = self.validar_stock_para_items(items)
        return faltantes

    def descargar_stock_por_detalle(self, det, codigo_pedido):
        prod = self.db.row("SELECT id, nombre, costo FROM productos WHERE id=?", (det["producto_id"],))
        cantidad_det = float(det["cantidad"] or 0)
        receta_id = self.receta_id_de_producto(det["producto_id"])

        # Si el producto tiene receta, se descuentan los insumos.
        # Si NO tiene receta, se descuenta el stock directo del producto vendido.
        if prod and cantidad_det > 0 and not receta_id:
            self.db.execute(
                "UPDATE productos SET stock = CASE WHEN stock - ? < 0 THEN 0 ELSE stock - ? END WHERE id=?",
                (cantidad_det, cantidad_det, prod["id"]),
                commit=True
            )
            self.db.execute("""
                INSERT INTO movimientos_inventario(fecha_hora, producto_id, producto, tipo, cantidad, referencia, costo_unit)
                VALUES (?, ?, ?, 'SALIDA', ?, ?, ?)
            """, (
                self.ahora_trabajo().strftime("%Y-%m-%d %H:%M:%S"), prod["id"], prod["nombre"], cantidad_det,
                f"VENTA {codigo_pedido}", float(prod["costo"] or 0),
            ), commit=True)

        if receta_id:
            insumos = self.db.rows("""
                SELECT p.id, p.nombre, p.costo, rd.cantidad
                FROM receta_detalle rd
                INNER JOIN productos p ON p.id=rd.insumo_id
                WHERE rd.receta_id=?
            """, (receta_id,))
            for ins in insumos:
                cantidad_total = float(ins["cantidad"] or 0) * cantidad_det
                self.db.execute(
                    "UPDATE productos SET stock = CASE WHEN stock - ? < 0 THEN 0 ELSE stock - ? END WHERE id=?",
                    (cantidad_total, cantidad_total, ins["id"]),
                    commit=True
                )
                self.db.execute("""
                    INSERT INTO movimientos_inventario(fecha_hora, producto_id, producto, tipo, cantidad, referencia, costo_unit)
                    VALUES (?, ?, ?, 'SALIDA', ?, ?, ?)
                """, (
                    self.ahora_trabajo().strftime("%Y-%m-%d %H:%M:%S"), ins["id"], ins["nombre"], cantidad_total,
                    f"VENTA {codigo_pedido} / {det['producto']}", float(ins["costo"] or 0),
                ), commit=True)

    def procesar_pago_pedido(self, pedido_id, metodo):
        pedido = self.db.row("SELECT * FROM pedidos WHERE id=?", (pedido_id,))
        if not pedido:
            return
        if pedido["estado"] == "ANULADO":
            raise ValueError("El pedido está anulado.")
        if int(pedido["pagado"]) == 1:
            return
        if not self.caja_abierta_actual():
            raise ValueError("Primero abre caja.")
        faltantes = self.validar_stock_para_pago(pedido_id)
        if faltantes:
            raise ValueError("No hay stock suficiente para cobrar:\n\n" + "\n".join(faltantes[:20]))
        detalles = self.db.rows("SELECT * FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
        for det in detalles:
            self.descargar_stock_por_detalle(det, pedido["codigo"])
        self.registrar_movimiento_caja("INGRESO", f"Venta {pedido['codigo']}", float(pedido["total"]), metodo, pedido["codigo"])
        self.db.execute("""
            UPDATE pedidos
            SET estado='PAGADO', pagado=1, metodo_pago=?
            WHERE id=?
        """, (metodo, pedido_id), commit=True)
        if pedido["mesa_id"]:
            pendientes = self.db.scalar("SELECT COUNT(*) FROM pedidos WHERE COALESCE(mesa_id,0)=COALESCE(?,0) AND pagado=0 AND estado NOT IN ('ANULADO','PAGADO')", (pedido["mesa_id"],))
            if int(pendientes) == 0:
                self.db.execute("UPDATE mesas SET estado='LIBRE' WHERE id=?", (pedido["mesa_id"],), commit=True)
        self.exportar_respaldo_dia(pedido["fecha"])
        self.set_status(f"Pedido cobrado: {pedido['codigo']} - {metodo}")

    def generar_ticket_txt(self, pedido_id, pedido_ids_relacionados=None):
        pedido = self.db.row("""
            SELECT p.*, COALESCE(m.nombre,'-') mesa
            FROM pedidos p
            LEFT JOIN mesas m ON m.id=p.mesa_id
            WHERE p.id=?
        """, (pedido_id,))
        if not pedido:
            return
        pedido_ids = list(pedido_ids_relacionados or [pedido_id])
        detalles = []
        total_general = 0.0
        metodo_pago = pedido["metodo_pago"] or "-"
        for pid in pedido_ids:
            ped = self.db.row("SELECT codigo, total, metodo_pago FROM pedidos WHERE id=?", (pid,))
            if not ped:
                continue
            total_general += float(ped["total"] or 0)
            if ped["metodo_pago"]:
                metodo_pago = ped["metodo_pago"]
            for d in self.db.rows("SELECT * FROM pedido_detalle WHERE pedido_id=? ORDER BY id", (pid,)):
                detalles.append((ped["codigo"], d))
        lineas = []
        lineas.append(self.sucursal_var.get().strip().upper())
        lineas.append("TICKET DE VENTA")
        lineas.append("-" * 40)
        lineas.append(f"Pedido base: {pedido['codigo']}")
        if len(pedido_ids) > 1:
            lineas.append(f"Pedidos   : {', '.join(str(x) for x in pedido_ids)}")
        lineas.append(f"Fecha    : {pedido['fecha']} {pedido['hora']}")
        lineas.append(f"Usuario  : {pedido['usuario']}")
        lineas.append(f"Servicio : {pedido['tipo_servicio']}")
        lineas.append(f"Mesa     : {pedido['mesa']}")
        lineas.append(f"Cliente  : {pedido['cliente']}")
        if pedido["telefono"]:
            lineas.append(f"Telefono : {pedido['telefono']}")
        if pedido["direccion"]:
            lineas.append(f"Direccion: {pedido['direccion']}")
        if pedido["referencia"]:
            lineas.append(f"Ref.     : {pedido['referencia']}")
        lineas.append("-" * 40)
        for codigo, d in detalles:
            lineas.append(f"[{codigo}] {d['producto']}")
            lineas.append(f"  {d['cantidad']} x {dinero(d['precio_unit'])} = {dinero(d['subtotal'])}")
        lineas.append("-" * 40)
        lineas.append(f"TOTAL: {dinero(total_general)}")
        lineas.append(f"PAGO : {metodo_pago or '-'}")
        lineas.append("-" * 40)
        lineas.append("Gracias por su compra")

        nombre_archivo = f"{pedido['codigo']}.txt"
        ruta = os.path.join(self.tickets_dir, nombre_archivo)
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("\n".join(lineas))
        self.set_status(f"Ticket generado: {ruta}")
        return ruta

    def ticket_desde_pedido_seleccionado(self):
        if not self.selected_pedido_id:
            messagebox.showwarning("Ticket", "Selecciona un pedido.")
            return
        pedido_ids = list(self.selected_related_pedido_ids or [self.selected_pedido_id])
        ruta = self.generar_ticket_txt(self.selected_pedido_id, pedido_ids_relacionados=pedido_ids)
        try:
            if ruta and os.name == "nt":
                os.startfile(ruta)
        except Exception:
            pass

    def generar_reporte(self, silencioso=False):
        fi = self.reporte_fecha_inicio_var.get().strip()
        ff = self.reporte_fecha_fin_var.get().strip()
        txt = []

        total = self.db.scalar("""
            SELECT COALESCE(SUM(total),0)
            FROM pedidos
            WHERE fecha BETWEEN ? AND ? AND estado IN ('PAGADO','ENTREGADO')
        """, (fi, ff))
        cantidad = self.db.scalar("""
            SELECT COUNT(*)
            FROM pedidos
            WHERE fecha BETWEEN ? AND ? AND estado IN ('PAGADO','ENTREGADO')
        """, (fi, ff))
        pendientes = self.db.scalar("""
            SELECT COUNT(*)
            FROM pedidos
            WHERE fecha BETWEEN ? AND ? AND estado IN ('PENDIENTE','EN PREPARACIÓN','LISTO')
        """, (fi, ff))
        txt.append("REPORTE GENERAL")
        txt.append("=" * 70)
        txt.append(f"Rango           : {fi} a {ff}")
        txt.append(f"Ventas netas    : {dinero(total)}")
        txt.append(f"Pedidos pagados : {int(cantidad)}")
        txt.append(f"Pendientes      : {int(pendientes)}")
        txt.append("")
        txt.append("TOP PRODUCTOS")
        txt.append("-" * 70)
        for r in self.db.rows("""
            SELECT producto, SUM(cantidad) cant, SUM(subtotal) total
            FROM pedido_detalle d
            INNER JOIN pedidos p ON p.id=d.pedido_id
            WHERE p.fecha BETWEEN ? AND ? AND p.estado IN ('PAGADO','ENTREGADO')
            GROUP BY producto
            ORDER BY cant DESC
            LIMIT 10
        """, (fi, ff)):
            txt.append(f"{r['producto']:<30} | Cant: {r['cant']:<8} | Total: {dinero(r['total'])}")
        txt.append("")
        txt.append("STOCK BAJO")
        txt.append("-" * 70)
        bajos = self.db.rows("""
            SELECT nombre, stock, stock_min, unidad
            FROM productos
            WHERE activo=1 AND stock <= stock_min
            ORDER BY nombre
            LIMIT 20
        """)
        if bajos:
            for r in bajos:
                txt.append(f"{r['nombre']:<30} | Stock: {r['stock']} {r['unidad']} | Minimo: {r['stock_min']}")
        else:
            txt.append("No hay productos en stock bajo.")
        txt.append("")
        txt.append("CLIENTES RECIENTES")
        txt.append("-" * 70)
        for r in self.db.rows("SELECT nombre, telefono, direccion FROM clientes ORDER BY id DESC LIMIT 10"):
            txt.append(f"{r['nombre']:<25} | {r['telefono']:<12} | {r['direccion']}")
        self.txt_reporte.delete("1.0", tk.END)
        self.txt_reporte.insert("1.0", "\n".join(txt))
        if not silencioso:
            self.set_status("Reporte generado.")

    def exportar_ventas_csv(self):
        ruta = filedialog.asksaveasfilename(
            title="Guardar ventas CSV",
            defaultextension=".csv",
            initialfile=f"ventas_{fecha_hoy()}.csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not ruta:
            return
        rows = self.db.rows("""
            SELECT p.id, p.codigo, p.fecha, p.hora, COALESCE(m.nombre,'-') mesa, p.cliente,
                   p.tipo_servicio, p.estado, p.total, p.pagado, p.metodo_pago, p.usuario
            FROM pedidos p
            LEFT JOIN mesas m ON m.id=p.mesa_id
            ORDER BY p.id DESC
        """)
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["ID", "CODIGO", "FECHA", "HORA", "MESA", "CLIENTE", "TIPO", "ESTADO", "TOTAL", "PAGADO", "METODO_PAGO", "USUARIO"])
            for r in rows:
                w.writerow([r["id"], r["codigo"], r["fecha"], r["hora"], r["mesa"], r["cliente"], r["tipo_servicio"], r["estado"], r["total"], r["pagado"], r["metodo_pago"], r["usuario"]])
        self.set_status(f"Ventas exportadas a {ruta}")

    # Admin / demo
    def sembrar_demo(self):
        try:
            if not self.caja_abierta_actual():
                self.db.execute("""
                    INSERT INTO caja(fecha, turno, apertura, estado, usuario, hora_apertura)
                    VALUES (?, 'MAÑANA', 150, 'ABIERTA', ?, ?)
                """, (fecha_hoy(), self.user["usuario"], ahora().strftime("%H:%M:%S")), commit=True)
            if not self.cliente_por_nombre("Carlos Delivery"):
                self.db.execute("""
                    INSERT INTO clientes(nombre, telefono, direccion, referencia, notas)
                    VALUES ('Carlos Delivery', '999111222', 'Av. Principal 123', 'Puerta azul', 'Cliente frecuente')
                """, commit=True)
            self.refrescar_todo()
            self.set_status("Datos demo generados.")
        except Exception as e:
            messagebox.showerror("Demo", str(e))

    def registrar_sucursal_admin(self):
        try:
            nombre = self.normalizar_texto(self.nueva_sucursal_var.get())
            if not nombre:
                raise ValueError("Ingresa el nombre de la sucursal.")
            existentes = self.sucursales_disponibles()
            if nombre in existentes:
                self.sucursal_var.set(nombre)
                self.set_status(f"La sucursal ya existe: {nombre}")
                return
            clave = f"sucursal_extra_{int(ahora().timestamp())}"
            self.set_config(clave, nombre)
            self.sucursal_var.set(nombre)
            self.nueva_sucursal_var.set("")
            self.actualizar_combos()
            self.set_status(f"Sucursal registrada: {nombre}")
        except Exception as e:
            messagebox.showerror("Sucursales", str(e))

    def guardar_usuario_admin(self):
        try:
            usuario = self.normalizar_texto(self.nuevo_usuario_var.get())
            nombre = self.normalizar_texto(self.nuevo_nombre_usuario_var.get())
            clave = str(self.nueva_clave_usuario_var.get()).strip()
            rol = self.normalizar_texto(self.nuevo_rol_usuario_var.get()) or "MESERO"
            if not usuario:
                raise ValueError("Ingresa el usuario.")
            if not nombre:
                raise ValueError("Ingresa el nombre.")
            if not clave:
                raise ValueError("Ingresa la clave.")
            row = self.db.row("SELECT id FROM usuarios WHERE UPPER(usuario)=?", (usuario,))
            if row:
                self.db.execute("UPDATE usuarios SET usuario=?, clave=?, nombre=?, rol=?, activo=1 WHERE id=?", (usuario, clave, nombre, rol, row["id"]), commit=True)
                msg = f"Usuario actualizado: {usuario}"
            else:
                self.db.execute("INSERT INTO usuarios(usuario, clave, nombre, rol, activo) VALUES (?, ?, ?, ?, 1)", (usuario, clave, nombre, rol), commit=True)
                msg = f"Usuario registrado: {usuario}"
            self.nuevo_usuario_var.set("")
            self.nuevo_nombre_usuario_var.set("")
            self.nueva_clave_usuario_var.set("")
            self.nuevo_rol_usuario_var.set("MESERO")
            self.actualizar_combos()
            self.set_status(msg)
            messagebox.showinfo("Usuarios", msg)
        except Exception as e:
            messagebox.showerror("Usuarios", str(e))

    def crear_usuario_demo(self):
        try:
            base = f"user{int(ahora().timestamp())%10000}"
            self.db.execute("""
                INSERT INTO usuarios(usuario, clave, nombre, rol, activo)
                VALUES (?, '1234', ?, 'MESERO', 1)
            """, (base, f"Usuario {base}"), commit=True)
            self.set_status(f"Usuario demo creado: {base} / 1234")
            messagebox.showinfo("Usuarios", f"Usuario creado:\n{base}\nClave: 1234")
        except Exception as e:
            messagebox.showerror("Usuarios", str(e))

    def run(self):
        self.root.mainloop()


def main():
    db = Database(os.path.join(APP_DIR, DB_NAME))
    login = LoginWindow(db)
    user = login.run()
    if not user:
        return
    app = RestauranteAppV3(db, user)
    app.run()


if __name__ == "__main__":
    main()