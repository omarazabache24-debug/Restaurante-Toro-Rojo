# -*- coding: utf-8 -*-
"""
RESTAURANTE AORIX - APP WEB PRO COMPLETA
Adaptación Flask lista para GitHub + Render.
Integra la lógica principal del Restaurante_AORIX desktop:
- Panel principal
- Ventas / pedidos
- Inventario
- Recetas con descuento de insumos
- Caja
- Delivery / clientes
- Indicadores
- Reportes + Excel/CSV
- Administrador / usuarios / sucursales
- Log

Usuarios iniciales:
- admin / admin123
- caja / caja123
"""

import csv
import os
import sqlite3
from datetime import datetime
from functools import wraps
from io import BytesIO, StringIO
from zoneinfo import ZoneInfo

from flask import (
    Flask, flash, redirect, render_template_string, request,
    send_file, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash

try:
    from openpyxl import Workbook, load_workbook
    OPENPYXL = True
except Exception:
    Workbook = None
    load_workbook = None
    OPENPYXL = False

# =========================
# CONFIG RENDER / LOCAL
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else os.path.join(BASE_DIR, "data"))
os.makedirs(PERSIST_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, "restaurante_aorix_pro.db")
APP_TZ = ZoneInfo(os.getenv("APP_TIMEZONE", "America/Lima"))

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "aorix-restaurante-pro-2026")

APP_TITLE = "RESTAURANTE AORIX"
APP_SUBTITLE = "Sistema de Control y Gestión de Alimentos"
BRAND = "AORIX SYSTEMS - Automatizamos tu empresa"

# =========================
# HELPERS DB
# =========================
def now():
    return datetime.now(APP_TZ)

def today():
    return now().date().isoformat()

def hour():
    return now().strftime("%H:%M:%S")

def money(v):
    try:
        return "S/ {:,.2f}".format(float(v or 0))
    except Exception:
        return "S/ 0.00"

def clean(v):
    return str(v or "").strip()

def up(v):
    return clean(v).upper()

def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def q_all(sql, params=()):
    with conn() as c:
        return c.execute(sql, params).fetchall()

def q_one(sql, params=()):
    rows = q_all(sql, params)
    return rows[0] if rows else None

def q_exec(sql, params=()):
    with conn() as c:
        cur = c.execute(sql, params)
        c.commit()
        return cur.lastrowid

def log_event(accion, detalle=""):
    try:
        q_exec(
            "INSERT INTO logs(fecha,hora,usuario,accion,detalle) VALUES(?,?,?,?,?)",
            (today(), hour(), session.get("user", "sistema"), accion, detalle),
        )
    except Exception:
        pass

# =========================
# INIT DB
# =========================
def init_db():
    with conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS usuarios(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT UNIQUE NOT NULL,
                nombre TEXT DEFAULT '',
                clave_hash TEXT NOT NULL,
                rol TEXT DEFAULT 'CAJA',
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS sucursales(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS productos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                nombre TEXT UNIQUE NOT NULL,
                categoria TEXT DEFAULT 'PLATOS',
                tipo TEXT DEFAULT 'VENTA',
                unidad TEXT DEFAULT 'PLATO',
                precio REAL DEFAULT 0,
                costo REAL DEFAULT 0,
                stock REAL DEFAULT 0,
                stock_min REAL DEFAULT 0,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS insumos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                unidad TEXT DEFAULT 'UND',
                stock REAL DEFAULT 0,
                stock_min REAL DEFAULT 0,
                costo REAL DEFAULT 0,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS recetas(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                producto_id INTEGER NOT NULL,
                insumo_id INTEGER NOT NULL,
                cantidad REAL DEFAULT 0,
                observacion TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS clientes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                telefono TEXT DEFAULT '',
                direccion TEXT DEFAULT '',
                referencia TEXT DEFAULT '',
                notas TEXT DEFAULT '',
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS pedidos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                fecha TEXT,
                hora TEXT,
                mesa TEXT DEFAULT '',
                cliente TEXT DEFAULT '',
                telefono TEXT DEFAULT '',
                direccion TEXT DEFAULT '',
                referencia TEXT DEFAULT '',
                servicio TEXT DEFAULT 'SALÓN',
                metodo_pago TEXT DEFAULT 'EFECTIVO',
                subtotal REAL DEFAULT 0,
                descuento REAL DEFAULT 0,
                total REAL DEFAULT 0,
                estado TEXT DEFAULT 'PENDIENTE',
                pagado TEXT DEFAULT 'NO',
                usuario TEXT DEFAULT '',
                observacion TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS pedido_detalle(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id INTEGER,
                producto_id INTEGER,
                producto TEXT,
                cantidad REAL,
                precio REAL,
                total REAL
            );
            CREATE TABLE IF NOT EXISTS ventas(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                pedido_id INTEGER DEFAULT NULL,
                cliente TEXT DEFAULT '',
                servicio TEXT DEFAULT 'SALÓN',
                metodo_pago TEXT DEFAULT 'EFECTIVO',
                subtotal REAL DEFAULT 0,
                descuento REAL DEFAULT 0,
                total REAL DEFAULT 0,
                usuario TEXT DEFAULT '',
                estado TEXT DEFAULT 'PAGADO'
            );
            CREATE TABLE IF NOT EXISTS venta_detalle(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                venta_id INTEGER,
                producto_id INTEGER,
                producto TEXT,
                cantidad REAL,
                precio REAL,
                total REAL
            );
            CREATE TABLE IF NOT EXISTS caja(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                tipo TEXT,
                concepto TEXT,
                monto REAL,
                usuario TEXT,
                venta_id INTEGER DEFAULT NULL,
                estado TEXT DEFAULT 'OK'
            );
            CREATE TABLE IF NOT EXISTS contexto(
                clave TEXT PRIMARY KEY,
                valor TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS logs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                usuario TEXT,
                accion TEXT,
                detalle TEXT
            );
            """
        )
        c.commit()

    defaults = [("Sucursal Principal",), ("Delivery",)]
    for (nombre,) in defaults:
        if not q_one("SELECT id FROM sucursales WHERE nombre=?", (nombre,)):
            q_exec("INSERT INTO sucursales(nombre,activo) VALUES(?,1)", (nombre,))

    for usuario, nombre, clave, rol in [
        ("admin", "Administrador", "admin123", "ADMIN"),
        ("caja", "Caja", "caja123", "CAJA"),
        ("mozo", "Mozo", "mozo123", "MESERO"),
    ]:
        if not q_one("SELECT id FROM usuarios WHERE usuario=?", (usuario,)):
            q_exec(
                "INSERT INTO usuarios(usuario,nombre,clave_hash,rol,activo) VALUES(?,?,?,?,1)",
                (usuario, nombre, generate_password_hash(clave), rol),
            )

    for k, v in {
        "sucursal": "Sucursal Principal",
        "turno": "MAÑANA",
        "dia_abierto": today(),
        "caja_abierta": "0",
        "monto_apertura": "0",
    }.items():
        if not q_one("SELECT clave FROM contexto WHERE clave=?", (k,)):
            q_exec("INSERT INTO contexto(clave,valor) VALUES(?,?)", (k, v))

    if not q_one("SELECT id FROM productos LIMIT 1"):
        demo_productos = [
            ("P001", "1/2 POLLO", "PLATOS", "VENTA", "PLATO", 35, 10, 40, 5),
            ("P002", "1/4 POLLO", "PLATOS", "VENTA", "PLATO", 18, 8, 0, 0),
            ("P003", "1/4 POLLO BROSTER", "PLATOS", "VENTA", "PLATO", 22, 0, 0, 0),
            ("P004", "1/8 POLLO", "PLATOS", "VENTA", "PLATO", 12, 0, 0, 0),
            ("P005", "ADICIONAL CARNE", "ADICIONALES", "VENTA", "UND", 5, 0, 0, 0),
            ("P006", "GASEOSA PERSONAL", "BEBIDAS", "VENTA", "UND", 4, 2, 30, 5),
        ]
        for p in demo_productos:
            q_exec(
                "INSERT INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)",
                p,
            )
    if not q_one("SELECT id FROM insumos LIMIT 1"):
        for nombre, unidad, stock, stock_min, costo in [
            ("POLLO ENTERO", "UND", 20, 3, 20),
            ("PAPA KG", "KG", 30, 5, 3),
            ("ENSALADA", "PORCION", 50, 10, 1),
            ("MAYONESA", "PORCION", 50, 10, 0.5),
            ("AJI", "PORCION", 50, 10, 0.5),
            ("ENVASE", "UND", 50, 10, 1),
        ]:
            q_exec(
                "INSERT INTO insumos(nombre,unidad,stock,stock_min,costo,activo) VALUES(?,?,?,?,?,1)",
                (nombre, unidad, stock, stock_min, costo),
            )

init_db()

# =========================
# AUTH
# =========================
def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("rol") != "ADMIN":
            flash("Solo el administrador puede ingresar a esta opción.", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper

# =========================
# UI
# =========================
BASE_HTML = r'''
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{title}}</title>
<style>
:root{--navy:#082238;--navy2:#0B2D4A;--green:#66D17B;--green2:#16a34a;--blue:#0d73b8;--bg:#f2f4f7;--line:#cbd5e1;--text:#111827;--muted:#64748b;--red:#dc2626;--orange:#f97316}*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text)}a{text-decoration:none;color:inherit}button,.btn{border:0;background:#e5e7eb;border-radius:4px;padding:10px 16px;font-weight:800;cursor:pointer;box-shadow:inset 0 0 0 1px #999;color:#0f172a;display:inline-block}.btn-primary,button.primary{background:#0B2D4A;color:white}.btn-green{background:#66D17B;color:#082238}.btn-red{background:#fecaca;color:#991b1b}.btn-orange{background:#fb923c;color:white}.btn-blue{background:#0d73b8;color:white}input,select,textarea{width:100%;border:1px solid #b6b6b6;background:white;padding:8px 9px;min-height:34px;color:#0f172a}label{font-weight:800;color:#334155}.login-page{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#061b2b,#0B2D4A);padding:22px}.login-card{width:min(430px,94vw);background:white;border-radius:16px;padding:28px;box-shadow:0 25px 70px rgba(0,0,0,.28);text-align:center}.logo{font-size:44px;font-weight:950;color:#0B2D4A;letter-spacing:-2px}.logo span{color:#f97316}.login-card label{text-align:left;display:block;margin:12px 0 5px}.login-card button{width:100%;margin-top:16px;background:#0B2D4A;color:white}.hint{font-size:12px;color:#64748b;margin-top:14px;line-height:1.5}.app{display:grid;grid-template-columns:210px minmax(0,1fr);min-height:100vh}.side{background:#082238;color:white;height:100vh;position:sticky;top:0;overflow:auto}.brand{padding:16px;text-align:center;border-bottom:1px solid rgba(255,255,255,.15)}.brand .logo{font-size:34px;color:white}.brand small{color:#dbeafe;font-weight:800}.nav a{display:flex;align-items:center;gap:8px;padding:12px 14px;font-weight:900;border-bottom:1px solid rgba(255,255,255,.05);color:#eef6ff}.nav a.on,.nav a:hover{background:#66D17B;color:#082238}.main{min-width:0}.topbar{background:#082238;color:white;padding:14px 20px;text-align:center}.topbar h1{margin:0;font-size:28px}.topbar p{margin:3px 0 0;color:#dbeafe;font-weight:700}.content{padding:14px;max-width:1700px;margin:0 auto}.flash{padding:12px 14px;border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;font-weight:800;margin-bottom:10px}.flash.error{background:#fff1f2;border-color:#fecaca;color:#991b1b}.flash.ok{background:#ecfdf5;border-color:#bbf7d0;color:#166534}.tabs{display:flex;gap:0;flex-wrap:wrap;background:#ececec;border-bottom:1px solid #aaa;margin:-14px -14px 14px}.tabs a{padding:16px 24px;font-size:18px;font-weight:950;border:1px solid #aaa;border-bottom:0;background:#e5e5e5;color:#082238}.tabs a.on{background:#66D17B}.panel{border:1px solid #999;background:#f8fafc;padding:14px;margin-bottom:14px}.panel legend,.box-title{font-weight:900;color:#0f172a}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;align-items:end}.grid5{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;align-items:end}.grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}.actions{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.kpi{background:white;border:1px solid #222;padding:24px;text-align:center}.kpi h3{margin:0 0 16px;color:#334155}.kpi b{font-size:32px;color:#082238}.kpi .red{color:#dc2626}.table-wrap{overflow:auto;border:1px solid #999;background:white}.table-wrap.small{max-height:360px}.table-wrap table{width:100%;border-collapse:collapse;min-width:900px}th{background:#0B2D4A;color:white;font-size:15px}th,td{border:1px solid #b6b6b6;padding:9px;text-align:center;white-space:nowrap}.row-ok td{background:#e8fff1;color:#047857}.row-bad td{background:#ffe1e1;color:#dc2626}.row-sel td{background:#51728d!important;color:white!important}.badge{display:inline-block;padding:5px 10px;border-radius:999px;font-weight:950;font-size:12px}.ok{background:#dcfce7;color:#166534}.warn{background:#fef3c7;color:#92400e}.off{background:#fee2e2;color:#991b1b}.muted{color:#64748b}.report-box{width:100%;min-height:420px;font-family:Consolas,monospace;white-space:pre;background:white}.chart{height:360px;border:1px solid #cbd5e1;background:white;padding:18px;display:flex;align-items:end;gap:18px}.bar{width:70px;background:#66D17B;border:1px solid #0f766e;display:flex;align-items:flex-start;justify-content:center;font-weight:900;padding-top:4px;color:#082238}.bar-wrap{text-align:center}.mobile-card{display:none}@media(max-width:900px){.app{display:block}.side{height:auto;position:relative}.brand{display:none}.nav{display:grid;grid-template-columns:repeat(2,1fr)}.nav a{justify-content:center;text-align:center}.topbar h1{font-size:22px}.tabs{overflow-x:auto;flex-wrap:nowrap}.tabs a{font-size:14px;padding:12px}.grid,.grid2,.grid5,.kpis{grid-template-columns:1fr}.content{padding:10px}.panel{padding:10px}.table-wrap{max-height:55vh}button,.btn{width:100%;text-align:center}.desktop-only{display:none}.mobile-card{display:block;background:white;border:1px solid #cbd5e1;padding:12px;margin-bottom:10px;border-radius:10px}.mobile-card b{display:inline-block;min-width:115px;color:#64748b}}
</style>
</head>
<body>
{% if session.get('user') %}
<div class="app">
  <aside class="side">
    <div class="brand"><div class="logo">AOR<span>IX</span></div><small>{{brand}}</small><br><small>{{session.get('user')}} - {{session.get('rol')}}</small></div>
    <nav class="nav">
      <a class="{{'on' if active=='dashboard' else ''}}" href="{{url_for('dashboard')}}">📊 Panel principal</a>
      <a class="{{'on' if active=='ventas' else ''}}" href="{{url_for('ventas')}}">🧾 Ventas</a>
      <a class="{{'on' if active=='pedidos' else ''}}" href="{{url_for('pedidos')}}">🚚 Pedidos</a>
      <a class="{{'on' if active=='inventario' else ''}}" href="{{url_for('inventario')}}">📦 Inventario</a>
      <a class="{{'on' if active=='recetas' else ''}}" href="{{url_for('recetas')}}">🍽️ Recetas</a>
      <a class="{{'on' if active=='caja' else ''}}" href="{{url_for('caja')}}">💵 Caja</a>
      <a class="{{'on' if active=='delivery' else ''}}" href="{{url_for('delivery')}}">🛵 Delivery</a>
      <a class="{{'on' if active=='indicadores' else ''}}" href="{{url_for('indicadores')}}">📈 Indicadores</a>
      <a class="{{'on' if active=='reportes' else ''}}" href="{{url_for('reportes')}}">📄 Reportes</a>
      <a class="{{'on' if active=='admin' else ''}}" href="{{url_for('admin')}}">⚙️ Administrador</a>
      <a class="{{'on' if active=='log' else ''}}" href="{{url_for('logs')}}">🧾 Log</a>
      <a href="{{url_for('logout')}}">🚪 Salir</a>
    </nav>
  </aside>
  <main class="main">
    <header class="topbar"><h1>{{title}}</h1><p>{{subtitle}}</p></header>
    <div class="content">
      <div class="tabs">
        {% for key,label,endpoint in tabs %}<a class="{{'on' if active==key else ''}}" href="{{url_for(endpoint)}}">{{label}}</a>{% endfor %}
      </div>
      {% with msgs=get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}{% endwith %}
      {{content|safe}}
    </div>
  </main>
</div>
{% else %}
{{content|safe}}
{% endif %}
</body>
</html>
'''

def get_ctx(k, default=""):
    r = q_one("SELECT valor FROM contexto WHERE clave=?", (k,))
    return r["valor"] if r else default

def set_ctx(k, v):
    if q_one("SELECT clave FROM contexto WHERE clave=?", (k,)):
        q_exec("UPDATE contexto SET valor=? WHERE clave=?", (str(v), k))
    else:
        q_exec("INSERT INTO contexto(clave,valor) VALUES(?,?)", (k, str(v)))

def tabs():
    return [
        ("dashboard", "Panel Principal", "dashboard"),
        ("ventas", "Ventas", "ventas"),
        ("pedidos", "Pedidos", "pedidos"),
        ("inventario", "Inventario", "inventario"),
        ("recetas", "Recetas", "recetas"),
        ("caja", "Caja", "caja"),
        ("delivery", "Delivery", "delivery"),
        ("indicadores", "Indicadores", "indicadores"),
        ("reportes", "Reportes", "reportes"),
        ("admin", "Administrador", "admin"),
        ("log", "Log", "logs"),
    ]

def page(content, active="dashboard"):
    return render_template_string(
        BASE_HTML,
        content=content,
        active=active,
        title=APP_TITLE,
        subtitle=APP_SUBTITLE,
        brand=BRAND,
        tabs=tabs(),
        money=money,
    )

def select_options(rows, value_key="id", text_key="nombre", selected=None):
    out = []
    for r in rows:
        val = str(r[value_key])
        sel = "selected" if selected is not None and str(selected) == val else ""
        out.append(f'<option value="{val}" {sel}>{r[text_key]}</option>')
    return "".join(out)

# =========================
# LÓGICA NEGOCIO
# =========================
def descontar_receta(producto_id, cantidad_producto):
    for r in q_all("SELECT * FROM recetas WHERE producto_id=?", (producto_id,)):
        q_exec(
            "UPDATE insumos SET stock=COALESCE(stock,0)-? WHERE id=?",
            (float(r["cantidad"] or 0) * float(cantidad_producto or 0), r["insumo_id"]),
        )

def descontar_producto(producto_id, cantidad):
    q_exec("UPDATE productos SET stock=COALESCE(stock,0)-? WHERE id=?", (float(cantidad or 0), producto_id))
    descontar_receta(producto_id, cantidad)

def crear_venta_desde_pedido(pedido_id, metodo_pago="EFECTIVO"):
    p = q_one("SELECT * FROM pedidos WHERE id=?", (pedido_id,))
    if not p:
        return None
    if p["pagado"] == "SI":
        return None
    vid = q_exec(
        "INSERT INTO ventas(fecha,hora,pedido_id,cliente,servicio,metodo_pago,subtotal,descuento,total,usuario,estado) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
        (today(), hour(), pedido_id, p["cliente"], p["servicio"], metodo_pago, p["subtotal"], p["descuento"], p["total"], session.get("user", ""), "PAGADO"),
    )
    for d in q_all("SELECT * FROM pedido_detalle WHERE pedido_id=?", (pedido_id,)):
        q_exec(
            "INSERT INTO venta_detalle(venta_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)",
            (vid, d["producto_id"], d["producto"], d["cantidad"], d["precio"], d["total"]),
        )
        descontar_producto(d["producto_id"], d["cantidad"])
    q_exec("UPDATE pedidos SET pagado='SI', estado='PAGADO' WHERE id=?", (pedido_id,))
    q_exec(
        "INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario,venta_id) VALUES(?,?,?,?,?,?,?)",
        (today(), hour(), "INGRESO", f"VENTA PEDIDO {p['codigo']}", p["total"], session.get("user", ""), vid),
    )
    log_event("VENTA", f"Pedido cobrado {p['codigo']}")
    return vid

# =========================
# ROUTES AUTH
# =========================
@app.route("/", methods=["GET", "POST"])
def login():
    if session.get("user"):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        usuario = clean(request.form.get("usuario"))
        clave = request.form.get("clave", "")
        r = q_one("SELECT * FROM usuarios WHERE usuario=? AND activo=1", (usuario,))
        if r and check_password_hash(r["clave_hash"], clave):
            session["user"] = r["usuario"]
            session["nombre"] = r["nombre"]
            session["rol"] = r["rol"]
            log_event("LOGIN", "Ingreso correcto")
            return redirect(url_for("dashboard"))
        flash("Usuario o clave incorrectos.", "error")
    html = """
    <div class="login-page">
      <form class="login-card" method="post">
        <div class="logo">AOR<span>IX</span></div>
        <h2>Restaurante AORIX</h2>
        <p class="muted">Acceso al sistema</p>
        <label>Usuario</label><input name="usuario" placeholder="Ingrese su usuario" autofocus>
        <label>Clave</label><input name="clave" type="password" placeholder="Ingrese su clave">
        <button>Ingresar</button>
        <div class="hint">Demo: admin / admin123<br>Caja: caja / caja123<br>Mozo: mozo / mozo123</div>
      </form>
    </div>
    """
    return page(html)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# =========================
# DASHBOARD
# =========================
@app.route("/dashboard", methods=["GET", "POST"])
@login_required
def dashboard():
    if request.method == "POST":
        set_ctx("sucursal", request.form.get("sucursal", "Sucursal Principal"))
        set_ctx("turno", request.form.get("turno", "MAÑANA"))
        set_ctx("dia_abierto", request.form.get("fecha", today()))
        flash("Contexto guardado.", "ok")
        return redirect(url_for("dashboard"))

    f = get_ctx("dia_abierto", today())
    ventas_hoy = q_one("SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha=?", (f,))
    pedidos_act = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha=? AND estado NOT IN ('ENTREGADO','PAGADO')", (f,))["c"]
    mesas = q_one("SELECT COUNT(DISTINCT mesa) c FROM pedidos WHERE fecha=? AND mesa<>'' AND estado NOT IN ('ENTREGADO','PAGADO')", (f,))["c"]
    stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min", ())["c"]
    sucursales = q_all("SELECT * FROM sucursales WHERE activo=1 ORDER BY nombre")
    usuarios = q_all("SELECT * FROM usuarios WHERE activo=1 ORDER BY usuario")
    opts_suc = select_options(sucursales, "nombre", "nombre", get_ctx("sucursal"))
    opts_user = select_options(usuarios, "usuario", "usuario", session.get("user"))
    html = f"""
    <form method="post" class="panel">
      <div class="box-title">Contexto de trabajo</div><br>
      <div class="grid5">
        <div><label>Sucursal</label><select name="sucursal">{opts_suc}</select></div>
        <div><label>Fecha</label><input type="date" name="fecha" value="{f}"></div>
        <div><label>Turno</label><select name="turno"><option>MAÑANA</option><option>TARDE</option><option>NOCHE</option></select></div>
        <div><label>Usuario</label><select name="usuario" disabled>{opts_user}</select></div>
        <div><label>Rol</label><input value="{session.get('rol','')}" readonly></div>
      </div><br>
      <div class="actions"><button class="primary">Guardar contexto</button><a class="btn" href="{url_for('dashboard')}">Recargar panel</a></div>
    </form>
    <div class="actions" style="margin-bottom:14px"><b style="font-size:22px;color:#047857">● DÍA ABIERTO: {f}</b><a class="btn btn-red" href="{url_for('cerrar_dia')}">🔒 Cerrar día</a><a class="btn" href="{url_for('reportes',fi=f,ff=f)}">Ver resumen</a><a class="btn" href="{url_for('reabrir_dia')}">Reabrir día</a></div>
    <div class="panel"><div class="box-title">Indicadores</div><br><div class="kpis">
      <div class="kpi"><h3>Ventas hoy</h3><b>{money(ventas_hoy['t'])}</b></div>
      <div class="kpi"><h3>Pedidos activos</h3><b>{pedidos_act}</b></div>
      <div class="kpi"><h3>Mesas ocupadas</h3><b class="red">{mesas}</b></div>
      <div class="kpi"><h3>Stock bajo</h3><b class="red">{stock_bajo}</b></div>
    </div></div>
    """
    return page(html, "dashboard")

@app.route("/cerrar_dia")
@login_required
def cerrar_dia():
    set_ctx("dia_cerrado", get_ctx("dia_abierto", today()))
    flash("Día cerrado de forma lógica.", "ok")
    return redirect(url_for("dashboard"))

@app.route("/reabrir_dia")
@login_required
def reabrir_dia():
    set_ctx("dia_cerrado", "")
    flash("Día reabierto.", "ok")
    return redirect(url_for("dashboard"))

# =========================
# VENTAS
# =========================
@app.route("/ventas", methods=["GET", "POST"])
@login_required
def ventas():
    if request.method == "POST":
        accion = request.form.get("accion", "guardar_pedido")
        if accion == "guardar_pedido":
            producto_id = int(request.form.get("producto_id") or 0)
            prod = q_one("SELECT * FROM productos WHERE id=? AND activo=1", (producto_id,))
            if not prod:
                flash("Selecciona un producto válido.", "error")
                return redirect(url_for("ventas"))
            cantidad = float(request.form.get("cantidad") or 1)
            descuento = float(request.form.get("descuento") or 0)
            precio = float(prod["precio"] or 0)
            subtotal = precio * cantidad
            total = max(subtotal - descuento, 0)
            codigo = "PED-" + now().strftime("%Y%m%d-%H%M%S")
            pedido_id = q_exec(
                "INSERT INTO pedidos(codigo,fecha,hora,mesa,cliente,telefono,direccion,referencia,servicio,metodo_pago,subtotal,descuento,total,estado,pagado,usuario,observacion) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (codigo, today(), hour(), request.form.get("mesa", ""), up(request.form.get("cliente")), clean(request.form.get("telefono")), up(request.form.get("direccion")), up(request.form.get("referencia")), request.form.get("servicio", "SALÓN"), request.form.get("metodo_pago", "EFECTIVO"), subtotal, descuento, total, "PENDIENTE", "NO", session.get("user"), ""),
            )
            q_exec(
                "INSERT INTO pedido_detalle(pedido_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)",
                (pedido_id, producto_id, prod["nombre"], cantidad, precio, total),
            )
            log_event("PEDIDO", f"Creado {codigo}")
            flash("Pedido guardado correctamente.", "ok")
            return redirect(url_for("ventas"))
        if accion == "cobrar_ticket":
            pedido_id = int(request.form.get("pedido_id") or 0)
            metodo = request.form.get("metodo_pago", "EFECTIVO")
            vid = crear_venta_desde_pedido(pedido_id, metodo)
            if vid:
                flash("Pedido cobrado y ticket generado.", "ok")
            else:
                flash("No se pudo cobrar: verifica que el pedido exista y no esté pagado.", "error")
            return redirect(url_for("ventas"))

    productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY nombre")
    opts_prod = "".join(f'<option value="{p["id"]}">{p["nombre"]} - {money(p["precio"])} - Stock {p["stock"]}</option>' for p in productos)
    rows = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY CASE WHEN stock>0 THEN 0 ELSE 1 END, nombre LIMIT 120")
    tr_prod = "".join(
        f'<tr class="{"row-ok" if float(r["stock"] or 0)>0 else "row-bad"}"><td>{r["id"]}</td><td>{r["nombre"]}</td><td>{r["categoria"]}</td><td>{money(r["precio"])}</td><td>{r["stock"]}</td><td>{"DISPONIBLE" if float(r["stock"] or 0)>0 else "SIN STOCK"}</td></tr>'
        for r in rows
    )
    pedidos_pend = q_all("SELECT * FROM pedidos ORDER BY id DESC LIMIT 20")
    opts_ped = "".join(f'<option value="{p["id"]}">{p["codigo"]} - {p["cliente"]} - {money(p["total"])} - {p["estado"]}</option>' for p in pedidos_pend)
    detalle = q_all("SELECT d.*,p.codigo FROM pedido_detalle d JOIN pedidos p ON p.id=d.pedido_id ORDER BY d.id DESC LIMIT 30")
    tr_det = "".join(f'<tr><td>{r["codigo"]}</td><td>{r["producto"]}</td><td>{r["cantidad"]}</td><td>{money(r["precio"])}</td><td>{money(r["total"])}</td></tr>' for r in detalle) or '<tr><td colspan="5">Sin detalle.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Nueva venta / pedido</div><br>
      <form method="post">
        <input type="hidden" name="accion" value="guardar_pedido">
        <div class="grid5">
          <div><label>Mesa</label><select name="mesa"><option></option><option>MESA 1</option><option>MESA 2</option><option>MESA 3</option><option>MESA 4</option></select></div>
          <div><label>Tipo servicio</label><select name="servicio"><option>SALÓN</option><option>DELIVERY</option><option>RECOJO</option></select></div>
          <div><label>Cliente</label><input name="cliente" placeholder="Cliente"></div>
          <div><label>Teléfono</label><input name="telefono" placeholder="Teléfono"></div>
          <div><label>Dirección</label><input name="direccion" placeholder="Dirección"></div>
          <div><label>Referencia</label><input name="referencia" placeholder="Referencia"></div>
          <div><label>Producto</label><select name="producto_id">{opts_prod}</select></div>
          <div><label>Cantidad</label><input name="cantidad" type="number" step="0.01" value="1"></div>
          <div><label>Método pago</label><select name="metodo_pago"><option>EFECTIVO</option><option>YAPE</option><option>TARJETA</option><option>TRANSFERENCIA</option></select></div>
          <div><label>Descuento</label><input name="descuento" type="number" step="0.01" value="0.00"></div>
        </div><br>
        <div class="actions"><button class="primary">Guardar pedido</button><button type="reset">Limpiar venta</button><a class="btn" href="{url_for('inventario')}">Nuevo prod. venta +</a><a class="btn" href="{url_for('ventas')}">Refrescar productos</a></div>
      </form>
      <form method="post" style="margin-top:12px" class="actions">
        <input type="hidden" name="accion" value="cobrar_ticket">
        <select name="pedido_id" style="max-width:520px">{opts_ped}</select>
        <select name="metodo_pago" style="max-width:220px"><option>EFECTIVO</option><option>YAPE</option><option>TARJETA</option><option>TRANSFERENCIA</option></select>
        <button class="btn-green">Cobrar y ticket</button>
      </form>
    </div>
    <div class="panel"><div class="box-title">Catálogo dinámico de productos para venta</div><br><div class="table-wrap small"><table><thead><tr><th>ID</th><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock</th><th>Estado</th></tr></thead><tbody>{tr_prod}</tbody></table></div></div>
    <div class="panel"><div class="box-title">Detalle actual / últimos ítems</div><br><div class="table-wrap small"><table><thead><tr><th>Pedido</th><th>Producto</th><th>Cantidad</th><th>Precio</th><th>Subtotal</th></tr></thead><tbody>{tr_det}</tbody></table></div></div>
    """
    return page(html, "ventas")

# =========================
# PEDIDOS
# =========================
@app.route("/pedidos", methods=["GET", "POST"])
@login_required
def pedidos():
    if request.method == "POST":
        accion = request.form.get("accion")
        pedido_id = int(request.form.get("pedido_id") or 0)
        if not pedido_id:
            flash("Selecciona un pedido.", "error")
            return redirect(url_for("pedidos"))
        if accion == "estado":
            estado = request.form.get("estado", "PENDIENTE")
            q_exec("UPDATE pedidos SET estado=? WHERE id=?", (estado, pedido_id))
            log_event("PEDIDO", f"Estado {estado} pedido {pedido_id}")
            flash("Estado actualizado.", "ok")
        elif accion == "pagado":
            crear_venta_desde_pedido(pedido_id, "EFECTIVO")
            flash("Pedido marcado como pagado.", "ok")
        elif accion == "limpiar":
            q_exec("DELETE FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
            q_exec("DELETE FROM pedidos WHERE id=?", (pedido_id,))
            flash("Pedido eliminado.", "ok")
        return redirect(url_for("pedidos"))

    estado = request.args.get("estado", "TODOS")
    if estado == "TODOS":
        rows = q_all("SELECT * FROM pedidos ORDER BY id DESC LIMIT 200")
    else:
        rows = q_all("SELECT * FROM pedidos WHERE estado=? ORDER BY id DESC", (estado,))
    opts_p = "".join(f'<option value="{r["id"]}">{r["codigo"]} - {r["cliente"]} - {r["estado"]}</option>' for r in rows)
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["codigo"]}</td><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["mesa"]}</td><td>{r["cliente"]}</td><td>{r["servicio"]}</td><td>{r["estado"]}</td><td>{money(r["total"])}</td><td>{r["pagado"]}</td></tr>' for r in rows) or '<tr><td colspan="10">Sin pedidos.</td></tr>'
    detalles = q_all("SELECT d.*,p.codigo FROM pedido_detalle d JOIN pedidos p ON p.id=d.pedido_id ORDER BY d.id DESC LIMIT 80")
    trd = "".join(f'<tr><td>{r["pedido_id"]}</td><td>{r["codigo"]}</td><td>{r["id"]}</td><td>{r["producto"]}</td><td>{r["cantidad"]}</td><td>{money(r["precio"])}</td><td>{money(r["total"])}</td></tr>' for r in detalles) or '<tr><td colspan="7">Sin detalle.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Control de pedidos</div><br>
      <form method="get" class="actions"><label>Estado</label><select name="estado" style="max-width:220px"><option>TODOS</option><option>PENDIENTE</option><option>PREPARACIÓN</option><option>LISTO</option><option>ENTREGADO</option><option>PAGADO</option></select><button>Refrescar</button></form><br>
      <form method="post" class="actions"><select name="pedido_id" style="max-width:520px">{opts_p}</select><input type="hidden" name="accion" value="estado"><select name="estado" style="max-width:220px"><option>PREPARACIÓN</option><option>LISTO</option><option>ENTREGADO</option><option>PAGADO</option></select><button>A estado</button></form><br>
      <form method="post" class="actions"><select name="pedido_id" style="max-width:520px">{opts_p}</select><button name="accion" value="pagado">Marcar pagado</button><button name="accion" value="limpiar" class="btn-red">Limpiar pedido</button><a class="btn" href="{url_for('ticket', pedido_id=0)}">Imprimir ticket</a></form>
    </div>
    <div class="panel"><div class="box-title">Listado de pedidos</div><br><div class="table-wrap"><table><thead><tr><th>ID</th><th>Código</th><th>Fecha</th><th>Hora</th><th>Mesa</th><th>Cliente</th><th>Servicio</th><th>Estado</th><th>Total</th><th>Pagado</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    <div class="panel"><div class="box-title">Detalle del pedido seleccionado</div><br><div class="table-wrap small"><table><thead><tr><th>Pedido ID</th><th>Código</th><th>Item</th><th>Producto</th><th>Cantidad</th><th>Precio</th><th>Subtotal</th></tr></thead><tbody>{trd}</tbody></table></div></div>
    """
    return page(html, "pedidos")

@app.route("/ticket/<int:pedido_id>")
@login_required
def ticket(pedido_id):
    p = q_one("SELECT * FROM pedidos WHERE id=?", (pedido_id,)) if pedido_id else q_one("SELECT * FROM pedidos ORDER BY id DESC LIMIT 1")
    if not p:
        return "Sin pedido"
    det = q_all("SELECT * FROM pedido_detalle WHERE pedido_id=?", (p["id"],))
    lines = ["RESTAURANTE AORIX", "TICKET", f"Pedido: {p['codigo']}", f"Fecha: {p['fecha']} {p['hora']}", f"Cliente: {p['cliente']}", "-" * 32]
    for d in det:
        lines.append(f"{d['cantidad']} x {d['producto']} {money(d['total'])}")
    lines += ["-" * 32, f"TOTAL: {money(p['total'])}"]
    bio = BytesIO("\n".join(lines).encode("utf-8"))
    return send_file(bio, as_attachment=True, download_name=f"ticket_{p['codigo']}.txt", mimetype="text/plain")

# =========================
# INVENTARIO
# =========================
@app.route("/inventario", methods=["GET", "POST"])
@login_required
def inventario():
    if request.method == "POST":
        accion = request.form.get("accion", "producto")
        if accion == "producto":
            nombre = up(request.form.get("nombre"))
            if not nombre:
                flash("Producto obligatorio.", "error")
                return redirect(url_for("inventario"))
            codigo = clean(request.form.get("codigo")) or "P" + now().strftime("%H%M%S")
            existe = q_one("SELECT id FROM productos WHERE nombre=?", (nombre,))
            data = (codigo, nombre, up(request.form.get("categoria") or "PLATOS"), request.form.get("tipo", "VENTA"), up(request.form.get("unidad") or "PLATO"), float(request.form.get("precio") or 0), float(request.form.get("costo") or 0), float(request.form.get("stock") or 0), float(request.form.get("stock_min") or 0))
            if existe:
                q_exec("UPDATE productos SET codigo=?,categoria=?,tipo=?,unidad=?,precio=?,costo=?,stock=?,stock_min=?,activo=1 WHERE nombre=?", data[:1]+data[2:]+(nombre,))
            else:
                q_exec("INSERT INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", data)
            flash("Producto guardado.", "ok")
        elif accion == "stock_in":
            pid = int(request.form.get("producto_id"))
            q_exec("UPDATE productos SET stock=stock+? WHERE id=?", (float(request.form.get("cantidad") or 0), pid))
            flash("Entrada de stock registrada.", "ok")
        elif accion == "stock_out":
            pid = int(request.form.get("producto_id"))
            q_exec("UPDATE productos SET stock=stock-? WHERE id=?", (float(request.form.get("cantidad") or 0), pid))
            flash("Salida de stock registrada.", "ok")
        elif accion == "importar" and "archivo" in request.files:
            file = request.files["archivo"]
            count = importar_productos(file)
            flash(f"Importación completada: {count} productos.", "ok")
        return redirect(url_for("inventario"))

    productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY CASE WHEN stock<=stock_min THEN 0 ELSE 1 END,nombre")
    opts_prod = select_options(productos)
    trp = "".join(f'<tr class="{"row-bad" if float(p["stock"] or 0)<=float(p["stock_min"] or 0) else ""}"><td>{p["id"]}</td><td>{p["nombre"]}</td><td>{p["categoria"]}</td><td>{p["tipo"]}</td><td>{p["unidad"]}</td><td>{money(p["precio"])}</td><td>{money(p["costo"])}</td><td>{p["stock"]}</td><td>{p["stock_min"]}</td><td>{"SIN STOCK" if float(p["stock"] or 0)<=0 else "OK"}</td></tr>' for p in productos)
    stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min")["c"]
    html = f"""
    <div class="panel"><div class="box-title">Productos e inventario</div><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="producto">
        <div><label>Código</label><input name="codigo" placeholder="Código"></div><div><label>Nombre producto</label><input name="nombre" placeholder="Nombre producto"></div>
        <div><label>Categoría</label><select name="categoria"><option>PLATOS</option><option>PARRILLAS</option><option>BEBIDAS</option><option>ADICIONALES</option><option>INSUMOS</option></select></div>
        <div><label>Tipo</label><select name="tipo"><option>VENTA</option><option>INSUMO</option></select></div><div><label>Unidad</label><select name="unidad"><option>PLATO</option><option>UND</option><option>KG</option><option>PORCION</option></select></div>
        <div><label>Precio venta</label><input name="precio" type="number" step="0.01"></div><div><label>Costo</label><input name="costo" type="number" step="0.01"></div><div><label>Stock</label><input name="stock" type="number" step="0.01"></div><div><label>Stock mínimo</label><input name="stock_min" type="number" step="0.01"></div><button>Guardar producto</button>
      </form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="stock_in"><select name="producto_id" style="max-width:380px">{opts_prod}</select><input name="cantidad" type="number" step="0.01" placeholder="Cantidad" style="max-width:160px"><button>Entrada stock</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="stock_out"><select name="producto_id" style="max-width:380px">{opts_prod}</select><input name="cantidad" type="number" step="0.01" placeholder="Cantidad" style="max-width:160px"><button>Salida stock</button><a class="btn" href="{url_for('plantilla_inventario')}">Descargar plantilla</a><a class="btn" href="{url_for('export_inventario')}">Exportar inventario</a><b style="color:#c2410c">Alertas: {stock_bajo} con stock bajo</b></form><br>
      <form method="post" enctype="multipart/form-data" class="actions"><input type="hidden" name="accion" value="importar"><input type="file" name="archivo" accept=".xlsx,.csv" style="max-width:340px"><button class="btn-orange">Importar Excel</button></form>
    </div>
    <div class="panel"><div class="box-title">Stock y productos</div><br><div class="table-wrap"><table><thead><tr><th>ID</th><th>Producto</th><th>Categoría</th><th>Tipo</th><th>Unidad</th><th>Precio</th><th>Costo</th><th>Stock</th><th>Mínimo</th><th>Estado</th></tr></thead><tbody>{trp}</tbody></table></div></div>
    """
    return page(html, "inventario")

def importar_productos(file_storage):
    count = 0
    filename = (file_storage.filename or "").lower()
    if filename.endswith(".xlsx") and OPENPYXL:
        wb = load_workbook(file_storage, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return 0
        headers = [up(h).replace(" ", "_") for h in rows[0]]
        for row in rows[1:]:
            d = dict(zip(headers, row))
            nombre = up(d.get("PRODUCTO") or d.get("NOMBRE") or "")
            if not nombre:
                continue
            codigo = clean(d.get("CODIGO") or d.get("CÓDIGO") or "") or "P" + str(count + 1).zfill(4)
            q_exec("INSERT OR REPLACE INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", (codigo, nombre, up(d.get("CATEGORIA") or "PLATOS"), up(d.get("TIPO") or "VENTA"), up(d.get("UNIDAD") or "UND"), float(d.get("PRECIO") or 0), float(d.get("COSTO") or 0), float(d.get("STOCK") or 0), float(d.get("MINIMO") or d.get("STOCK_MIN") or 0)))
            count += 1
    else:
        content = file_storage.read().decode("utf-8-sig", errors="ignore")
        for d in csv.DictReader(StringIO(content)):
            nombre = up(d.get("PRODUCTO") or d.get("NOMBRE") or "")
            if not nombre:
                continue
            codigo = clean(d.get("CODIGO") or "") or "P" + str(count + 1).zfill(4)
            q_exec("INSERT OR REPLACE INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", (codigo, nombre, up(d.get("CATEGORIA") or "PLATOS"), up(d.get("TIPO") or "VENTA"), up(d.get("UNIDAD") or "UND"), float(d.get("PRECIO") or 0), float(d.get("COSTO") or 0), float(d.get("STOCK") or 0), float(d.get("MINIMO") or 0)))
            count += 1
    return count

# =========================
# RECETAS
# =========================
@app.route("/recetas", methods=["GET", "POST"])
@login_required
def recetas():
    if request.method == "POST":
        accion = request.form.get("accion", "agregar")
        if accion == "insumo":
            nombre = up(request.form.get("nombre"))
            if nombre:
                q_exec("INSERT OR REPLACE INTO insumos(nombre,unidad,stock,stock_min,costo,activo) VALUES(?,?,?,?,?,1)", (nombre, up(request.form.get("unidad") or "UND"), float(request.form.get("stock") or 0), float(request.form.get("stock_min") or 0), float(request.form.get("costo") or 0)))
                flash("Insumo guardado.", "ok")
        else:
            q_exec("INSERT INTO recetas(producto_id,insumo_id,cantidad,observacion) VALUES(?,?,?,?)", (int(request.form.get("producto_id")), int(request.form.get("insumo_id")), float(request.form.get("cantidad") or 0), up(request.form.get("observacion"))))
            flash("Insumo agregado a receta.", "ok")
        return redirect(url_for("recetas"))
    productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY nombre")
    insumos = q_all("SELECT * FROM insumos WHERE activo=1 ORDER BY nombre")
    opts_p = select_options(productos)
    opts_i = select_options(insumos)
    rows = q_all("SELECT r.id,p.nombre producto,i.nombre insumo,i.unidad,r.cantidad,r.observacion FROM recetas r JOIN productos p ON p.id=r.producto_id JOIN insumos i ON i.id=r.insumo_id ORDER BY p.nombre,i.nombre")
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["producto"]}</td><td>{r["insumo"]}</td><td>{r["unidad"]}</td><td>{r["cantidad"]}</td><td>{r["observacion"]}</td></tr>' for r in rows) or '<tr><td colspan="6">Sin receta.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Recetas detalladas</div><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="agregar"><div><label>Producto de venta</label><select name="producto_id">{opts_p}</select></div><div><label>Insumo</label><select name="insumo_id">{opts_i}</select></div><div><label>Cantidad del insumo</label><input name="cantidad" type="number" step="0.0001" value="1"></div><div><label>Observación</label><input name="observacion"></div><button>Agregar insumo</button></form><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="insumo"><div><label>Nuevo insumo</label><input name="nombre" placeholder="Ej. AJI"></div><div><label>Unidad</label><select name="unidad"><option>PORCION</option><option>UND</option><option>KG</option></select></div><div><label>Stock</label><input name="stock" type="number" step="0.01"></div><div><label>Mínimo</label><input name="stock_min" type="number" step="0.01"></div><button>Guardar insumo</button></form>
    </div>
    <div class="panel"><div class="box-title">Detalle de receta</div><br><div class="table-wrap"><table><thead><tr><th>ID</th><th>Producto</th><th>Insumo</th><th>Unidad</th><th>Cantidad</th><th>Observación</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "recetas")

# =========================
# CAJA
# =========================
@app.route("/caja", methods=["GET", "POST"])
@login_required
def caja():
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "abrir":
            monto = float(request.form.get("monto_apertura") or 0)
            set_ctx("caja_abierta", "1")
            set_ctx("monto_apertura", monto)
            q_exec("INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario) VALUES(?,?,?,?,?,?)", (today(), hour(), "INGRESO", "APERTURA CAJA", monto, session.get("user")))
            flash("Caja abierta.", "ok")
        elif accion == "cerrar":
            set_ctx("caja_abierta", "0")
            flash("Caja cerrada.", "ok")
        elif accion == "gasto":
            q_exec("INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario) VALUES(?,?,?,?,?,?)", (today(), hour(), "EGRESO", up(request.form.get("concepto")), float(request.form.get("monto") or 0), session.get("user")))
            flash("Gasto registrado.", "ok")
        return redirect(url_for("caja"))
    f = request.args.get("fecha", today())
    rows = q_all("SELECT * FROM caja WHERE fecha=? ORDER BY id DESC", (f,))
    ingresos = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='INGRESO'", (f,))["t"]
    egresos = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='EGRESO'", (f,))["t"]
    trs = "".join(f'<tr><td>{r["hora"]}</td><td>{r["tipo"]}</td><td>{r["concepto"]}</td><td>{money(r["monto"])}</td><td>{r["usuario"]}</td></tr>' for r in rows) or '<tr><td colspan="5">Sin movimientos.</td></tr>'
    estado = "ABIERTA" if get_ctx("caja_abierta", "0") == "1" else "SIN ABRIR"
    html = f"""
    <div class="panel"><div class="box-title">Caja</div><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="abrir"><label>Monto apertura:</label><input name="monto_apertura" type="number" step="0.01" value="100.00" style="max-width:180px"><button>Abrir caja</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="cerrar"><button>Cerrar caja</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="gasto"><label>Egreso:</label><input name="concepto" placeholder="Concepto" style="max-width:240px"><input name="monto" type="number" step="0.01" value="0.00" style="max-width:180px"><button>Registrar gasto</button></form>
    </div>
    <div class="panel"><div class="box-title">Resumen de caja</div><h2 style="color:#dc2626">Caja: {estado}</h2><div class="grid"><div>Apertura: <b>{money(get_ctx('monto_apertura','0'))}</b></div><div>Efectivo sistema: <b>{money(ingresos-egresos)}</b></div><div>Ingresos: <b>{money(ingresos)}</b></div><div>Gastos: <b>{money(egresos)}</b></div></div></div>
    <div class="panel"><form method="get" class="actions"><input type="date" name="fecha" value="{f}" style="max-width:180px"><button>Filtrar</button></form><br><div class="table-wrap"><table><thead><tr><th>Hora</th><th>Tipo</th><th>Concepto</th><th>Monto</th><th>Usuario</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "caja")

# =========================
# DELIVERY / CLIENTES
# =========================
@app.route("/delivery", methods=["GET", "POST"])
@login_required
def delivery():
    if request.method == "POST":
        q_exec("INSERT INTO clientes(nombre,telefono,direccion,referencia,notas,activo) VALUES(?,?,?,?,?,1)", (up(request.form.get("nombre")), clean(request.form.get("telefono")), up(request.form.get("direccion")), up(request.form.get("referencia")), up(request.form.get("notas"))))
        flash("Cliente guardado.", "ok")
        return redirect(url_for("delivery"))
    rows = q_all("SELECT * FROM clientes WHERE activo=1 ORDER BY id DESC")
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["nombre"]}</td><td>{r["telefono"]}</td><td>{r["direccion"]}</td><td>{r["referencia"]}</td><td>{r["notas"]}</td></tr>' for r in rows) or '<tr><td colspan="6">Sin clientes.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Clientes delivery / frecuentes</div><br><form method="post" class="grid5"><div><label>Nombre</label><input name="nombre"></div><div><label>Teléfono</label><input name="telefono"></div><div><label>Dirección</label><input name="direccion"></div><div><label>Referencia</label><input name="referencia"></div><div><label>Notas</label><input name="notas"></div><button>Guardar cliente</button><a class="btn" href="{url_for('ventas')}">Cargar en venta</a><button type="reset">Llenar datos</button><a class="btn" href="{url_for('delivery')}">Refrescar</a></form></div>
    <div class="panel"><div class="box-title">Base de clientes</div><br><div class="table-wrap"><table><thead><tr><th>ID</th><th>Nombre</th><th>Teléfono</th><th>Dirección</th><th>Referencia</th><th>Notas</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "delivery")

# =========================
# INDICADORES
# =========================
@app.route("/indicadores")
@login_required
def indicadores():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    agrup = request.args.get("agrup", "DÍA")
    ventas = q_one("SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha BETWEEN ? AND ?", (fi, ff))
    pendientes = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha BETWEEN ? AND ? AND estado NOT IN ('PAGADO','ENTREGADO')", (fi, ff))["c"]
    stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min")["c"]
    ticket = (float(ventas["t"] or 0) / int(ventas["c"] or 1)) if int(ventas["c"] or 0) else 0
    rows = q_all("SELECT fecha periodo, COALESCE(SUM(total),0) ventas, COUNT(*) pedidos FROM ventas WHERE fecha BETWEEN ? AND ? GROUP BY fecha ORDER BY fecha", (fi, ff))
    maxv = max([float(r["ventas"] or 0) for r in rows] + [1])
    bars = "".join(f'<div class="bar-wrap"><div class="bar" style="height:{max(20, int((float(r["ventas"] or 0)/maxv)*300))}px">{money(r["ventas"])}</div><small>{r["periodo"]}</small></div>' for r in rows) or '<div class="muted">Sin ventas.</div>'
    trs = "".join(f'<tr><td>{r["periodo"]}</td><td>{money(r["ventas"])}</td><td>{r["pedidos"]}</td><td>{money(float(r["ventas"] or 0)/int(r["pedidos"] or 1))}</td></tr>' for r in rows) or '<tr><td colspan="4">Sin detalle.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Filtros de indicadores</div><br><form method="get" class="actions"><label>Agrupar por:</label><select name="agrup" style="max-width:150px"><option>DÍA</option><option>MES</option></select><label>Fecha inicio:</label><input type="date" name="fi" value="{fi}" style="max-width:170px"><label>Fecha fin:</label><input type="date" name="ff" value="{ff}" style="max-width:170px"><label>Gráfica:</label><select style="max-width:180px"><option>VENTAS S/</option></select><button>Actualizar indicadores</button><a class="btn" href="{url_for('indicadores')}">Hoy</a></form></div>
    <div class="kpis"><div class="kpi"><h3>Ventas netas</h3><b>{money(ventas['t'])}</b></div><div class="kpi"><h3>Pedidos pagados</h3><b>{ventas['c']}</b></div><div class="kpi"><h3>Ticket promedio</h3><b>{money(ticket)}</b></div><div class="kpi"><h3>Pedidos pendientes</h3><b>{pendientes}</b></div></div>
    <div class="grid2"><div class="panel"><div class="box-title">Gráfica comparativa</div><div class="chart">{bars}</div></div><div class="panel"><div class="box-title">Detalle por periodo</div><br><div class="table-wrap"><table><thead><tr><th>Periodo</th><th>Ventas S/</th><th>Pedidos</th><th>Ticket promedio</th></tr></thead><tbody>{trs}</tbody></table></div></div></div>
    """
    return page(html, "indicadores")

# =========================
# REPORTES
# =========================
@app.route("/reportes")
@login_required
def reportes():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    ventas = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    stock = q_all("SELECT * FROM productos WHERE activo=1 AND stock<=stock_min ORDER BY nombre")
    lines = ["REPORTE RESTAURANTE AORIX", f"PERIODO: {fi} al {ff}", "", "VENTAS", "-"*70]
    total = 0
    for v in ventas:
        total += float(v["total"] or 0)
        lines.append(f"{v['fecha']} {v['hora']} | {v['cliente']:<25} | {v['metodo_pago']:<12} | {money(v['total'])}")
    lines += ["-"*70, f"TOTAL VENTAS: {money(total)}", "", "STOCK BAJO", "-"*70]
    for s in stock:
        lines.append(f"{s['nombre']:<30} | Stock: {s['stock']} {s['unidad']} | Mínimo: {s['stock_min']}")
    reporte = "\n".join(lines)
    html = f"""
    <div class="panel"><div class="box-title">Filtros</div><br><form method="get" class="actions"><label>Fecha inicio:</label><input type="date" name="fi" value="{fi}" style="max-width:170px"><label>Fecha fin:</label><input type="date" name="ff" value="{ff}" style="max-width:170px"><button>Generar</button><a class="btn" href="{url_for('export_excel',fi=fi,ff=ff)}">Exportar Excel</a><a class="btn" href="{url_for('export_csv',fi=fi,ff=ff)}">Exportar CSV</a></form></div>
    <div class="panel"><div class="box-title">Reporte</div><textarea class="report-box" readonly>{reporte}</textarea></div>
    """
    return page(html, "reportes")

@app.route("/export_excel")
@login_required
def export_excel():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    rows = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    if OPENPYXL:
        wb = Workbook()
        ws = wb.active
        ws.title = "Ventas"
        ws.append(["ID", "Fecha", "Hora", "Cliente", "Servicio", "Pago", "Subtotal", "Descuento", "Total", "Usuario", "Estado"])
        for r in rows:
            ws.append([r["id"], r["fecha"], r["hora"], r["cliente"], r["servicio"], r["metodo_pago"], r["subtotal"], r["descuento"], r["total"], r["usuario"], r["estado"]])
        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return send_file(bio, as_attachment=True, download_name=f"reporte_aorix_{fi}_{ff}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return redirect(url_for("export_csv", fi=fi, ff=ff))

@app.route("/export_csv")
@login_required
def export_csv():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    rows = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["ID", "Fecha", "Hora", "Cliente", "Servicio", "Pago", "Total"])
    for r in rows:
        w.writerow([r["id"], r["fecha"], r["hora"], r["cliente"], r["servicio"], r["metodo_pago"], r["total"]])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name=f"reporte_aorix_{fi}_{ff}.csv", mimetype="text/csv")

@app.route("/export_inventario")
@login_required
def export_inventario():
    rows = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY nombre")
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["CODIGO", "PRODUCTO", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "MINIMO"])
    for r in rows:
        w.writerow([r["codigo"], r["nombre"], r["categoria"], r["tipo"], r["unidad"], r["precio"], r["costo"], r["stock"], r["stock_min"]])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name="inventario_aorix.csv", mimetype="text/csv")

@app.route("/plantilla_inventario")
@login_required
def plantilla_inventario():
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["CODIGO", "PRODUCTO", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "MINIMO"])
    w.writerow(["P001", "1/2 POLLO", "PLATOS", "VENTA", "PLATO", 35, 10, 40, 5])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name="plantilla_inventario_aorix.csv", mimetype="text/csv")

# =========================
# ADMIN / LOG
# =========================
@app.route("/admin", methods=["GET", "POST"])
@login_required
@admin_required
def admin():
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "sucursal":
            nombre = up(request.form.get("nombre"))
            if nombre:
                q_exec("INSERT OR IGNORE INTO sucursales(nombre,activo) VALUES(?,1)", (nombre,))
                flash("Sucursal registrada.", "ok")
        elif accion == "usuario":
            usuario = clean(request.form.get("usuario"))
            clave = clean(request.form.get("clave"))
            nombre = up(request.form.get("nombre"))
            rol = request.form.get("rol", "MESERO")
            if usuario and clave:
                q_exec("INSERT OR REPLACE INTO usuarios(usuario,nombre,clave_hash,rol,activo) VALUES(?,?,?,?,1)", (usuario, nombre, generate_password_hash(clave), rol))
                flash("Usuario guardado.", "ok")
        elif accion == "sembrar":
            init_db()
            flash("Demo sembrada / verificada.", "ok")
        return redirect(url_for("admin"))
    sucursales = q_all("SELECT * FROM sucursales ORDER BY nombre")
    usuarios = q_all("SELECT usuario,nombre,rol,activo FROM usuarios ORDER BY usuario")
    tr_s = "".join(f'<tr><td>{s["id"]}</td><td>{s["nombre"]}</td><td>{s["activo"]}</td></tr>' for s in sucursales)
    tr_u = "".join(f'<tr><td>{u["usuario"]}</td><td>{u["nombre"]}</td><td>{u["rol"]}</td><td>{u["activo"]}</td></tr>' for u in usuarios)
    html = f"""
    <div class="panel"><div class="box-title">Administrador</div><br><form method="post" class="actions"><input type="hidden" name="accion" value="negocio"><label>Nombre negocio / sucursal:</label><select style="max-width:300px">{select_options(sucursales,'nombre','nombre')}</select><button>Guardar negocio</button><button name="accion" value="sembrar">Sembrar demo</button></form></div>
    <div class="panel"><div class="box-title">Registrar sucursal</div><br><form method="post" class="actions"><input type="hidden" name="accion" value="sucursal"><label>Nueva sucursal:</label><input name="nombre" style="max-width:280px"><button>Registrar sucursal</button></form></div>
    <div class="panel"><div class="box-title">Registrar usuario</div><br><form method="post" class="grid5"><input type="hidden" name="accion" value="usuario"><div><label>Usuario</label><input name="usuario"></div><div><label>Nombre</label><input name="nombre"></div><div><label>Clave</label><input name="clave"></div><div><label>Rol</label><select name="rol"><option>ADMIN</option><option>MESERO</option><option>CAJA</option><option>COCINA</option></select></div><button>Guardar usuario</button></form></div>
    <div class="grid2"><div class="panel"><div class="box-title">Sucursales</div><div class="table-wrap small"><table><thead><tr><th>ID</th><th>Sucursal</th><th>Activo</th></tr></thead><tbody>{tr_s}</tbody></table></div></div><div class="panel"><div class="box-title">Usuarios</div><div class="table-wrap small"><table><thead><tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Activo</th></tr></thead><tbody>{tr_u}</tbody></table></div></div></div>
    <div class="panel"><div class="box-title">Resumen funcional</div><textarea class="report-box" readonly>VERSION WEB PRO AORIX\n\nMEJORAS ACTIVAS:\n- Sucursales registrables\n- Usuarios registrables\n- Pedido unificado para mismo cliente / mesa / servicio\n- Detalle de ítems en pestaña Pedidos\n- Caja, inventario, recetas, delivery, indicadores, reportes y log\n- Listo para Render con SQLite persistente en /data si está disponible\n\nBase actual: {DB_PATH}</textarea></div>
    """
    return page(html, "admin")

@app.route("/logs")
@login_required
def logs():
    rows = q_all("SELECT * FROM logs ORDER BY id DESC LIMIT 300")
    trs = "".join(f'<tr><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["usuario"]}</td><td>{r["accion"]}</td><td>{r["detalle"]}</td></tr>' for r in rows) or '<tr><td colspan="5">Sin log.</td></tr>'
    return page(f'<div class="panel"><div class="box-title">Log del sistema</div><br><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Usuario</th><th>Acción</th><th>Detalle</th></tr></thead><tbody>{trs}</tbody></table></div></div>', "log")

@app.errorhandler(500)
def err500(e):
    try:
        log_event("ERROR 500", str(e))
    except Exception:
        pass
    return page('<div class="panel"><h2>Error interno controlado</h2><p>No se perdió información. Revisa los logs de Render o vuelve al panel.</p><a class="btn btn-primary" href="/dashboard">Volver al panel</a></div>'), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=False)
