# -*- coding: utf-8 -*-
"""
EL TORO Restaurant Grill - POS PRO FINAL
Render / Web / Celular
Usuarios demo:
- admin / admin123
- vendedor1 / vendedor1
"""
import os, sqlite3, csv, io, base64
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, redirect, url_for, session, flash, send_file, jsonify, render_template_string
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
try:
    import pandas as pd
except Exception:
    pd = None
try:
    import qrcode
except Exception:
    qrcode = None

BASE_DIR=os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR=os.getenv('PERSIST_DIR', '/data' if os.path.isdir('/data') else BASE_DIR)
STATIC_DIR=os.path.join(BASE_DIR,'static')
UPLOAD_DIR=os.path.join(STATIC_DIR,'uploads')
DATA_DIR=os.path.join(PERSIST_DIR,'data')
REPORT_DIR=os.path.join(PERSIST_DIR,'reportes')
for d in (STATIC_DIR,UPLOAD_DIR,DATA_DIR,REPORT_DIR): os.makedirs(d,exist_ok=True)
DB=os.path.join(DATA_DIR,'el_toro_pos.db')
app=Flask(__name__, static_folder='static')
app.secret_key=os.getenv('SECRET_KEY','el-toro-final-pro-2026')

# ---------------- DB ----------------
def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def q(sql,p=(),one=False):
    with conn() as c:
        cur=c.execute(sql,p); rows=cur.fetchall();
    return rows[0] if one and rows else (None if one else rows)

def x(sql,p=()):
    with conn() as c:
        cur=c.execute(sql,p); c.commit(); return cur.lastrowid

def init_db():
    with conn() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS sucursales(id INTEGER PRIMARY KEY AUTOINCREMENT,nombre TEXT UNIQUE, activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS usuarios(id INTEGER PRIMARY KEY AUTOINCREMENT,usuario TEXT UNIQUE,nombre TEXT,clave_hash TEXT,clave_plain TEXT,rol TEXT,sucursal_id INTEGER,activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS productos(id INTEGER PRIMARY KEY AUTOINCREMENT,sku TEXT UNIQUE,nombre TEXT,categoria TEXT,precio INTEGER DEFAULT 0,costo INTEGER DEFAULT 0,stock INTEGER DEFAULT 0,stock_min INTEGER DEFAULT 0,imagen TEXT,activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS clientes(id INTEGER PRIMARY KEY AUTOINCREMENT,nombre TEXT,telefono TEXT,direccion TEXT,referencia TEXT,notas TEXT);
        CREATE TABLE IF NOT EXISTS pedidos(id INTEGER PRIMARY KEY AUTOINCREMENT,codigo TEXT,fecha TEXT,hora TEXT,mesa TEXT,servicio TEXT,cliente TEXT,telefono TEXT,direccion TEXT,referencia TEXT,total INTEGER,descuento INTEGER DEFAULT 0,metodo TEXT,estado TEXT DEFAULT 'PENDIENTE',pagado INTEGER DEFAULT 0,sucursal_id INTEGER,usuario TEXT);
        CREATE TABLE IF NOT EXISTS pedido_items(id INTEGER PRIMARY KEY AUTOINCREMENT,pedido_id INTEGER,producto_id INTEGER,sku TEXT,producto TEXT,cantidad INTEGER,precio INTEGER,subtotal INTEGER);
        CREATE TABLE IF NOT EXISTS cierres(id INTEGER PRIMARY KEY AUTOINCREMENT,fecha TEXT,sucursal_id INTEGER,usuario TEXT,total_ventas INTEGER,total_pedidos INTEGER,cerrado_en TEXT);
        CREATE TABLE IF NOT EXISTS catalogo_cfg(id INTEGER PRIMARY KEY CHECK(id=1),nombre TEXT,slug TEXT,whatsapp TEXT,descripcion TEXT);
        ''')
        c.commit()
    if not q('SELECT id FROM sucursales WHERE nombre=?',('Sucursal Principal',),True):
        x('INSERT INTO sucursales(nombre) VALUES(?)',('Sucursal Principal',))
    sid=q('SELECT id FROM sucursales WHERE nombre=?',('Sucursal Principal',),True)['id']
    for u,n,cl,r in [('admin','Administrador','admin123','ADMIN'),('vendedor1','Vendedor 1','vendedor1','VENDEDOR')]:
        if not q('SELECT id FROM usuarios WHERE usuario=?',(u,),True):
            x('INSERT INTO usuarios(usuario,nombre,clave_hash,clave_plain,rol,sucursal_id) VALUES(?,?,?,?,?,?)',(u,n,generate_password_hash(cl),cl,r,sid))
    demos=[('P001','1/2 POLLO','PARRILLAS',35,18,40,5),('P002','1/4 POLLO','PARRILLAS',18,9,25,5),('P003','PIZZA MUZZARELLA','PIZZAS',32,14,30,5),('P004','PIZZA AMERICANA','PIZZAS',38,16,30,5),('B001','GASEOSA 500ML','BEBIDAS',5,3,60,10),('B002','CHICHA MORADA','BEBIDAS',6,2,50,10)]
    for d in demos:
        if not q('SELECT id FROM productos WHERE sku=?',(d[0],),True): x('INSERT INTO productos(sku,nombre,categoria,precio,costo,stock,stock_min) VALUES(?,?,?,?,?,?,?)',d)
    if not q('SELECT id FROM catalogo_cfg WHERE id=1',(),True):
        x('INSERT INTO catalogo_cfg(id,nombre,slug,whatsapp,descripcion) VALUES(1,?,?,?,?)',('EL TORO Restaurant Grill','el-toro','51999999999','Parrillas, pizzas y delivery'))
init_db()

# ---------------- helpers ----------------
def current_user(): return q('SELECT u.*,s.nombre sucursal FROM usuarios u LEFT JOIN sucursales s ON s.id=u.sucursal_id WHERE usuario=?',(session.get('user',''),),True)
def login_required(f):
    @wraps(f)
    def w(*a,**k):
        if not session.get('user'): return redirect(url_for('login'))
        return f(*a,**k)
    return w
def admin_required(f):
    @wraps(f)
    def w(*a,**k):
        if session.get('rol')!='ADMIN': flash('Solo administrador.','err'); return redirect(url_for('ventas'))
        return f(*a,**k)
    return w
def money(n): return 'S/ {:,.0f}'.format(int(n or 0))
def today(): return date.today().isoformat()
def nowtime(): return datetime.now().strftime('%H:%M:%S')
def sku_ok(s):
    import re
    return bool(re.match(r'^[A-Za-z0-9\-]{2,20}$', s or ''))

def save_qr_pago(codigo,total,metodo='YAPE/PLIN'):
    data=f'EL TORO|{metodo}|PEDIDO:{codigo}|MONTO:{int(total)}'
    path=os.path.join(STATIC_DIR,'qr_pago.png')
    if qrcode:
        img=qrcode.make(data); img.save(path)
    else:
        # tiny fallback svg-like text file not used as image
        pass
    return '/static/qr_pago.png'

# ---------------- UI ----------------
CSS='''
:root{--side:250px;--red:#ff002b;--red2:#b40000;--orange:#ff751f;--dark:#070812;--panel:#0e101b;--text:#101827;--muted:#667085;--line:#e8edf5;--bg:#f3f6fb;--card:#fff;--green:#20c997}
*{box-sizing:border-box} html,body{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:radial-gradient(circle at top right,#ffe1e8 0,#f3f6fb 34%,#eef2f7 100%);color:var(--text);overflow-x:hidden} a{text-decoration:none;color:inherit}.hidden{display:none!important}.muted{color:var(--muted)}
.btn,button{border:0;border-radius:18px;padding:14px 20px;font-weight:950;cursor:pointer;background:#eef2f7;color:#101827;box-shadow:0 12px 28px rgba(16,24,40,.08)}.btn-red{background:linear-gradient(135deg,var(--red),#ff6b45);color:white;box-shadow:0 18px 36px rgba(255,0,43,.28)}.btn-dark{background:#0b1020;color:#fff}.btn-green{background:linear-gradient(135deg,#15c47e,#19a9ff);color:#fff}.btn-sm{padding:9px 12px;border-radius:12px;font-size:13px}input,select,textarea{width:100%;border:1px solid #dbe4f0;border-radius:18px;padding:15px 18px;background:white;font-size:16px;outline:none}input:focus,select:focus{border-color:var(--red);box-shadow:0 0 0 4px rgba(255,0,43,.12)}label{display:block;font-weight:950;margin:0 0 7px;color:#243044}.flash{padding:13px 16px;border-radius:16px;background:#fff3f3;border:1px solid #ffc7cf;color:#800;font-weight:900;margin-bottom:14px}.ok{background:#effff7;border-color:#b7f7d6;color:#05603a}.err{background:#fff1f2;color:#9f1239}
.login-page{min-height:100vh;display:grid;place-items:center;background:radial-gradient(circle at 15% 0,rgba(255,0,43,.22),transparent 35%),linear-gradient(160deg,#070812,#080312 65%,#230006);padding:20px}.login-card{width:min(560px,94vw);background:rgba(12,13,28,.92);border:1px solid rgba(255,255,255,.12);border-radius:36px;padding:44px 38px;color:white;box-shadow:0 40px 90px rgba(0,0,0,.55),0 0 50px rgba(255,0,43,.12);text-align:center}.logo-img{width:170px;max-width:70%;height:auto;object-fit:contain;display:block;margin:0 auto 18px}.login-card h1{font-size:42px;margin:0;color:white}.login-card p{font-size:18px;color:#c8cfdd;font-weight:800}.login-card label{text-align:left;color:#fff}.login-card input{background:#1b1e2f;color:white;border-color:#31364c}.promo{margin-top:34px;text-align:left;background:linear-gradient(135deg,#ff0038,#ff6d4c);border-radius:24px;padding:28px;color:white;box-shadow:0 22px 45px rgba(255,0,43,.35)}.promo b{font-size:30px;display:block}
.sidebar{position:fixed;left:0;top:0;width:var(--side);height:100vh;background:radial-gradient(circle at 0 0,rgba(255,0,43,.18),transparent 35%),linear-gradient(180deg,#090b17,#08010a 70%,#001b14);color:white;z-index:50;overflow-y:auto;border-right:1px solid rgba(255,255,255,.12)}.brand{padding:24px 18px;text-align:center;border-bottom:1px solid rgba(255,255,255,.1)}.brand .logo-img{width:150px;margin-bottom:8px}.brand-title{font-weight:1000;font-size:20px}.brand-sub{font-weight:800;color:#c9d2df;font-size:13px}.user-mini{font-weight:950;color:#d9e1ef;margin-top:18px}.nav{padding:18px 12px 90px;display:flex;flex-direction:column;gap:9px}.nav a{display:flex;align-items:center;gap:12px;padding:15px 18px;border-radius:18px;font-weight:950;color:#eef2ff}.nav a:hover,.nav a.on{background:linear-gradient(135deg,var(--red),#ff7258);box-shadow:0 16px 32px rgba(255,0,43,.32)}.main{margin-left:var(--side);min-height:100vh}.top{position:sticky;top:0;z-index:30;background:rgba(255,255,255,.82);backdrop-filter:blur(16px);border-bottom:1px solid rgba(228,233,242,.8);padding:18px 28px;display:flex;align-items:center;justify-content:space-between}.top h1{margin:0;font-size:30px}.pill{background:#101827;color:#fff;border-radius:999px;padding:10px 14px;font-weight:900}.content{padding:28px;max-width:1650px}.tabs{display:none}.card{background:rgba(255,255,255,.90);border:1px solid rgba(219,228,240,.9);border-radius:30px;padding:26px;box-shadow:0 28px 55px rgba(16,24,40,.08);margin-bottom:24px}.card h2{margin:0 0 18px;font-size:28px}.grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:16px}.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:20px}.actions{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:22px;background:#fff}.table-wrap table{width:100%;border-collapse:collapse;min-width:760px}th,td{padding:15px;border-bottom:1px solid #e6edf5;text-align:left;white-space:nowrap}th{background:#071827;color:white;font-weight:950}.kpi{background:linear-gradient(135deg,#fff,#fff5f5);border:1px solid #ffe1e1;border-radius:26px;padding:24px}.kpi b{font-size:32px;display:block;margin-top:8px}.products{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:16px}.prod{border:1px solid #edf1f7;border-radius:24px;padding:15px;background:white;box-shadow:0 15px 30px rgba(16,24,40,.06)}.prod img{width:100%;height:110px;object-fit:cover;border-radius:18px;background:#f1f5f9}.prod h3{font-size:16px;margin:10px 0 5px}.cart{position:sticky;top:100px}.pay-box{display:grid;grid-template-columns:1fr 260px;gap:20px}.qr{width:220px;max-width:100%;background:#fff;border-radius:20px;padding:12px;border:1px solid #ddd}.mobile-nav{display:none}.only-admin{}.quick-note{background:#fff5f5;border:1px solid #ffc9d0;color:#7f0000;border-radius:20px;padding:16px;font-weight:900}.sku-help{font-size:13px;background:#fff5f5;border:1px solid #ffc9d0;color:#8a0000;border-radius:16px;padding:12px;font-weight:900}.catalog-card{background:#0d1020;color:white;border-radius:28px;padding:22px}.catalog-card img{width:100%;height:160px;object-fit:cover;border-radius:22px}.bottom-space{height:20px}
@media(max-width:900px){:root{--side:0}.sidebar{display:none}.main{margin-left:0}.top{padding:12px 14px}.top h1{font-size:20px}.content{padding:12px 10px 90px}.card{padding:18px;border-radius:24px}.grid,.grid2,.grid3{grid-template-columns:1fr}.pay-box{grid-template-columns:1fr}.mobile-nav{position:fixed;left:0;right:0;bottom:0;height:68px;background:rgba(5,5,9,.96);display:flex;justify-content:space-around;align-items:center;z-index:999;border-top:1px solid rgba(255,255,255,.1);backdrop-filter:blur(14px)}.mobile-nav a{color:white;font-weight:900;font-size:12px;display:grid;place-items:center;gap:3px}.mobile-nav a.on{color:#ff4b5f}.mobile-nav span{font-size:20px}.tabs{display:none!important}.products{grid-template-columns:repeat(2,1fr)}.prod img{height:90px}.login-card{padding:30px 22px}.login-card h1{font-size:32px}.promo{padding:22px}.nav{display:none}.table-wrap table{min-width:640px}input,select{font-size:15px}.cart{position:relative;top:auto}}
'''

def nav(active):
    u=current_user(); rol=session.get('rol')
    items=[('dashboard','📊','Panel principal'),('ventas','🧾','Venta'),('pedidos','🚚','Pedido'),('cierre','🔒','Cierre'),('inventario','📦','Inventario'),('catalogo','🖼️','Catálogo / QR'),('recetas','🍽️','Recetas'),('caja','💵','Caja'),('delivery','🛵','Delivery'),('indicadores','📈','Indicadores'),('reportes','📄','Reportes')]
    if rol=='ADMIN': items.append(('usuarios','⚙️','Usuarios / Admin'))
    items.append(('logout','🚪','Salir'))
    links=''.join([f'<a class="{"on" if active==r else ""}" href="/{r}"><span>{ico}</span>{txt}</a>' for r,ico,txt in items if (rol=='ADMIN' or r in ['ventas','pedidos','catalogo','cierre','logout'])])
    mob=''.join([f'<a class="{"on" if active==r else ""}" href="/{r}"><span>{ico}</span>{txt.split()[0]}</a>' for r,ico,txt in items if r in (['ventas','pedidos','catalogo','cierre','logout'] if rol!='ADMIN' else ['dashboard','ventas','pedidos','catalogo','reportes'])])
    return f'''<aside class="sidebar"><div class="brand"><img class="logo-img" src="/static/logo_toro.jpg"><div class="brand-title">EL TORO</div><div class="brand-sub">Restaurant Grill</div><div class="user-mini">{session.get('user')} · {rol}</div></div><nav class="nav">{links}</nav></aside><div class="mobile-nav">{mob}</div>'''

def page(title, active, body):
    u=current_user()
    return render_template_string(f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>EL TORO POS</title><style>{CSS}</style></head><body>{nav(active)}<main class="main"><header class="top"><h1>{title}</h1><div class="actions"><span class="pill">{(u['sucursal'] if u else '')}</span><span class="pill">{session.get('rol','')}</span></div></header><section class="content">{{% with messages=get_flashed_messages(with_categories=true) %}}{{% for c,m in messages %}}<div class="flash {{{{c}}}}">{{{{m}}}}</div>{{% endfor %}}{{% endwith %}}{body}<div class="bottom-space"></div></section></main><script>function tog(id){{let e=document.getElementById(id);e.type=e.type==='password'?'text':'password'}};function printPage(){{window.print()}}</script></body></html>''')

@app.route('/',methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=q('SELECT * FROM usuarios WHERE usuario=? AND activo=1',(request.form.get('usuario','').strip(),),True)
        if u and check_password_hash(u['clave_hash'],request.form.get('clave','')):
            session['user']=u['usuario']; session['rol']=u['rol']; session['sucursal_id']=u['sucursal_id']; return redirect(url_for('dashboard' if u['rol']=='ADMIN' else 'ventas'))
        flash('Usuario o clave incorrecta','err')
    return render_template_string(f'''<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>{CSS}</style></head><body><div class="login-page"><form class="login-card" method="post"><img class="logo-img" src="/static/logo_toro.jpg"><h1>Restaurant Grill</h1><p>Restaurante · Pizzería · Parrillas · Delivery · Caja</p><label>Usuario</label><input name="usuario" placeholder="Ingrese su usuario" autofocus><br><br><label>Clave</label><input name="clave" type="password" placeholder="Ingrese su clave"><br><br><button class="btn-red" style="width:100%">Ingresar</button><div class="promo"><b>15% descuento</b><strong>primera compra · catálogo digital</strong></div></form></div></body></html>''')

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    sid=session.get('sucursal_id'); fecha=today()
    ventas=q('SELECT COUNT(*) c, COALESCE(SUM(total),0) t FROM pedidos WHERE fecha=? AND sucursal_id=?',(fecha,sid),True)
    stock=q('SELECT COUNT(*) c FROM productos WHERE stock<=stock_min AND activo=1',(),True)['c']
    body=f'''<div class="grid"><div class="kpi">Ventas hoy<b>{money(ventas['t'])}</b></div><div class="kpi">Pedidos hoy<b>{ventas['c']}</b></div><div class="kpi">Stock bajo<b>{stock}</b></div><div class="kpi">Modo<b>POS</b></div></div><div class="card"><h2>🚀 Accesos rápidos</h2><div class="actions"><a class="btn btn-red" href="/ventas">Nueva venta</a><a class="btn" href="/pedidos">Ver pedidos</a><a class="btn" href="/catalogo">Catálogo QR</a><a class="btn" href="/indicadores">Indicadores</a></div></div>'''
    return page('EL TORO RESTAURANT GRILL','dashboard',body)

@app.route('/ventas',methods=['GET','POST'])
@login_required
def ventas():
    if request.method=='POST':
        prod_ids=request.form.getlist('producto_id'); cants=request.form.getlist('cantidad')
        cliente=request.form.get('cliente') or 'Cliente general'; servicio=request.form.get('servicio','SALÓN'); metodo=request.form.get('metodo','EFECTIVO'); descuento=int(float(request.form.get('descuento') or 0))
        total=0; items=[]
        for pid,can in zip(prod_ids,cants):
            try: can=int(can)
            except: can=0
            if can<=0: continue
            p=q('SELECT * FROM productos WHERE id=?',(pid,),True)
            if p:
                sub=int(p['precio'])*can; total+=sub; items.append((p,can,sub))
        total=max(0,total-descuento)
        if not items: flash('Agrega por lo menos un producto.','err'); return redirect(url_for('ventas'))
        codigo='PED-'+datetime.now().strftime('%Y%m%d-%H%M%S')
        pedido_id=x('INSERT INTO pedidos(codigo,fecha,hora,mesa,servicio,cliente,telefono,direccion,referencia,total,descuento,metodo,sucursal_id,usuario) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(codigo,today(),nowtime(),request.form.get('mesa',''),servicio,cliente,request.form.get('telefono',''),request.form.get('direccion',''),request.form.get('referencia',''),total,descuento,metodo,session.get('sucursal_id'),session.get('user')))
        for p,can,sub in items:
            x('INSERT INTO pedido_items(pedido_id,producto_id,sku,producto,cantidad,precio,subtotal) VALUES(?,?,?,?,?,?,?)',(pedido_id,p['id'],p['sku'],p['nombre'],can,p['precio'],sub))
            x('UPDATE productos SET stock=MAX(stock-?,0) WHERE id=?',(can,p['id']))
        qr=save_qr_pago(codigo,total,metodo)
        flash(f'Pedido guardado {codigo}. Total {money(total)}. QR generado.','ok')
        return redirect(url_for('pagar',pedido_id=pedido_id))
    buscar=request.args.get('buscar','')
    where='WHERE activo=1'; params=[]
    if buscar: where+=' AND (nombre LIKE ? OR sku LIKE ? OR categoria LIKE ?)'; params+=['%'+buscar+'%']*3
    prods=q(f'SELECT * FROM productos {where} ORDER BY categoria,nombre',params)
    cards=''.join([f'''<div class="prod"><img src="{('/static/uploads/'+p['imagen']) if p['imagen'] else '/static/logo_toro.jpg'}"><h3>{p['nombre']}</h3><b>{money(p['precio'])}</b><p class="muted">Stock {int(p['stock'])}</p><label>Cantidad</label><input type="number" name="cantidad" value="0" min="0" step="1"><input type="hidden" name="producto_id" value="{p['id']}"></div>''' for p in prods])
    body=f'''<form method="get" class="card"><h2>🔍 Buscar producto</h2><div class="actions"><input style="max-width:520px" name="buscar" value="{buscar}" placeholder="Buscar producto por nombre, código o categoría"><button>Buscar</button><a class="btn" href="/ventas">Limpiar</a></div></form><form method="post"><div class="grid2"><div class="card"><h2>🧾 Nueva venta / POS rápido</h2><div class="grid"><div><label>Mesa</label><input name="mesa" placeholder="Mesa 1"></div><div><label>Servicio</label><select name="servicio"><option>SALÓN</option><option>DELIVERY</option><option>RECOJO</option></select></div><div><label>Cliente</label><input name="cliente" placeholder="Cliente general"></div><div><label>Teléfono</label><input name="telefono" inputmode="numeric"></div><div><label>Dirección</label><input name="direccion"></div><div><label>Referencia</label><input name="referencia"></div><div><label>Método pago</label><select name="metodo"><option>EFECTIVO</option><option>YAPE</option><option>PLIN</option><option>TARJETA</option><option>TRANSFERENCIA</option></select></div><div><label>Descuento entero</label><input name="descuento" type="number" step="1" value="0"></div></div><br><button class="btn-red">Guardar y generar QR de pago</button></div><div class="card cart"><h2>📦 Catálogo de productos</h2><div class="products">{cards or '<p>No hay productos.</p>'}</div></div></div></form>{'' if session.get('rol')!='ADMIN' else carga_dia_html()}'''
    return page('Ventas / POS rápido','ventas',body)

def carga_dia_html():
    return '''<div class="card only-admin"><h2>📥 Carga de inicio de día</h2><div class="quick-note">Solo administrador: importa Excel/CSV para actualizar productos, precios y stock antes de vender.</div><br><form action="/importar_productos" method="post" enctype="multipart/form-data" class="actions"><input type="file" name="archivo" accept=".xlsx,.csv" required><button class="btn-red">Cargar día / importar Excel productos</button><a class="btn" href="/plantilla_productos">Descargar plantilla</a></form></div>'''

@app.route('/pagar/<int:pedido_id>')
@login_required
def pagar(pedido_id):
    p=q('SELECT * FROM pedidos WHERE id=?',(pedido_id,),True); items=q('SELECT * FROM pedido_items WHERE pedido_id=?',(pedido_id,))
    qr=save_qr_pago(p['codigo'],p['total'],p['metodo'])
    rows=''.join([f'<tr><td>{i["producto"]}</td><td>{i["cantidad"]}</td><td>{money(i["subtotal"])}</td></tr>' for i in items])
    body=f'''<div class="card"><h2>💳 Cobro QR</h2><div class="pay-box"><div><h3>{p['codigo']}</h3><p>Cliente: <b>{p['cliente']}</b></p><p>Método: <b>{p['metodo']}</b></p><h1>Total: {money(p['total'])}</h1><div class="actions"><a class="btn btn-red" href="/marcar_pagado/{p['id']}">Confirmar pago</a><button onclick="printPage()" type="button">Imprimir ticket</button><a class="btn" href="/ventas">Nueva venta</a></div></div><div><img class="qr" src="{qr}"><p class="muted">Escanear con Yape, Plin o app bancaria. Confirmación manual.</p></div></div></div><div class="card"><h2>Detalle</h2><div class="table-wrap"><table><tr><th>Producto</th><th>Cantidad</th><th>Subtotal</th></tr>{rows}</table></div></div>'''
    return page('Pago QR','ventas',body)

@app.route('/marcar_pagado/<int:pedido_id>')
@login_required
def marcar_pagado(pedido_id): x("UPDATE pedidos SET pagado=1,estado='PAGADO' WHERE id=?",(pedido_id,)); flash('Pago confirmado.','ok'); return redirect(url_for('pedidos'))

@app.route('/pedidos')
@login_required
def pedidos():
    estado=request.args.get('estado','TODOS'); where='WHERE sucursal_id=?'; params=[session.get('sucursal_id')]
    if estado!='TODOS': where+=' AND estado=?'; params.append(estado)
    rows=q(f'SELECT * FROM pedidos {where} ORDER BY id DESC LIMIT 100',params)
    opts=''.join([f'<option {"selected" if estado==e else ""}>{e}</option>' for e in ['TODOS','PENDIENTE','PREPARACIÓN','LISTO','PAGADO','ANULADO']])
    trs=''.join([f'<tr><td>{r["codigo"]}</td><td>{r["fecha"]}</td><td>{r["cliente"]}</td><td>{r["servicio"]}</td><td>{r["estado"]}</td><td>{money(r["total"])}</td><td>{"SI" if r["pagado"] else "NO"}</td><td><a class="btn btn-sm" href="/estado/{r["id"]}/PREPARACIÓN">Preparar</a> <a class="btn btn-sm btn-green" href="/estado/{r["id"]}/LISTO">Listo</a> <a class="btn btn-sm btn-red" href="/pagar/{r["id"]}">Cobrar</a></td></tr>' for r in rows])
    body=f'''<div class="card"><h2>🚚 Control de pedidos / cocina</h2><form class="actions"><label>Filtrar por estado</label><select name="estado" style="max-width:260px">{opts}</select><button>Refrescar</button></form></div><div class="card"><h2>Listado de pedidos</h2><div class="table-wrap"><table><tr><th>Código</th><th>Fecha</th><th>Cliente</th><th>Servicio</th><th>Estado</th><th>Total</th><th>Pagado</th><th>Acciones</th></tr>{trs or '<tr><td colspan=8>Sin pedidos.</td></tr>'}</table></div></div>'''
    return page('Pedidos','pedidos',body)

@app.route('/estado/<int:pid>/<estado>')
@login_required
def estado(pid,estado): x('UPDATE pedidos SET estado=? WHERE id=?',(estado,pid)); return redirect(url_for('pedidos'))

@app.route('/inventario',methods=['GET','POST'])
@login_required
def inventario():
    if request.method=='POST':
        sku=request.form.get('sku','').upper().strip();
        if not sku_ok(sku): flash('SKU inválido. Usa letras, números y guion. Ej: PIZ-MUZ-01','err'); return redirect(url_for('inventario'))
        imgname=''; f=request.files.get('imagen')
        if f and f.filename:
            imgname=secure_filename(datetime.now().strftime('%Y%m%d%H%M%S_')+f.filename); f.save(os.path.join(UPLOAD_DIR,imgname))
        data=(sku,request.form.get('nombre','').upper(),request.form.get('categoria','PIZZAS'),int(float(request.form.get('precio') or 0)),int(float(request.form.get('costo') or 0)),int(float(request.form.get('stock') or 0)),int(float(request.form.get('stock_min') or 0)),imgname)
        if q('SELECT id FROM productos WHERE sku=?',(sku,),True): x('UPDATE productos SET nombre=?,categoria=?,precio=?,costo=?,stock=?,stock_min=?,imagen=COALESCE(NULLIF(?,""),imagen) WHERE sku=?',data[1:]+(sku,))
        else: x('INSERT INTO productos(sku,nombre,categoria,precio,costo,stock,stock_min,imagen) VALUES(?,?,?,?,?,?,?,?)',data)
        flash('Producto guardado.','ok'); return redirect(url_for('inventario'))
    prods=q('SELECT * FROM productos ORDER BY categoria,nombre')
    skus=''.join([f'<option value="{p["sku"]}">{p["sku"]} - {p["nombre"]}</option>' for p in prods])
    trs=''.join([f'<tr><td>{p["sku"]}</td><td>{p["nombre"]}</td><td>{p["categoria"]}</td><td>{money(p["precio"])}</td><td>{int(p["stock"])}</td><td>{int(p["stock_min"])}</td></tr>' for p in prods])
    body=f'''<div class="card"><h2>📦 Productos e inventario</h2><div class="quick-note">SKU recomendado para cadena: PIZ-MUZ-01, BEB-COCA-500, PAR-POLLO-01. Letras + números + guion.</div><br><form method="post" enctype="multipart/form-data"><div class="grid"><div><label>Código/SKU</label><input list="skuopts" name="sku" placeholder="PIZ-MUZ-01" required><datalist id="skuopts">{skus}</datalist><div class="sku-help">Lista con sugerencias. Puedes escribir uno nuevo.</div></div><div><label>Nombre</label><input name="nombre" required></div><div><label>Categoría</label><select name="categoria"><option>PIZZAS</option><option>PARRILLAS</option><option>BEBIDAS</option><option>POSTRES</option><option>EXTRAS</option></select></div><div><label>Precio entero</label><input name="precio" type="number" step="1"></div><div><label>Costo entero</label><input name="costo" type="number" step="1"></div><div><label>Stock entero</label><input name="stock" type="number" step="1"></div><div><label>Stock mínimo</label><input name="stock_min" type="number" step="1"></div><div><label>Imagen catálogo</label><input type="file" name="imagen" accept="image/*"></div></div><br><button class="btn-red">Guardar producto</button></form></div><div class="card"><h2>Lista de productos</h2><div class="table-wrap"><table><tr><th>SKU</th><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock</th><th>Mínimo</th></tr>{trs}</table></div></div>'''
    return page('Inventario','inventario',body)

@app.route('/catalogo',methods=['GET','POST'])
@login_required
def catalogo():
    cfg=q('SELECT * FROM catalogo_cfg WHERE id=1',(),True)
    if request.method=='POST':
        x('UPDATE catalogo_cfg SET nombre=?,slug=?,whatsapp=?,descripcion=? WHERE id=1',(request.form.get('nombre'),request.form.get('slug'),request.form.get('whatsapp'),request.form.get('descripcion'))); flash('Catálogo actualizado.','ok'); return redirect(url_for('catalogo'))
    link=request.host_url.rstrip()+'/c/'+cfg['slug']
    if qrcode:
        img=qrcode.make(link); img.save(os.path.join(STATIC_DIR,'qr_catalogo.png'))
    prods=q('SELECT * FROM productos WHERE activo=1 ORDER BY categoria,nombre')
    cards=''.join([f'<div class="catalog-card"><img src="{('/static/uploads/'+p['imagen']) if p['imagen'] else '/static/logo_toro.jpg'}"><h3>{p['nombre']}</h3><b>{money(p['precio'])}</b></div>' for p in prods[:12]])
    body=f'''<div class="grid2"><div class="card"><h2>🖼️ Catálogo online / QR</h2><form method="post"><label>Nombre comercio</label><input name="nombre" value="{cfg['nombre']}"><label>Slug enlace</label><input name="slug" value="{cfg['slug']}"><label>WhatsApp</label><input name="whatsapp" value="{cfg['whatsapp']}"><label>Descripción</label><input name="descripcion" value="{cfg['descripcion']}"><br><button class="btn-red">Guardar catálogo</button></form></div><div class="card"><h2>Compartir</h2><input value="{link}" readonly><br><br><img class="qr" src="/static/qr_catalogo.png"><br><br><a class="btn btn-green" target="_blank" href="{link}">Abrir catálogo público</a></div></div><div class="card"><h2>Vista previa</h2><div class="products">{cards}</div></div>'''
    return page('Catálogo / QR','catalogo',body)

@app.route('/c/<slug>')
def catalogo_publico(slug):
    cfg=q('SELECT * FROM catalogo_cfg WHERE slug=?',(slug,),True) or q('SELECT * FROM catalogo_cfg WHERE id=1',(),True)
    prods=q('SELECT * FROM productos WHERE activo=1 ORDER BY categoria,nombre')
    cards=''.join([f'<div class="prod"><img src="{('/static/uploads/'+p['imagen']) if p['imagen'] else '/static/logo_toro.jpg'}"><h3>{p['nombre']}</h3><b>{money(p['precio'])}</b><a class="btn btn-sm btn-red" href="https://wa.me/{cfg['whatsapp']}?text=Hola quiero {p['nombre']}">Pedir</a></div>' for p in prods])
    return render_template_string(f'<html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>{CSS}.main{{margin:0}}body{{background:#080812;color:white}}.content{{max-width:1100px;margin:auto}}</style></head><body><section class="content"><div class="login-card" style="margin:20px auto"><img class="logo-img" src="/static/logo_toro.jpg"><h1>{cfg["nombre"]}</h1><p>{cfg["descripcion"]}</p></div><div class="products">{cards}</div></section></body></html>')

@app.route('/usuarios',methods=['GET','POST'])
@login_required
@admin_required
def usuarios():
    if request.method=='POST':
        u=request.form.get('usuario'); clave=request.form.get('clave'); nombre=request.form.get('nombre'); rol=request.form.get('rol','VENDEDOR'); sid=request.form.get('sucursal_id')
        if q('SELECT id FROM usuarios WHERE usuario=?',(u,),True): flash('Usuario ya existe.','err')
        else: x('INSERT INTO usuarios(usuario,nombre,clave_hash,clave_plain,rol,sucursal_id) VALUES(?,?,?,?,?,?)',(u,nombre,generate_password_hash(clave),clave,rol,sid)); flash('Usuario creado.','ok')
    suc=q('SELECT * FROM sucursales ORDER BY nombre'); users=q('SELECT u.*,s.nombre sucursal FROM usuarios u LEFT JOIN sucursales s ON s.id=u.sucursal_id ORDER BY u.id')
    sucopts=''.join([f'<option value="{s["id"]}">{s["nombre"]}</option>' for s in suc])
    rows=''.join([f'<tr><td>{u["usuario"]}</td><td>{u["nombre"]}</td><td>{u["rol"]}</td><td>{u["sucursal"]}</td><td><input id="p{u["id"]}" type="password" value="{u["clave_plain"]}" readonly style="width:140px;padding:8px"><button class="btn-sm" onclick="tog(\'p{u["id"]}\')" type="button">👁</button></td><td><a class="btn btn-sm btn-red" href="/desactivar_usuario/{u["id"]}">Desactivar</a></td></tr>' for u in users])
    body=f'''<div class="grid2"><div class="card"><h2>👥 Crear usuario</h2><form method="post"><div class="grid2"><div><label>Usuario</label><input name="usuario" placeholder="vendedor2" required></div><div><label>Nombre</label><input name="nombre" required></div><div><label>Clave</label><input name="clave" required></div><div><label>Rol</label><select name="rol"><option>VENDEDOR</option><option>ADMIN</option></select></div><div><label>Sucursal</label><select name="sucursal_id">{sucopts}</select></div></div><br><button class="btn-red">Guardar usuario</button></form></div><div class="card"><h2>🏬 Crear sucursal</h2><form action="/crear_sucursal" method="post" class="actions"><input name="nombre" placeholder="Nueva sucursal" required><button class="btn-red">Registrar sucursal</button></form><p class="quick-note">Admin controla todo; vendedores trabajan ventas, pedidos, catálogo y cierre por sucursal.</p></div></div><div class="card"><h2>Usuarios creados</h2><div class="table-wrap"><table><tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Sucursal</th><th>Clave</th><th>Acción</th></tr>{rows}</table></div></div>'''
    return page('Usuarios / Admin','usuarios',body)

@app.route('/crear_sucursal',methods=['POST'])
@login_required
@admin_required
def crear_sucursal():
    try: x('INSERT INTO sucursales(nombre) VALUES(?)',(request.form.get('nombre'),)); flash('Sucursal creada.','ok')
    except Exception: flash('La sucursal ya existe.','err')
    return redirect(url_for('usuarios'))
@app.route('/desactivar_usuario/<int:uid>')
@login_required
@admin_required
def desactivar_usuario(uid): x('UPDATE usuarios SET activo=0 WHERE id=?',(uid,)); return redirect(url_for('usuarios'))

@app.route('/indicadores')
@login_required
def indicadores():
    sid=session.get('sucursal_id')
    total=q('SELECT COUNT(*) c, COALESCE(SUM(total),0) t FROM pedidos WHERE sucursal_id=?',(sid,),True)
    pagos=q('SELECT metodo, SUM(total) t FROM pedidos WHERE sucursal_id=? GROUP BY metodo',(sid,))
    top=q('SELECT producto,SUM(cantidad) c,SUM(subtotal) t FROM pedido_items GROUP BY producto ORDER BY c DESC LIMIT 10')
    pagos_rows=''.join([f'<tr><td>{r["metodo"]}</td><td>{money(r["t"])}</td></tr>' for r in pagos])
    top_rows=''.join([f'<tr><td>{r["producto"]}</td><td>{int(r["c"])}</td><td>{money(r["t"])}</td></tr>' for r in top])
    body=f'''<div class="grid"><div class="kpi">Facturación<b>{money(total['t'])}</b></div><div class="kpi">Ventas<b>{total['c']}</b></div><div class="kpi">Ticket promedio<b>{money((total['t'] or 0)/(total['c'] or 1))}</b></div><div class="kpi">Rentabilidad<b>PRO</b></div></div><div class="grid2"><div class="card"><h2>💳 Ventas por método de pago</h2><div class="table-wrap"><table><tr><th>Método</th><th>Total</th></tr>{pagos_rows or '<tr><td colspan=2>Sin pagos.</td></tr>'}</table></div></div><div class="card"><h2>🏆 Top productos vendidos</h2><div class="table-wrap"><table><tr><th>Producto</th><th>Cantidad</th><th>Total</th></tr>{top_rows or '<tr><td colspan=3>Sin productos vendidos.</td></tr>'}</table></div></div></div><div class="quick-note">Indicadores visibles completos: controla ticket promedio, productos top, pagos, stock y velocidad de atención.</div>'''
    return page('Indicadores','indicadores',body)

@app.route('/reportes')
@login_required
def reportes():
    rows=q('SELECT * FROM pedidos WHERE sucursal_id=? ORDER BY id DESC LIMIT 200',(session.get('sucursal_id'),))
    trs=''.join([f'<tr><td>{r["fecha"]}</td><td>{r["codigo"]}</td><td>{r["cliente"]}</td><td>{r["metodo"]}</td><td>{money(r["total"])}</td></tr>' for r in rows])
    body=f'<div class="card"><h2>📄 Reportería</h2><div class="actions"><a class="btn btn-red" href="/exportar_ventas">Exportar CSV</a><button onclick="printPage()">Imprimir</button></div><br><div class="table-wrap"><table><tr><th>Fecha</th><th>Código</th><th>Cliente</th><th>Método</th><th>Total</th></tr>{trs}</table></div></div>'
    return page('Reportes','reportes',body)

@app.route('/exportar_ventas')
@login_required
def exportar_ventas():
    rows=q('SELECT fecha,codigo,cliente,servicio,metodo,total,estado FROM pedidos WHERE sucursal_id=? ORDER BY id DESC',(session.get('sucursal_id'),))
    out=io.StringIO(); w=csv.writer(out); w.writerow(['fecha','codigo','cliente','servicio','metodo','total','estado']); [w.writerow([r[k] for k in r.keys()]) for r in rows]
    return send_file(io.BytesIO(out.getvalue().encode('utf-8-sig')), as_attachment=True, download_name='ventas_el_toro.csv', mimetype='text/csv')

@app.route('/plantilla_productos')
@login_required
def plantilla_productos():
    data='sku,nombre,categoria,precio,costo,stock,stock_min\nPIZ-MUZ-01,PIZZA MUZZARELLA,PIZZAS,32,14,50,5\n'
    return send_file(io.BytesIO(data.encode('utf-8-sig')), as_attachment=True, download_name='plantilla_productos.csv', mimetype='text/csv')

@app.route('/importar_productos',methods=['POST'])
@login_required
@admin_required
def importar_productos():
    f=request.files.get('archivo')
    if not f: flash('Selecciona archivo.','err'); return redirect(url_for('ventas'))
    name=f.filename.lower(); count=0
    if name.endswith('.csv'):
        text=f.read().decode('utf-8-sig'); rows=csv.DictReader(io.StringIO(text))
    else:
        if pd is None: flash('Falta pandas/openpyxl para Excel. Usa CSV.','err'); return redirect(url_for('ventas'))
        df=pd.read_excel(f); rows=df.fillna('').to_dict('records')
    for r in rows:
        sku=str(r.get('sku') or r.get('SKU') or '').upper().strip()
        if not sku_ok(sku): continue
        vals=(str(r.get('nombre') or r.get('NOMBRE') or '').upper(),str(r.get('categoria') or 'PIZZAS'),int(float(r.get('precio') or 0)),int(float(r.get('costo') or 0)),int(float(r.get('stock') or 0)),int(float(r.get('stock_min') or 0)),sku)
        if q('SELECT id FROM productos WHERE sku=?',(sku,),True): x('UPDATE productos SET nombre=?,categoria=?,precio=?,costo=?,stock=?,stock_min=? WHERE sku=?',vals)
        else: x('INSERT INTO productos(nombre,categoria,precio,costo,stock,stock_min,sku) VALUES(?,?,?,?,?,?,?)',vals)
        count+=1
    flash(f'Carga de día completada: {count} productos actualizados.','ok'); return redirect(url_for('ventas'))

@app.route('/cierre')
@login_required
def cierre():
    sid=session.get('sucursal_id'); fecha=today(); r=q('SELECT COUNT(*) c, COALESCE(SUM(total),0) t FROM pedidos WHERE fecha=? AND sucursal_id=?',(fecha,sid),True)
    body=f'''<div class="card"><h2>🔒 Cierre de día</h2><div class="grid"><div class="kpi">Pedidos<b>{r['c']}</b></div><div class="kpi">Total ventas<b>{money(r['t'])}</b></div></div><br><a class="btn btn-red" href="/cerrar_dia">Cerrar día</a> <a class="btn" href="/reportes">Ver reporte</a></div>'''
    return page('Cierre','cierre',body)
@app.route('/cerrar_dia')
@login_required
def cerrar_dia():
    sid=session.get('sucursal_id'); fecha=today(); r=q('SELECT COUNT(*) c, COALESCE(SUM(total),0) t FROM pedidos WHERE fecha=? AND sucursal_id=?',(fecha,sid),True)
    x('INSERT INTO cierres(fecha,sucursal_id,usuario,total_ventas,total_pedidos,cerrado_en) VALUES(?,?,?,?,?,?)',(fecha,sid,session.get('user'),r['t'],r['c'],datetime.now().isoformat()))
    flash('Día cerrado correctamente.','ok'); return redirect(url_for('cierre'))

@app.route('/delivery',methods=['GET','POST'])
@login_required
def delivery():
    if request.method=='POST': x('INSERT INTO clientes(nombre,telefono,direccion,referencia,notas) VALUES(?,?,?,?,?)',(request.form.get('nombre'),request.form.get('telefono'),request.form.get('direccion'),request.form.get('referencia'),request.form.get('notas'))); flash('Cliente guardado.','ok')
    rows=q('SELECT * FROM clientes ORDER BY id DESC')
    trs=''.join([f'<tr><td>{r["nombre"]}</td><td>{r["telefono"]}</td><td>{r["direccion"]}</td><td>{r["referencia"]}</td></tr>' for r in rows])
    body=f'''<div class="card"><h2>🛵 Delivery / clientes frecuentes</h2><div class="quick-note">Este módulo no se repite: sirve como CRM de direcciones. En ventas solo eliges si el pedido es delivery.</div><form method="post"><div class="grid"><input name="nombre" placeholder="Nombre"><input name="telefono" placeholder="Teléfono"><input name="direccion" placeholder="Dirección"><input name="referencia" placeholder="Referencia"><input name="notas" placeholder="Notas"></div><br><button class="btn-red">Guardar cliente</button></form></div><div class="card"><div class="table-wrap"><table><tr><th>Nombre</th><th>Teléfono</th><th>Dirección</th><th>Referencia</th></tr>{trs}</table></div></div>'''
    return page('Delivery','delivery',body)

@app.route('/recetas')
@login_required
def recetas(): return page('Recetas','recetas','<div class="card"><h2>🍽️ Recetas</h2><p>Módulo preparado para receta por producto y descuento automático de insumos.</p></div>')
@app.route('/caja')
@login_required
def caja(): return page('Caja','caja','<div class="card"><h2>💵 Caja</h2><p>Consulta pagos, arqueo y movimientos de caja.</p><a class="btn btn-red" href="/reportes">Ver ventas</a></div>')

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=True)
