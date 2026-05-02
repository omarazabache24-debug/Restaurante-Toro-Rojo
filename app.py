# -*- coding: utf-8 -*-
"""RESTAURANTE AORIX - Web Flask listo para GitHub + Render."""
import os, sqlite3, csv
from io import BytesIO, StringIO
from datetime import datetime
from functools import wraps
from zoneinfo import ZoneInfo
from flask import Flask, request, redirect, url_for, session, render_template_string, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
try:
    from openpyxl import Workbook
    OPENPYXL=True
except Exception:
    OPENPYXL=False

BASE_DIR=os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR=os.getenv('PERSIST_DIR','/data' if os.path.isdir('/data') else os.path.join(BASE_DIR,'data'))
os.makedirs(PERSIST_DIR,exist_ok=True)
DB_PATH=os.path.join(PERSIST_DIR,'restaurante_aorix.db')
APP_TZ=ZoneInfo(os.getenv('APP_TIMEZONE','America/Lima'))
app=Flask(__name__)
app.secret_key=os.getenv('SECRET_KEY','aorix-restaurante-2026')
APP_TITLE='RESTAURANTE AORIX'
APP_SUBTITLE='Sistema de Control y Gestión de Alimentos'
BRAND='AORIX SYSTEMS - Automatizamos tu empresa'

def now(): return datetime.now(APP_TZ)
def today(): return now().date().isoformat()
def hour(): return now().strftime('%H:%M:%S')
def money(v):
    try: return 'S/ {:,.2f}'.format(float(v or 0))
    except Exception: return 'S/ 0.00'
def conn():
    c=sqlite3.connect(DB_PATH); c.row_factory=sqlite3.Row; return c
def q_all(sql,params=()):
    with conn() as c: return c.execute(sql,params).fetchall()
def q_one(sql,params=()):
    r=q_all(sql,params); return r[0] if r else None
def q_exec(sql,params=()):
    with conn() as c:
        cur=c.execute(sql,params); c.commit(); return cur.lastrowid
def log(accion,detalle=''):
    try: q_exec('INSERT INTO logs(fecha,hora,usuario,accion,detalle) VALUES(?,?,?,?,?)',(today(),hour(),session.get('user','sistema'),accion,detalle))
    except Exception: pass

def init_db():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS usuarios(id INTEGER PRIMARY KEY AUTOINCREMENT, usuario TEXT UNIQUE, clave_hash TEXT, rol TEXT DEFAULT 'admin', activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS productos(id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT UNIQUE, nombre TEXT, categoria TEXT DEFAULT '', precio REAL DEFAULT 0, stock REAL DEFAULT 0, stock_min REAL DEFAULT 0, unidad TEXT DEFAULT 'UND', activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS insumos(id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE, unidad TEXT DEFAULT 'UND', stock REAL DEFAULT 0, stock_min REAL DEFAULT 0, costo REAL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS recetas(id INTEGER PRIMARY KEY AUTOINCREMENT, producto_id INTEGER, insumo_id INTEGER, cantidad REAL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS clientes(id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT, telefono TEXT DEFAULT '', direccion TEXT DEFAULT '', referencia TEXT DEFAULT '', activo INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS ventas(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, hora TEXT, cliente TEXT DEFAULT '', tipo TEXT DEFAULT 'MOSTRADOR', metodo_pago TEXT DEFAULT 'EFECTIVO', subtotal REAL DEFAULT 0, descuento REAL DEFAULT 0, total REAL DEFAULT 0, usuario TEXT DEFAULT '', estado TEXT DEFAULT 'PAGADO');
        CREATE TABLE IF NOT EXISTS venta_detalle(id INTEGER PRIMARY KEY AUTOINCREMENT, venta_id INTEGER, producto_id INTEGER, producto TEXT, cantidad REAL, precio REAL, total REAL);
        CREATE TABLE IF NOT EXISTS pedidos(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, hora TEXT, cliente_id INTEGER, cliente TEXT, telefono TEXT, direccion TEXT, estado TEXT DEFAULT 'PENDIENTE', total REAL DEFAULT 0, observacion TEXT DEFAULT '', usuario TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS pedido_detalle(id INTEGER PRIMARY KEY AUTOINCREMENT, pedido_id INTEGER, producto_id INTEGER, producto TEXT, cantidad REAL, precio REAL, total REAL);
        CREATE TABLE IF NOT EXISTS caja(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, hora TEXT, tipo TEXT, concepto TEXT, monto REAL, usuario TEXT, venta_id INTEGER DEFAULT NULL);
        CREATE TABLE IF NOT EXISTS logs(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, hora TEXT, usuario TEXT, accion TEXT, detalle TEXT);
        """)
        c.commit()
    if not q_one('SELECT id FROM usuarios WHERE usuario=?',('admin',)):
        q_exec('INSERT INTO usuarios(usuario,clave_hash,rol,activo) VALUES(?,?,?,1)',('admin',generate_password_hash('admin123'),'admin'))
    if not q_one('SELECT id FROM usuarios WHERE usuario=?',('caja',)):
        q_exec('INSERT INTO usuarios(usuario,clave_hash,rol,activo) VALUES(?,?,?,1)',('caja',generate_password_hash('caja123'),'caja'))
    if not q_one('SELECT id FROM productos LIMIT 1'):
        for cod,nombre,cat,precio,stock in [('P001','MENU EJECUTIVO','COMIDA',12,100),('P002','ALMUERZO','COMIDA',10,100),('P003','GASEOSA','BEBIDA',4,50),('P004','POLLO A LA PLANCHA','COMIDA',18,40)]:
            q_exec('INSERT INTO productos(codigo,nombre,categoria,precio,stock,stock_min,unidad,activo) VALUES(?,?,?,?,?,?,?,1)',(cod,nombre,cat,precio,stock,5,'UND'))
init_db()

def login_required(fn):
    @wraps(fn)
    def w(*a,**kw):
        if not session.get('user'): return redirect(url_for('login'))
        return fn(*a,**kw)
    return w
def admin_required(fn):
    @wraps(fn)
    def w(*a,**kw):
        if session.get('rol')!='admin':
            flash('Solo administrador puede ingresar a esta opción.','error')
            return redirect(url_for('dashboard'))
        return fn(*a,**kw)
    return w

BASE = r'''
<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{title}}</title>
<style>
:root{--green:#16a34a;--blue:#0d73b8;--bg:#eef4f8;--line:#dbe7f0;--muted:#64748b}*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:#0f172a}a{text-decoration:none;color:inherit}button,.btn{border:0;border-radius:12px;padding:12px 16px;background:linear-gradient(135deg,var(--green),#0f8a3a);color:white;font-weight:900;cursor:pointer;display:inline-block}.btn-blue{background:linear-gradient(135deg,#1585cf,#075f9e)}.btn-orange{background:linear-gradient(135deg,#ff7a1a,#f05a05)}.btn-red{background:linear-gradient(135deg,#ef4444,#991b1b)}input,select,textarea{width:100%;border:1px solid var(--line);border-radius:12px;padding:12px 13px;background:white;outline:none}.login-page{min-height:100vh;display:grid;place-items:center;padding:20px;background:radial-gradient(circle at 10% 90%,rgba(13,115,184,.2),transparent 28%),radial-gradient(circle at 95% 95%,rgba(22,163,74,.22),transparent 32%),linear-gradient(135deg,#061b2b,#0b2d4a)}.login-card{width:min(440px,94vw);background:#111827e8;color:white;border:1px solid rgba(255,255,255,.18);border-radius:24px;padding:34px;box-shadow:0 24px 70px rgba(0,0,0,.35);text-align:center}.logo{font-size:48px;font-weight:950;letter-spacing:-2px;color:#bff0ff}.logo span{color:#ff7a1a}.login-card input{background:#1f2937;color:white;border-color:#374151}.login-card label{display:block;text-align:left;margin:14px 0 6px;font-weight:900}.login-card button{width:100%;margin-top:18px}.hint{color:#cbd5e1;font-size:12px;margin-top:16px;line-height:1.5}.app{display:grid;grid-template-columns:220px 1fr;min-height:100vh}.side{background:linear-gradient(180deg,#05243a,#041827);color:white;padding:18px 12px;position:sticky;top:0;height:100vh;overflow:auto}.side .brand{text-align:center;border-bottom:1px solid rgba(255,255,255,.15);padding-bottom:18px;margin-bottom:12px}.side .brand .logo{font-size:38px}.side small{color:#cfe6f3;font-weight:800}.nav a{display:flex;gap:9px;align-items:center;padding:12px 10px;border-radius:11px;font-weight:900;font-size:13px;margin:5px 0;color:#e6f6ff}.nav a:hover,.nav a.on{background:linear-gradient(90deg,#147a4a,#0d73b8)}.main{min-width:0}.top{background:linear-gradient(135deg,#061b2b,#0b2d4a);color:white;padding:18px 22px;text-align:center}.top h1{margin:0;font-size:32px}.top p{margin:5px 0 0;color:#d8e8f2;font-weight:800}.content{padding:18px;max-width:1500px;margin:0 auto}.head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:16px}.head h2{margin:0;font-size:28px}.head p{margin:4px 0 0;color:var(--muted)}.card{background:#fff;border:1px solid var(--line);border-radius:18px;padding:18px;box-shadow:0 12px 30px rgba(15,35,55,.07);margin-bottom:16px}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.kpi{display:flex;gap:12px;align-items:center}.ico{width:54px;height:54px;border-radius:50%;display:grid;place-items:center;font-size:26px;background:#e6f6ee}.kpi b{font-size:26px}.muted{color:var(--muted)}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:14px;background:white}table{width:100%;border-collapse:collapse;min-width:920px}th,td{padding:11px;border-bottom:1px solid #eef2f7;text-align:left;white-space:nowrap;font-size:13px}th{background:#f8fafc;font-weight:950}.badge{border-radius:999px;padding:6px 10px;font-weight:950;font-size:12px;display:inline-block}.ok{background:#dcfce7;color:#166534}.warn{background:#fef3c7;color:#92400e}.off{background:#fee2e2;color:#991b1b}.flash{padding:12px 14px;border-radius:12px;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;font-weight:850;margin-bottom:12px}.flash.error{background:#fff1f2;color:#991b1b;border-color:#fecaca}.mini{font-size:12px;color:var(--muted)}@media(max-width:850px){.app{display:block}.side{height:auto;position:relative;padding:8px}.side .brand{display:none}.nav{display:grid;grid-template-columns:repeat(2,1fr);gap:7px}.nav a{justify-content:center;text-align:center;background:rgba(255,255,255,.07);margin:0}.top h1{font-size:23px}.content{padding:10px}.head{display:block}.grid,.grid2,.kpis{grid-template-columns:1fr}.card{padding:13px}.table-wrap{max-height:55vh}button,.btn{width:100%;text-align:center}table{min-width:760px}}
</style></head><body>{% if session.get('user') %}<div class="app"><aside class="side"><div class="brand"><div class="logo">AOR<span>IX</span></div><small>{{brand}}</small><br><small>{{session.get('user')}} - {{session.get('rol')}}</small></div><nav class="nav"><a class="{{'on' if active=='dashboard' else ''}}" href="{{url_for('dashboard')}}">📊 Dashboard</a><a class="{{'on' if active=='ventas' else ''}}" href="{{url_for('ventas')}}">🧾 Ventas</a><a class="{{'on' if active=='pedidos' else ''}}" href="{{url_for('pedidos')}}">🚚 Pedidos</a><a class="{{'on' if active=='inventario' else ''}}" href="{{url_for('inventario')}}">📦 Inventario</a><a class="{{'on' if active=='recetas' else ''}}" href="{{url_for('recetas')}}">🍽️ Recetas</a><a class="{{'on' if active=='caja' else ''}}" href="{{url_for('caja')}}">💵 Caja</a><a class="{{'on' if active=='clientes' else ''}}" href="{{url_for('clientes')}}">👥 Clientes</a><a class="{{'on' if active=='reportes' else ''}}" href="{{url_for('reportes')}}">📈 Reportes</a><a class="{{'on' if active=='admin' else ''}}" href="{{url_for('admin')}}">⚙️ Admin</a><a href="{{url_for('logout')}}">🚪 Salir</a></nav></aside><main class="main"><header class="top"><h1>{{title}}</h1><p>{{subtitle}}</p></header><div class="content">{% with msgs=get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}{% endwith %}{{content|safe}}</div></main></div>{% else %}{{content|safe}}{% endif %}</body></html>
'''
def page(content,active='dashboard'):
    return render_template_string(BASE,content=content,active=active,title=APP_TITLE,subtitle=APP_SUBTITLE,brand=BRAND,money=money)

@app.route('/',methods=['GET','POST'])
def login():
    if session.get('user'): return redirect(url_for('dashboard'))
    if request.method=='POST':
        u=request.form.get('usuario','').strip(); p=request.form.get('clave','')
        r=q_one('SELECT * FROM usuarios WHERE usuario=? AND activo=1',(u,))
        if r and check_password_hash(r['clave_hash'],p):
            session['user']=r['usuario']; session['rol']=r['rol']; log('LOGIN','Ingreso correcto'); return redirect(url_for('dashboard'))
        flash('Usuario o clave incorrectos.','error')
    return page('<div class="login-page"><form class="login-card" method="post"><div class="logo">AOR<span>IX</span></div><h2>Restaurante AORIX</h2><p class="muted">Acceso al sistema</p><label>Usuario</label><input name="usuario" placeholder="Ingrese su usuario" autofocus><label>Clave</label><input type="password" name="clave" placeholder="Ingrese su clave"><button>Ingresar</button><div class="hint">Demo inicial: admin / admin123<br>Usuario caja: caja / caja123</div></form></div>')
@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    f=today(); total=q_one('SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha=?',(f,)); pedidos=q_one('SELECT COUNT(*) c FROM pedidos WHERE fecha=? AND estado<>"ENTREGADO"',(f,))['c']; stock_bajo=q_one('SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min')['c']; caja_total=q_one('SELECT COALESCE(SUM(CASE WHEN tipo="INGRESO" THEN monto ELSE -monto END),0) t FROM caja WHERE fecha=?',(f,))['t']
    rows=q_all('SELECT * FROM ventas ORDER BY id DESC LIMIT 10')
    trs=''.join(f'<tr><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["cliente"]}</td><td>{r["tipo"]}</td><td>{money(r["total"])}</td><td><span class="badge ok">{r["estado"]}</span></td></tr>' for r in rows) or '<tr><td colspan="6">Sin ventas.</td></tr>'
    return page(f'<div class="head"><div><h2>Panel principal</h2><p>Resumen del día {f}</p></div></div><div class="kpis"><div class="card kpi"><div class="ico">💰</div><div><div class="muted">Ventas hoy</div><b>{money(total["t"])}</b><div class="mini">{total["c"]} comprobantes</div></div></div><div class="card kpi"><div class="ico">🚚</div><div><div class="muted">Pedidos pendientes</div><b>{pedidos}</b></div></div><div class="card kpi"><div class="ico">📦</div><div><div class="muted">Stock bajo</div><b>{stock_bajo}</b></div></div><div class="card kpi"><div class="ico">💵</div><div><div class="muted">Caja neta</div><b>{money(caja_total)}</b></div></div></div><div class="card"><h3>Últimas ventas</h3><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Cliente</th><th>Tipo</th><th>Total</th><th>Estado</th></tr></thead><tbody>{trs}</tbody></table></div></div>','dashboard')

def descontar_receta(producto_id,cantidad_producto):
    for r in q_all('SELECT * FROM recetas WHERE producto_id=?',(producto_id,)):
        q_exec('UPDATE insumos SET stock=stock-? WHERE id=?',(float(r['cantidad'])*float(cantidad_producto),r['insumo_id']))

@app.route('/ventas',methods=['GET','POST'])
@login_required
def ventas():
    if request.method=='POST':
        producto_id=int(request.form.get('producto_id') or 0); cantidad=float(request.form.get('cantidad') or 1); prod=q_one('SELECT * FROM productos WHERE id=? AND activo=1',(producto_id,))
        if not prod: flash('Selecciona un producto válido.','error'); return redirect(url_for('ventas'))
        precio=float(request.form.get('precio') or prod['precio']); desc=float(request.form.get('descuento') or 0); total=max(cantidad*precio-desc,0)
        vid=q_exec('INSERT INTO ventas(fecha,hora,cliente,tipo,metodo_pago,subtotal,descuento,total,usuario,estado) VALUES(?,?,?,?,?,?,?,?,?,?)',(today(),hour(),request.form.get('cliente','MOSTRADOR').upper(),request.form.get('tipo','MOSTRADOR'),request.form.get('metodo_pago','EFECTIVO'),cantidad*precio,desc,total,session['user'],'PAGADO'))
        q_exec('INSERT INTO venta_detalle(venta_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)',(vid,producto_id,prod['nombre'],cantidad,precio,total)); q_exec('UPDATE productos SET stock=stock-? WHERE id=?',(cantidad,producto_id)); q_exec('INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario,venta_id) VALUES(?,?,?,?,?,?,?)',(today(),hour(),'INGRESO','VENTA '+str(vid),total,session['user'],vid)); descontar_receta(producto_id,cantidad); log('VENTA',f'Venta {vid}'); flash('Venta registrada correctamente.','ok'); return redirect(url_for('ventas'))
    productos=q_all('SELECT * FROM productos WHERE activo=1 ORDER BY nombre'); opts=''.join(f'<option value="{p["id"]}" data-precio="{p["precio"]}">{p["nombre"]} - {money(p["precio"])} - Stock {p["stock"]}</option>' for p in productos); ventas=q_all('SELECT * FROM ventas ORDER BY id DESC LIMIT 80')
    trs=''.join(f'<tr><td>{v["id"]}</td><td>{v["fecha"]}</td><td>{v["hora"]}</td><td>{v["cliente"]}</td><td>{v["tipo"]}</td><td>{v["metodo_pago"]}</td><td>{money(v["total"])}</td><td><a class="btn btn-blue" href="{url_for("ticket",venta_id=v["id"])}">Ticket</a></td></tr>' for v in ventas) or '<tr><td colspan="8">Sin ventas.</td></tr>'
    return page(f'<div class="head"><div><h2>Ventas</h2><p>Registro rápido de ventas y ticket imprimible</p></div></div><div class="card"><form method="post" class="grid"><input name="cliente" placeholder="Cliente / Mostrador"><select name="producto_id" id="producto" onchange="precio.value=this.selectedOptions[0].dataset.precio">{opts}</select><input name="cantidad" type="number" step="0.01" value="1"><input id="precio" name="precio" type="number" step="0.01" placeholder="Precio"><select name="tipo"><option>MOSTRADOR</option><option>DELIVERY</option><option>MESA</option></select><select name="metodo_pago"><option>EFECTIVO</option><option>YAPE</option><option>PLIN</option><option>TARJETA</option><option>TRANSFERENCIA</option></select><input name="descuento" type="number" step="0.01" value="0"><button>Registrar venta</button></form></div><div class="card"><h3>Ventas recientes</h3><div class="table-wrap"><table><thead><tr><th>ID</th><th>Fecha</th><th>Hora</th><th>Cliente</th><th>Tipo</th><th>Pago</th><th>Total</th><th>Ticket</th></tr></thead><tbody>{trs}</tbody></table></div></div><script>window.addEventListener("load",()=>{{let s=document.getElementById("producto"); if(s&&s.selectedOptions[0]) precio.value=s.selectedOptions[0].dataset.precio;}})</script>','ventas')

@app.route('/ticket/<int:venta_id>')
@login_required
def ticket(venta_id):
    v=q_one('SELECT * FROM ventas WHERE id=?',(venta_id,)); det=q_all('SELECT * FROM venta_detalle WHERE venta_id=?',(venta_id,))
    if not v: return 'Venta no encontrada',404
    lines=[APP_TITLE,'TICKET DE VENTA',f'Nro: {v["id"]}',f'Fecha: {v["fecha"]} {v["hora"]}',f'Cliente: {v["cliente"]}','-'*32]
    for d in det: lines.append(f'{d["cantidad"]} x {d["producto"]}  {money(d["total"])}')
    lines+=['-'*32,f'TOTAL: {money(v["total"])}','Gracias por su compra']
    return send_file(BytesIO('\n'.join(lines).encode('utf-8')),as_attachment=True,download_name=f'ticket_{venta_id}.txt',mimetype='text/plain')

@app.route('/pedidos',methods=['GET','POST'])
@login_required
def pedidos():
    if request.method=='POST':
        cliente_id=int(request.form.get('cliente_id') or 0); producto_id=int(request.form.get('producto_id') or 0); cantidad=float(request.form.get('cantidad') or 1); cli=q_one('SELECT * FROM clientes WHERE id=?',(cliente_id,)); prod=q_one('SELECT * FROM productos WHERE id=?',(producto_id,))
        if not cli or not prod: flash('Cliente y producto son obligatorios.','error'); return redirect(url_for('pedidos'))
        total=cantidad*float(prod['precio']); pid=q_exec('INSERT INTO pedidos(fecha,hora,cliente_id,cliente,telefono,direccion,estado,total,observacion,usuario) VALUES(?,?,?,?,?,?,?,?,?,?)',(today(),hour(),cli['id'],cli['nombre'],cli['telefono'],cli['direccion'],'PENDIENTE',total,request.form.get('observacion',''),session['user'])); q_exec('INSERT INTO pedido_detalle(pedido_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)',(pid,prod['id'],prod['nombre'],cantidad,prod['precio'],total)); log('PEDIDO',f'Pedido {pid}'); flash('Pedido registrado.','ok'); return redirect(url_for('pedidos'))
    clientes=q_all('SELECT * FROM clientes WHERE activo=1 ORDER BY nombre'); productos=q_all('SELECT * FROM productos WHERE activo=1 ORDER BY nombre'); opts_cli=''.join(f'<option value="{c["id"]}">{c["nombre"]} - {c["telefono"]}</option>' for c in clientes); opts_prod=''.join(f'<option value="{p["id"]}">{p["nombre"]} - {money(p["precio"])}</option>' for p in productos); rows=q_all('SELECT * FROM pedidos ORDER BY id DESC LIMIT 80')
    trs=''.join(f'<tr><td>{r["id"]}</td><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["cliente"]}</td><td>{r["telefono"]}</td><td>{r["direccion"]}</td><td>{money(r["total"])}</td><td><span class="badge warn">{r["estado"]}</span></td><td><a class="btn btn-blue" href="{url_for("pedido_estado",pedido_id=r["id"],estado="ENTREGADO")}">Entregado</a></td></tr>' for r in rows) or '<tr><td colspan="9">Sin pedidos.</td></tr>'
    return page(f'<div class="head"><div><h2>Pedidos / Delivery</h2><p>Control de pedidos y entregas</p></div></div><div class="card"><form method="post" class="grid"><select name="cliente_id">{opts_cli}</select><select name="producto_id">{opts_prod}</select><input name="cantidad" type="number" step="0.01" value="1"><input name="observacion" placeholder="Observación"><button>Crear pedido</button></form><p class="mini">Primero registra clientes en el módulo Clientes.</p></div><div class="card"><h3>Pedidos</h3><div class="table-wrap"><table><thead><tr><th>ID</th><th>Fecha</th><th>Hora</th><th>Cliente</th><th>Teléfono</th><th>Dirección</th><th>Total</th><th>Estado</th><th>Acción</th></tr></thead><tbody>{trs}</tbody></table></div></div>','pedidos')
@app.route('/pedido_estado/<int:pedido_id>/<estado>')
@login_required
def pedido_estado(pedido_id,estado): q_exec('UPDATE pedidos SET estado=? WHERE id=?',(estado,pedido_id)); flash('Estado actualizado.','ok'); return redirect(url_for('pedidos'))

@app.route('/inventario',methods=['GET','POST'])
@login_required
def inventario():
    if request.method=='POST':
        tipo=request.form.get('tipo','producto')
        if tipo=='producto':
            codigo=request.form.get('codigo') or 'P'+now().strftime('%H%M%S'); q_exec('INSERT OR REPLACE INTO productos(codigo,nombre,categoria,precio,stock,stock_min,unidad,activo) VALUES(?,?,?,?,?,?,?,1)',(codigo,request.form.get('nombre','').upper(),request.form.get('categoria',''),float(request.form.get('precio') or 0),float(request.form.get('stock') or 0),float(request.form.get('stock_min') or 0),request.form.get('unidad','UND')))
        else: q_exec('INSERT OR REPLACE INTO insumos(nombre,unidad,stock,stock_min,costo) VALUES(?,?,?,?,?)',(request.form.get('nombre','').upper(),request.form.get('unidad','UND'),float(request.form.get('stock') or 0),float(request.form.get('stock_min') or 0),float(request.form.get('costo') or 0)))
        flash('Inventario actualizado.','ok'); return redirect(url_for('inventario'))
    productos=q_all('SELECT * FROM productos ORDER BY nombre'); insumos=q_all('SELECT * FROM insumos ORDER BY nombre')
    trp=''.join(f'<tr><td>{p["codigo"]}</td><td>{p["nombre"]}</td><td>{p["categoria"]}</td><td>{money(p["precio"])}</td><td>{p["stock"]}</td><td>{p["stock_min"]}</td><td><span class="badge {"off" if p["stock"]<=p["stock_min"] else "ok"}">{"BAJO" if p["stock"]<=p["stock_min"] else "OK"}</span></td></tr>' for p in productos); tri=''.join(f'<tr><td>{i["nombre"]}</td><td>{i["unidad"]}</td><td>{i["stock"]}</td><td>{i["stock_min"]}</td><td>{money(i["costo"])}</td></tr>' for i in insumos) or '<tr><td colspan="5">Sin insumos.</td></tr>'
    return page(f'<div class="head"><div><h2>Inventario</h2><p>Productos, insumos y stock mínimo</p></div></div><div class="grid2"><div class="card"><h3>Nuevo producto</h3><form method="post" class="grid2"><input type="hidden" name="tipo" value="producto"><input name="codigo" placeholder="Código"><input name="nombre" placeholder="Producto"><input name="categoria" placeholder="Categoría"><input name="unidad" value="UND"><input name="precio" type="number" step="0.01" placeholder="Precio"><input name="stock" type="number" step="0.01" placeholder="Stock"><input name="stock_min" type="number" step="0.01" placeholder="Stock mínimo"><button>Guardar producto</button></form></div><div class="card"><h3>Nuevo insumo</h3><form method="post" class="grid2"><input type="hidden" name="tipo" value="insumo"><input name="nombre" placeholder="Insumo"><input name="unidad" value="UND"><input name="stock" type="number" step="0.01" placeholder="Stock"><input name="stock_min" type="number" step="0.01" placeholder="Stock mínimo"><input name="costo" type="number" step="0.01" placeholder="Costo"><button>Guardar insumo</button></form></div></div><div class="card"><h3>Productos</h3><div class="table-wrap"><table><thead><tr><th>Código</th><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock</th><th>Mínimo</th><th>Estado</th></tr></thead><tbody>{trp}</tbody></table></div></div><div class="card"><h3>Insumos</h3><div class="table-wrap"><table><thead><tr><th>Insumo</th><th>Unidad</th><th>Stock</th><th>Mínimo</th><th>Costo</th></tr></thead><tbody>{tri}</tbody></table></div></div>','inventario')

@app.route('/recetas',methods=['GET','POST'])
@login_required
def recetas():
    if request.method=='POST': q_exec('INSERT INTO recetas(producto_id,insumo_id,cantidad) VALUES(?,?,?)',(int(request.form.get('producto_id')),int(request.form.get('insumo_id')),float(request.form.get('cantidad') or 0))); flash('Receta agregada.','ok'); return redirect(url_for('recetas'))
    productos=q_all('SELECT * FROM productos ORDER BY nombre'); insumos=q_all('SELECT * FROM insumos ORDER BY nombre'); opts_p=''.join(f'<option value="{p["id"]}">{p["nombre"]}</option>' for p in productos); opts_i=''.join(f'<option value="{i["id"]}">{i["nombre"]} ({i["unidad"]})</option>' for i in insumos); rows=q_all('SELECT r.id,p.nombre producto,i.nombre insumo,r.cantidad,i.unidad FROM recetas r JOIN productos p ON p.id=r.producto_id JOIN insumos i ON i.id=r.insumo_id ORDER BY p.nombre'); trs=''.join(f'<tr><td>{r["producto"]}</td><td>{r["insumo"]}</td><td>{r["cantidad"]}</td><td>{r["unidad"]}</td></tr>' for r in rows) or '<tr><td colspan="4">Sin recetas.</td></tr>'
    return page(f'<div class="head"><div><h2>Recetas</h2><p>Descuento automático de insumos por producto vendido</p></div></div><div class="card"><form method="post" class="grid"><select name="producto_id">{opts_p}</select><select name="insumo_id">{opts_i}</select><input name="cantidad" type="number" step="0.0001" placeholder="Cantidad por producto"><button>Agregar insumo a receta</button></form></div><div class="card"><div class="table-wrap"><table><thead><tr><th>Producto</th><th>Insumo</th><th>Cantidad</th><th>Unidad</th></tr></thead><tbody>{trs}</tbody></table></div></div>','recetas')

@app.route('/caja',methods=['GET','POST'])
@login_required
def caja():
    if request.method=='POST': q_exec('INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario) VALUES(?,?,?,?,?,?)',(today(),hour(),request.form.get('tipo'),request.form.get('concepto','').upper(),float(request.form.get('monto') or 0),session['user'])); flash('Movimiento registrado.','ok'); return redirect(url_for('caja'))
    fecha=request.args.get('fecha',today()); rows=q_all('SELECT * FROM caja WHERE fecha=? ORDER BY id DESC',(fecha,)); total=q_one('SELECT COALESCE(SUM(CASE WHEN tipo="INGRESO" THEN monto ELSE -monto END),0) t FROM caja WHERE fecha=?',(fecha,))['t']; trs=''.join(f'<tr><td>{r["hora"]}</td><td>{r["tipo"]}</td><td>{r["concepto"]}</td><td>{money(r["monto"])}</td><td>{r["usuario"]}</td></tr>' for r in rows) or '<tr><td colspan="5">Sin movimientos.</td></tr>'
    return page(f'<div class="head"><div><h2>Caja</h2><p>Saldo neto: <b>{money(total)}</b></p></div></div><div class="card"><form method="post" class="grid"><select name="tipo"><option>INGRESO</option><option>EGRESO</option></select><input name="concepto" placeholder="Concepto"><input name="monto" type="number" step="0.01" placeholder="Monto"><button>Registrar movimiento</button></form></div><div class="card"><form method="get" class="grid2"><input type="date" name="fecha" value="{fecha}"><button class="btn-blue">Filtrar</button></form></div><div class="card"><div class="table-wrap"><table><thead><tr><th>Hora</th><th>Tipo</th><th>Concepto</th><th>Monto</th><th>Usuario</th></tr></thead><tbody>{trs}</tbody></table></div></div>','caja')

@app.route('/clientes',methods=['GET','POST'])
@login_required
def clientes():
    if request.method=='POST': q_exec('INSERT INTO clientes(nombre,telefono,direccion,referencia,activo) VALUES(?,?,?,?,1)',(request.form.get('nombre','').upper(),request.form.get('telefono',''),request.form.get('direccion','').upper(),request.form.get('referencia','').upper())); flash('Cliente guardado.','ok'); return redirect(url_for('clientes'))
    rows=q_all('SELECT * FROM clientes ORDER BY id DESC'); trs=''.join(f'<tr><td>{r["nombre"]}</td><td>{r["telefono"]}</td><td>{r["direccion"]}</td><td>{r["referencia"]}</td><td><span class="badge ok">ACTIVO</span></td></tr>' for r in rows) or '<tr><td colspan="5">Sin clientes.</td></tr>'
    return page(f'<div class="head"><div><h2>Clientes</h2><p>Base para delivery y pedidos</p></div></div><div class="card"><form method="post" class="grid"><input name="nombre" placeholder="Nombre cliente"><input name="telefono" placeholder="Teléfono"><input name="direccion" placeholder="Dirección"><input name="referencia" placeholder="Referencia"><button>Guardar cliente</button></form></div><div class="card"><div class="table-wrap"><table><thead><tr><th>Cliente</th><th>Teléfono</th><th>Dirección</th><th>Referencia</th><th>Estado</th></tr></thead><tbody>{trs}</tbody></table></div></div>','clientes')

@app.route('/reportes')
@login_required
def reportes():
    fi=request.args.get('fi',today()); ff=request.args.get('ff',fi); ventas=q_all('SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha DESC,hora DESC',(fi,ff)); total=sum(float(v['total'] or 0) for v in ventas); trs=''.join(f'<tr><td>{v["fecha"]}</td><td>{v["hora"]}</td><td>{v["cliente"]}</td><td>{v["tipo"]}</td><td>{v["metodo_pago"]}</td><td>{money(v["total"])}</td></tr>' for v in ventas) or '<tr><td colspan="6">Sin datos.</td></tr>'
    return page(f'<div class="head"><div><h2>Reportes</h2><p>Total del periodo: <b>{money(total)}</b></p></div><a class="btn btn-blue" href="{url_for("export_excel",fi=fi,ff=ff)}">Exportar Excel</a></div><div class="card"><form method="get" class="grid"><input type="date" name="fi" value="{fi}"><input type="date" name="ff" value="{ff}"><button class="btn-blue">Filtrar</button></form></div><div class="card"><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Cliente</th><th>Tipo</th><th>Pago</th><th>Total</th></tr></thead><tbody>{trs}</tbody></table></div></div>','reportes')

@app.route('/export_excel')
@login_required
def export_excel():
    fi=request.args.get('fi',today()); ff=request.args.get('ff',fi); rows=q_all('SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora',(fi,ff))
    if OPENPYXL:
        wb=Workbook(); ws=wb.active; ws.title='Ventas'; ws.append(['ID','Fecha','Hora','Cliente','Tipo','Metodo pago','Subtotal','Descuento','Total','Usuario','Estado'])
        for r in rows: ws.append([r['id'],r['fecha'],r['hora'],r['cliente'],r['tipo'],r['metodo_pago'],r['subtotal'],r['descuento'],r['total'],r['usuario'],r['estado']])
        bio=BytesIO(); wb.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f'reporte_aorix_{fi}_{ff}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    out=StringIO(); w=csv.writer(out); w.writerow(['ID','Fecha','Hora','Cliente','Tipo','Metodo pago','Total'])
    for r in rows: w.writerow([r['id'],r['fecha'],r['hora'],r['cliente'],r['tipo'],r['metodo_pago'],r['total']])
    return send_file(BytesIO(out.getvalue().encode('utf-8-sig')),as_attachment=True,download_name='reporte_aorix.csv',mimetype='text/csv')

@app.route('/admin',methods=['GET','POST'])
@login_required
@admin_required
def admin():
    if request.method=='POST':
        u=request.form.get('usuario','').strip(); clave=request.form.get('clave','').strip(); rol=request.form.get('rol','caja')
        if not u or not clave: flash('Usuario y clave son obligatorios.','error')
        else: q_exec('INSERT OR REPLACE INTO usuarios(usuario,clave_hash,rol,activo) VALUES(?,?,?,1)',(u,generate_password_hash(clave),rol)); flash('Usuario guardado.','ok')
        return redirect(url_for('admin'))
    usuarios=q_all('SELECT id,usuario,rol,activo FROM usuarios ORDER BY usuario'); logs=q_all('SELECT * FROM logs ORDER BY id DESC LIMIT 100'); tru=''.join(f'<tr><td>{u["usuario"]}</td><td>{u["rol"]}</td><td>{u["activo"]}</td></tr>' for u in usuarios); trl=''.join(f'<tr><td>{l["fecha"]}</td><td>{l["hora"]}</td><td>{l["usuario"]}</td><td>{l["accion"]}</td><td>{l["detalle"]}</td></tr>' for l in logs) or '<tr><td colspan="5">Sin logs.</td></tr>'
    return page(f'<div class="head"><div><h2>Administrador</h2><p>Usuarios y bitácora</p></div></div><div class="card"><h3>Crear / actualizar usuario</h3><form method="post" class="grid"><input name="usuario" placeholder="Usuario"><input name="clave" placeholder="Clave"><select name="rol"><option>admin</option><option>caja</option><option>mozo</option></select><button>Guardar usuario</button></form></div><div class="grid2"><div class="card"><h3>Usuarios</h3><div class="table-wrap"><table><thead><tr><th>Usuario</th><th>Rol</th><th>Activo</th></tr></thead><tbody>{tru}</tbody></table></div></div><div class="card"><h3>Bitácora</h3><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Usuario</th><th>Acción</th><th>Detalle</th></tr></thead><tbody>{trl}</tbody></table></div></div></div>','admin')

@app.errorhandler(500)
def err500(e): return page('<div class="card"><h2>Error interno controlado</h2><p>Revisa los logs de Render. El sistema no eliminó información.</p><a class="btn btn-blue" href="/dashboard">Volver al panel</a></div>'),500
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT',5000)),debug=False)
