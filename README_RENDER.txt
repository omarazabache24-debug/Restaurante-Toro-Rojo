RESTAURANTE EL TORO - ZIP FINAL PRO

Archivos obligatorios para Render:
- app.py
- requirements.txt
- Procfile
- runtime.txt
- static/toro_logo.png

Deploy en Render:
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app

Usuarios iniciales:
admin / admin123
vendedor1 / vendedor1
vendedor2 / vendedor2
vendedor3 / vendedor3

Notas:
- El QR de pago se genera visualmente con el monto y datos del negocio.
- Para cobro automático real Yape/Plin se requieren credenciales/API del proveedor de pago.
