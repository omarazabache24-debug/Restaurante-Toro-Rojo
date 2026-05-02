# RESTAURANTE AORIX - Render/GitHub

Sistema web Flask adaptado del archivo Restaurante_AORIX.py para ejecutarse en Render.

Usuarios iniciales:
- admin / admin123
- caja / caja123

Render:
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Variables: `SECRET_KEY`, `APP_TIMEZONE=America/Lima`, `PERSIST_DIR=/data`

Nota: para producción en Render usa disco persistente o base externa para no perder la BD.
