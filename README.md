# Sistema Comedor PRIZE - Final PRO Render

## Mejoras incluidas
- Botón **Actualizar / refrescar** en Entregas corregido: ahora llama al API y recarga pedidos sin salir de la pantalla.
- Si no ingresas DNI en Entregas, el refresco muestra todos los pedidos del día.
- Usuarios guardados en SQLite con `commit` real y claves protegidas con hash.
- Auditoría de creación/actualización de usuarios.
- Persistencia local en `comedor_prize.db` para usuarios, consumos, entregas, reportes y cierres.
- Interfaz reforzada para celular desde login y módulos internos.
- Listo para GitHub + Render.

## Usuarios demo
- adm1 / adm1
- adm2 / adm2
- admin / admin123
- comedor / comedor123

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```

## Correo de auditoría de usuarios
Por seguridad, el sistema **no envía contraseñas por correo**. Sí puede enviar una notificación segura con usuario, rol, acción y fecha.

Variables opcionales en Render:
- ENABLE_ADMIN_USER_ALERTS=1
- ADMIN_AUDIT_EMAIL=omar.azabache24@gmail.com
- SMTP_HOST=smtp.gmail.com
- SMTP_PORT=587
- SMTP_USER=tu_correo@gmail.com
- SMTP_PASSWORD=tu_password_app
- SMTP_FROM=tu_correo@gmail.com

Si no configuras SMTP, el sistema genera el registro en `reportes_cierre/notificaciones_usuarios.txt`.

## Mejoras aplicadas - versión control consumos/cierre
- Administradores pueden quitar consumos registrados desde cualquier equipo conectado.
- Validación dinámica de DNI: al digitar DNI se muestra el trabajador.
- Responsable obligatorio en mayúsculas para registrar consumo.
- Botones de entrega corregidos: seleccionado y todos.
- Cierre por fecha seleccionable, con correo y hora de referencia para envío del Excel.
- Al elegir una fecha nueva se trabaja desde cero; al elegir una fecha cerrada se muestra su reporte histórico.

## Actualización PRO PostgreSQL + Lote + QR

Cambios incluidos:
- Persistencia en PostgreSQL para Render mediante `DATABASE_URL`. Ya no depende de archivos locales para usuarios, consumos, entregas y cierres.
- `render.yaml` crea/usa una base PostgreSQL llamada `comedor-prize-db` y conecta `DATABASE_URL` automáticamente.
- Registro masivo/en lote desde la pestaña Consumos con check "Registro masivo / lote". Valida cada DNI contra trabajadores activos y avisa DNI errados, no encontrados o duplicados.
- Escáner QR desde celular usando la cámara. El QR debe contener el DNI o un texto con el DNI.
- En local, si no existe `DATABASE_URL`, el sistema sigue funcionando con SQLite.

Importante para Render:
1. Subir todo el proyecto a GitHub.
2. En Render crear desde Blueprint usando este `render.yaml`, o crear una base PostgreSQL y colocar su `DATABASE_URL` en Environment.
3. Las imágenes/logos pueden quedar en `/static`, pero la información del sistema queda en PostgreSQL.
