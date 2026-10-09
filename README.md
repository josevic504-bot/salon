# Salón de Belleza · Sistema de gestión

Web app multiusuario para un salón: **agenda tipo calendario de eventos, citas, clientes, servicios, tratamientos por paquetes, productos con inventario, punto de venta y reportes**.

- **Frontend:** `index.html` (un solo archivo) publicado en **GitHub Pages**
- **Backend:** `Code.gs` en **Google Apps Script**
- **Base de datos:** una **Hoja de cálculo de Google** (en tu Drive)

---

## 1. Backend (Google Sheets + Apps Script)

1. En Google Drive crea una hoja nueva, por ejemplo **"Salón – Base de datos"**.
2. Menú **Extensiones → Apps Script**. Borra el código de ejemplo y pega todo `Code.gs`. Guarda.
3. En la barra superior elige la función **`instalar`** y pulsa **Ejecutar**. Autoriza los permisos (es tu propia cuenta).
   Se crean las hojas, la configuración, datos de ejemplo y el usuario inicial:
   - Usuario: **admin** · Contraseña: **admin123** (cámbiala al entrar)
4. **Implementar → Nueva implementación → ⚙ Tipo: App web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
5. Copia la **URL que termina en `/exec`**.

> Al modificar `Code.gs` más adelante: **Implementar → Gestionar implementaciones → ✏ Editar → Versión: Nueva versión**. Así la URL no cambia.

## 2. Frontend (GitHub Pages)

1. Abre `index.html` y pega la URL en la línea:
   ```js
   const API_URL_DEFAULT = 'https://script.google.com/macros/s/XXXX/exec';
   ```
   (Opcional: también se puede pegar desde el botón **"Configurar servidor"** en la pantalla de ingreso.)
2. Crea un repositorio en GitHub (ej. `salon`) y sube `index.html`.
3. **Settings → Pages → Deploy from a branch → `main` / `(root)`** → Save.
4. En 1–2 minutos tendrás el sistema en `https://TU-USUARIO.github.io/salon/`.

## 3. Primer uso

1. Ingresa con **admin / admin123** → **Ajustes**: nombre del salón, RTN, ISV, horario, mensaje de WhatsApp. Cambia tu contraseña.
2. **Usuarios**: crea a cada persona con su rol y un color (el color identifica sus citas en la agenda).
3. Ajusta **Servicios**, **Tratamientos** y **Productos** (vienen ejemplos que puedes editar o desactivar).

## Roles

| Rol | Puede |
|---|---|
| **Administrador** | Todo: catálogos, usuarios, ajustes, anular ventas y gastos, eliminar registros |
| **Recepción** | Agenda de todos, clientes, planes de tratamiento, caja/ventas, gastos, reportes |
| **Estilista** | Solo su propia agenda y sus ventas; cambia el estado de sus citas |

## Cómo funciona

- **Agenda:** vistas Mes / Semana / Día / Lista. Toca un espacio libre para crear una cita; arrastra o estira una cita para reprogramarla. Filtra por estilista con los chips y colorea por estilista o por servicio. Avisa si un estilista ya tiene una cita en ese horario.
- **Ciclo de una cita:** Pendiente → Confirmada → En proceso → Completada (o Cancelada / No asistió). Botón de **WhatsApp** con recordatorio listo para enviar. **Cobrar** pasa la cita directo a caja.
- **Tratamientos:** se venden como **paquete** (ej. 4 sesiones). Al venderlo se crea un *plan* para el cliente; cada sesión agendada y completada descuenta del plan. Se permiten **abonos** al saldo pendiente.
- **Ventas:** servicios, productos (descuenta inventario y no deja vender sin stock), paquetes y abonos. Descuento, método de pago, cálculo de cambio, ISV incluido o agregado, comprobante imprimible (80 mm). El admin puede **anular** (devuelve inventario y revierte abonos).
- **Varios servicios por cita:** una cita puede llevar uno o más servicios (ej. tinte + corte + peinado). La duración y el precio se suman solos, cada servicio puede ajustar su precio, y al cobrar se cargan por separado en el ticket.
- **Comprobante de pago digital:** al cobrar se genera una imagen elegante con el detalle, ISV, total, método de pago y cambio, lista para compartir por WhatsApp o descargar. Indica que no sustituye la factura fiscal.
- **Imagen para el cliente:** al guardar una cita se genera una tarjeta de confirmación (fecha, hora, servicio, estilista, valor, dirección y teléfono del salón) lista para **Compartir** por WhatsApp desde el teléfono o **Descargar**. También está en el detalle de cada cita. Llena dirección y teléfono en Ajustes para que salgan en la imagen.
- **Gastos:** registro por fecha, categoría, detalle, proveedor, monto y método de pago. Las categorías y detalles se sugieren según lo que ya se ha registrado (los más usados primero), con el último monto pagado; también se puede crear una categoría nueva. Recepción registra y corrige sus propios gastos; solo el administrador anula. Los anulados quedan visibles como respaldo.
- **Diario del mes** (Reportes): todos los ingresos y gastos del mes, agrupados por día, con totales diarios, acumulado y búsqueda; se descarga en CSV para Excel. Los anulados se ven tachados y no suman.
- **Historial y utilidades** (Reportes): resumen de cada mes (ingresos, gastos, utilidad). Cuando un mes cerrado deja utilidad, el administrador decide su destino: **enviarla a una cuenta** (ahorro, reinversión…) o **pagar algo** (préstamo, retiro del dueño…), en una o varias partes. El historial muestra el saldo acumulado por cuenta y los pagos hechos. Inicio avisa cuando hay utilidad sin destino.
- **Reportes:** ingresos, gastos, utilidad estimada (ventas sin ISV − gastos), ticket promedio, ISV, ingresos y gastos por día, gastos por categoría, por línea, más vendidos, desempeño por estilista, asistencia.

## Modo prueba

Si `index.html` no tiene URL de servidor, abre en **modo prueba**: corre la misma lógica de `Code.gs` dentro del navegador con datos ficticios (nada se envía a Google). Entra con los botones Administrador / Recepción / Estilista. La etiqueta **MODO PRUEBA** permite reiniciar los datos de ejemplo.

> Si modificas `Code.gs`, el modo prueba usa la copia que va incluida dentro de `index.html`; pide que se actualice para que ambos coincidan.

## Al actualizar a una versión nueva

Pega el `Code.gs` nuevo, guarda, **ejecuta `instalar()` otra vez** (no borra datos: solo agrega hojas y columnas nuevas) y actualiza la implementación con **Nueva versión**.

## Seguridad

- Contraseñas guardadas con **SHA‑256 + sal** (nunca en texto plano).
- Sesiones de 6 horas; bloqueo de 10 minutos tras 5 intentos fallidos.
- Cada acción se valida en el servidor según el rol. La URL `/exec` es pública, pero sin sesión válida no devuelve datos.
- Los datos viven en tu Google Drive; no compartas la hoja con quien no deba verla.

## Límites a tener en cuenta

- Cada operación tarda ~1–2 s (Apps Script). Rinde bien para un salón con varios miles de citas/ventas al año.
- Para borrar historial antiguo, archiva filas de las hojas **Citas**, **Ventas** y **VentaDetalle** a otra hoja.
- No edites los encabezados de las hojas; sí puedes consultar y filtrar los datos directamente en Sheets.
