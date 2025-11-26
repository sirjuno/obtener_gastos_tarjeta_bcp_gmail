# Gastos Tarjeta BCP – Extracción automática desde Gmail hacia Google Sheets

Este proyecto contiene un **script en Google Apps Script** que lee correos enviados por  
**notificaciones@notificacionesbcp.com.pe**, extrae los datos de cada operación financiera  
(notificaciones BCP) y los almacena en una hoja de cálculo llamada:

> **Gastos Tarjeta BCP**

El script clasifica qué correos deben procesarse y cuáles no, obtiene fecha, monto, operación realizada, destinatario e ID del mensaje, y etiqueta el hilo procesado con:

> **tcprocesadabcp**

---

## 🧩 **Objetivo del proyecto**

Automatizar el registro de movimientos enviados por BCP para:

- Control personal de gastos  
- Auditoría interna  
- Consolidación mensual  
- Integración con dashboards (Looker Studio, Power BI, etc.)

El proceso:

1. Lee correos nuevos.
2. Ignora asuntos no válidos o no procesables.
3. Extrae los campos requeridos usando reglas específicas por tipo de operación.
4. Registra una fila nueva en Google Sheets.
5. Etiqueta el correo como “procesado”.

---

## 📦 **Requisitos**

Cada persona que use este proyecto debe contar con:

- Una cuenta Google con acceso a **Gmail** y **Google Sheets**
- Permisos para ejecutar scripts en Google Apps Script
- Una hoja creada llamada **exactamente**: `Gastos Tarjeta BCP`

---

## 📁 Archivos del repositorio

| Archivo | Descripción |
|--------|-------------|
| **Código.gs** | Script completo de Apps Script. |
| **README.md** | Este documento. |
| **.gitignore** | Ignora archivos locales no necesarios. |
| **LICENSE** | Licencia del proyecto (MIT recomendada). |
| **CONTRIBUTING.md** | Guía básica para contribuir. |

---

## ⚙️ **Configuración del proyecto (por persona)**

### 1️⃣ Crear la hoja de Google Sheets

Nombre exacto: Gastos Tarjeta BCP

Puedes dejarla vacía.  
El script creará encabezados automáticamente si no existen.

---

### 2️⃣ Cargar el código en Google Apps Script

1. Ir a: https://script.google.com  
2. Crear un **Nuevo Proyecto**  
3. Reemplazar el contenido por el archivo `Código.gs` de este repositorio  
4. Guardar con el nombre que prefieras  
5. Ejecutar **por primera vez** la función: extraerDatosBCP

6. Aceptar los permisos solicitados:
   - Gmail (lectura y etiquetado)
   - Google Sheets (escritura)
   - Utilities

---

### 3️⃣ Crear el Trigger (opcional pero recomendado)

1. En Apps Script → menú lateral → **Triggers**  
2. Crear un nuevo Trigger:
   - Función: `extraerDatosBCP`
   - Tipo: **Time-driven**
   - Frecuencia recomendada: cada 10–15 minutos / cada hora

---

### 4️⃣ Verificar etiqueta en Gmail

Si no existe, se creará automáticamente: tcprocesadabcp

---

## 🧪 **Logs y detección de errores**

El script registra en `Logger.log(...)` todos los correos cuyo:

- Monto  
- Destinatario  
- Operación realizada  

no se hayan podido procesar correctamente.

Esto facilita la depuración sin romper el flujo general.

---

## 🧾 **Asuntos que NO se procesan**

El script descarta automáticamente estos asuntos (lista incluida en código):

- Constancia de transferencia propias
- Configuración de Tarjeta  
- Se rechazó tu compra por E-Commerce no permitido - Servicio de Notificaciones BCP  
- Realizamos una devolución de una operación a tu Tarjeta de Débito BCP - Servicio de Notificaciones BCP  

---

## 🛠️ **Modificaciones por operación (ya integradas en el código)**

El script reconoce montos según el tipo de operación:

| Operación | Texto que precede al monto |
|----------|-----------------------------|
| Yapear a celular | `Monto enviado` |
| Yapeo a celular | `Monto enviado` |
| Pago de servicios | `Monto total:` |
| Pago de tarjeta propia BCP | `Monto pagado` |
| Aporte voluntario | `Total aportado` |
| Retiro | `Monto retirado` o `Total retirado` |
| Transferencia a terceros BCP | `Monto transferido` |
| Otros montos | Reconoce también formato tipo `$ 1.56` |

---

## 🖥 Cómo ejecutar manualmente

Dentro del editor de Apps Script:

1. Seleccionar función: `extraerDatosBCP`  
2. Dar clic en **Run**  
3. Ver logs en:  
   → *View* → *Logs*

---

## 🛡 Permisos requeridos (scopes)

Google solicitará autorizaciones para:

- `GmailApp` – leer y etiquetar correos  
- `SpreadsheetApp` – leer y escribir datos  
- `Utilities` – funciones de fecha  

---

## 🧯 Solución de problemas comunes

| Problema | Solución |
|---------|----------|
| La hoja se queda vacía | Confirma que se llame **exactamente** `Gastos Tarjeta BCP` |
| No se etiqueta el correo | Revisa permisos de Gmail |
| Algunos campos salen ERROR | Revisa logs → ajustar regex si correo cambia |
| Trigger no ejecuta | Revisar permisos y que la función esté bien seleccionada |

---

## 🤝 Cómo contribuir

1. Haz **fork** del repositorio  
2. Crea una rama: feature/mi-mejora
3. Realiza cambios  
4. Abre un Pull Request indicando qué mejoras hiciste  

---

## 📄 Licencia

Este proyecto está bajo la licencia **MIT**.  
Puedes usarlo, modificarlo y distribuirlo libremente.

---

¡Disfruta la automatización!  
Si deseas agregar más operaciones o mejorar el parsing, abre un issue o PR.