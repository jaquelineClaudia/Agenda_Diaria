# 📅 Agenda Diaria Automatizada — Google Apps Script  

> Sistema corporativo para **MyCityHome** que centraliza calendarios, genera agendas diarias y las envía automáticamente por correo a cada empleado.  

---

## ✨ Funcionalidades principales
- 🗂️ **Lectura automática de empleados** desde Google Sheets.  
- 📅 **Consulta de Google Calendar corporativo** y calendarios adicionales por link.  
- 📧 **Envío diario de la agenda personalizada** a cada empleado (correo corporativo y personal).  
- ⏰ **Triggers automáticos** que ejecutan el envío todos los días a las 09:00 (con respaldo a las 09:05).  
- 🖥️ **Menú en la hoja de cálculo** para envíos manuales, crear/borrar triggers y pruebas rápidas.  
- 🌐 **Endpoint web (API)** para integrarse con el ERP o un widget externo.  

---

## ⚙️ Tecnologías usadas
- 🟨 Google Apps Script (JavaScript ES5).  
- 📊 Google Sheets (base de datos de empleados).  
- 📆 Google Calendar API (eventos corporativos).  
- 📧 GmailApp para envíos de correos automáticos.  

---

## 📂 Estructura del Proyecto
```

agenda-diaria-appscript/
│
├── main.gs       # Código completo del sistema
├── README.md     # Documentación (este archivo)

````

---

## 🔑 Funciones destacadas

| 🚀 Función | 📖 Descripción |
|------------|----------------|
| `onOpen()` | 📂 Añade un menú “Agenda diaria” en la hoja de Google Sheets. |
| `sendAllToday()` | 📧 Envía la agenda diaria a todos los empleados activos. |
| `sendSelectedRow()` | 👤 Envía la agenda solo al empleado de la fila seleccionada. |
| `setupDailyTriggerExact()` | ⏰ Crea un trigger exacto a las 09:00. |
| `setupDailyTriggerBackup()` | ⏰ Crea un trigger de respaldo a las 09:05. |
| `deleteAllTriggers()` | 🗑️ Borra todos los triggers activos del proyecto. |
| `doGet(e)` | 🌐 Endpoint web para integraciones (modo `dept`, `user`, `email`). |
| `renderPlannerHtml_()` | 🎨 Genera el HTML del correo con logo, branding y agenda del día. |

---

## 🚀 Cómo usarlo

1. Abre [Google Apps Script](https://script.google.com/).  
2. Crea un nuevo proyecto.  
3. Copia el contenido de `main.gs`.  
4. Ajusta la configuración en la sección **CONFIG**:
   - `SHEET_ID` → ID de tu Google Sheet con empleados.  
   - `SHEET_EMPLEADOS` / `SHEET_PRIORIDADES` / `SHEET_NOTAS`.  
   - `BRAND`, `LOGO_URL`, `HERO_BG` → Branding de tu empresa.  
   - `DOMAIN_FALLBACK` → Dominio por defecto para emails.  

5. Ve a tu Google Sheet, en el menú superior verás **Agenda diaria**.  
   - ▶️ *Enviar a todos (hoy)*  
   - ▶️ *Enviar fila seleccionada*  
   - ⚙️ *Crear trigger 09:00 exacto*  
   - 🗑️ *Borrar triggers*  

---

## 🌐 API REST (modo GET)

El script también expone un endpoint público que devuelve JSON:  

### 🔹 Agenda de un departamento
```http
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?token=XXX&mode=dept&dept=Atic
````

### 🔹 Agenda de un usuario

```http
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?token=XXX&mode=user&user=proyectostic@mycityhome.es
```

### 🔹 Agenda por email

```http
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?token=XXX&email=usuario@mycityhome.es
```

📦 Respuesta JSON ejemplo:

```json
{
  "ok": true,
  "mode": "user",
  "user": {
    "name": "Juan Pérez",
    "email": "juan.perez@mycityhome.es",
    "dept": "Atic"
  },
  "date": "2025-08-21",
  "tz": "Europe/Madrid",
  "events": [
    {
      "title": "Reunión equipo",
      "startISO": "2025-08-21T09:00:00Z",
      "endISO": "2025-08-21T10:00:00Z",
      "where": "Sala A",
      "allDay": false
    }
  ]
}
```

---

## 👀 Demo visual

### 📋 Ejemplo de correo de agenda

![Correo Agenda](assets/demo-agenda.png)

### 📊 Ejemplo de JSON de la API

![API JSON](assets/demo-json.png)

*(Crea screenshots de Gmail y del JSON para mostrar en tu repo)*

---

## ✨ Autor

👩‍💻 Creado por **\[Tu Nombre]**
⚡ Expert@ en automatización y productividad con Google Workspace.

---

### 🌟 ¿Por qué es un proyecto top?

✔️ Ahorra tiempo al equipo de RRHH y managers.
✔️ Centraliza calendarios dispersos en un solo email diario.
✔️ Integra un API lista para consumir en ERP o dashboards.
✔️ 100% sin servidores externos: todo corre en Google Workspace.


