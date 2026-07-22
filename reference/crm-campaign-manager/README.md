# Referencia — Campaign Manager (equipo CRM/Growth)

Copia de referencia del Apps Script **"Campaign Manager"** del equipo de CRM
(dueña del proyecto: Valentina Bobadilla; contacto operativo: Andrés Felipe
Cerinza), compartido a Legal el 2026-07-21/22 para integrarlo con RappiMind.

| Archivo | Qué es |
|---|---|
| `Code.gs` | Backend completo: router `doGet` (`?page=`), tickets Campañas (hoja `Base`, 55 col) y Rappicreditos, Slack (mensaje + hilo + revisor por país), CSV por ticket en Drive, `generateTicketId` (CAM-####). |
| `home.html` | **El formulario** ("Nueva Solicitud", 11 secciones) + página por defecto. |
| `status.html` | Consulta de ticket por ID o email. |
| `admin.html` | Panel del equipo CRM: stats, filtros, cambio de estado, Global Offer ID. |

⚠️ **Seguridad:**
- El `Code.gs` original traía un **token de bot de Slack hardcodeado** (línea 7).
  Aquí está **REDACTADO**. El token real quedó expuesto al compartirse el
  código → **hay que rotarlo** y moverlo a Script Properties. No lo re-agregues
  a este repo bajo ninguna circunstancia.
- IDs internos que sí se conservan (no son secretos, son configuración):
  spreadsheet `1yzcRTWhdVlm9G-M2--0KIS8_nnsGa93xg3xerL4ZQqc`, carpeta CSV
  `1pZ03_RgDTVtudaFDAmc8vK6-0k4qNLNR`, canal Slack `C09S1BGKUQJ`.

🚫 **Estos archivos NO se despliegan**: `.claspignore` es allowlist (ignora todo
salvo los 7 archivos del proyecto RappiMind), así que nada de `reference/` sube
al servidor de Apps Script con `clasp push`. Contiene un `doGet` propio que
**colisionaría** con el de `Código.js` si se subiera tal cual — la fusión se
hará renombrando y unificando el router (ver
`docs/INTEGRACION_CRM_CAMPAIGN_MANAGER.md`).
