# RappiMind — Motor Legal Rappi (Generador de T&C)

> Plataforma de **Legal Operations** de Rappi construida sobre **Google Apps
> Script**. Genera Términos & Condiciones de campañas promocionales (cashback,
> concursos) a partir de plantillas legales por país, con un panel de
> administración, flujo de aprobación de plantillas e importación asistida por
> **IA (Gemini)**.

**Estado:** Prototipo funcional en producción interna · Piloto multi-país en
preparación · Hoy operativo solo para **Colombia**.

- **Proyecto Apps Script:** https://script.google.com/home/projects/17LQqc40ukZJDLYaAexRZF7srT6Yb7JD7Fvjp9J8Y7mYh-L9qexptWy4V/edit
- **Script ID:** `17LQqc40ukZJDLYaAexRZF7srT6Yb7JD7Fvjp9J8Y7mYh-L9qexptWy4V`
- **Versión UI:** V2.3 · **Lógica interna:** hasta V3.4 ("GOD MODE")

---

## ⚠️ Avisos importantes

1. **El repo ya está sincronizado con el servidor.** Se hizo `clasp pull` y el
   repo contiene el código real y actual de Apps Script (`Código.js`,
   `Admin.gs.js`, `Config.gs.js`, `Helpers.gs.js`, `Setup.gs.js`, `WebApp.Html`,
   `appsscript.json`). A partir de aquí, **GitHub → Apps Script (`clasp push`) es
   seguro** → ver **[docs/DEPLOY_CLASP.md](docs/DEPLOY_CLASP.md)**.
2. **Multi-país: los datos están listos para 9 países, pero solo Colombia está
   sembrado y probado.** Ver [§7 Multi-país](#7-multi-país--localización-crítico-para-el-piloto).
3. **El chatbot y el feedback están a medias.** El frontend llama a `askGemini` y
   `saveFeedback`, pero esas funciones **no existen en el backend** → esos dos
   botones fallarían. (La generación de documentos sí funciona.)
4. **Hay funciones admin sin control de rol** y los T&C generados quedan públicos
   por enlace. Ver [§11 Seguridad](#11-seguridad-y-permisos).

---

## Tabla de contenido

1. [Qué es RappiMind](#1-qué-es-rappimind)
2. [Arquitectura](#2-arquitectura)
3. [Contenido del repositorio](#3-contenido-del-repositorio)
4. [Flujo del usuario final (generar un T&C)](#4-flujo-del-usuario-final-generar-un-tc)
5. [Panel de administración, roles y flujo de aprobación](#5-panel-de-administración-roles-y-flujo-de-aprobación)
6. [Modelo de datos](#6-modelo-de-datos)
7. [Multi-país / Localización](#7-multi-país--localización-crítico-para-el-piloto)
8. [Configuración y secretos](#8-configuración-y-secretos)
9. [Puesta en marcha (setup inicial)](#9-puesta-en-marcha-setup-inicial)
10. [Despliegue y sincronización (clasp)](#10-despliegue-y-sincronización-clasp)
11. [Seguridad y permisos](#11-seguridad-y-permisos)
12. [Problemas conocidos y riesgos](#12-problemas-conocidos-y-riesgos)
13. [Checklist pre-lanzamiento del piloto](#13-checklist-pre-lanzamiento-del-piloto)
14. [Roadmap y propuesta de rediseño](#14-roadmap-y-propuesta-de-rediseño)

---

## 1. Qué es RappiMind

RappiMind permite a los equipos comerciales/legales de Rappi **generar
documentos de Términos & Condiciones** para campañas promocionales en ~2 minutos,
sin redactar desde cero. El usuario elige país y tipo de campaña, llena un
formulario dinámico, y el motor rellena una **plantilla legal de Google Docs** con
los valores (fechas en letras, montos en palabras, jurisdicción, ley aplicable,
etc.) y entrega un Doc listo para publicar.

El backend vive en estos archivos de Apps Script (nombres tal como están en el
servidor): **`Código.js`** (core/motor + `doGet`), **`Admin.gs.js`** (panel admin
+ IA), **`Config.gs.js`** (configuración), **`Helpers.gs.js`** (utilidades) y
**`Setup.gs.js`** (instalación/migración). La base de datos es **un único
Spreadsheet**: `AUDIT_SHEET_ID = 1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI`.

> Nota: una versión vieja del repo incluía además un "Legal Team Tracker"
> (`Codigo.gs`, otro spreadsheet) que **no existe en este proyecto de Apps
> Script** — era código obsoleto y se eliminó al sincronizar con el servidor.

---

## 2. Arquitectura

```
┌─────────────────────────────────────────────────────────────────┐
│                      Google Apps Script (V8)                      │
│                                                                   │
│  doGet(e)  ──►  HtmlService.createTemplateFromFile('WebApp')       │
│                 (título "Motor Legal Rappi")                       │
│                                                                   │
│   Frontend (WebApp.Html, SPA)                                     │
│   ── google.script.run ──►  Funciones servidor (.js)              │
│                                                                   │
│   Backend:                                                        │
│     • Código.js     → core/motor: doGet, processWebPayload,      │
│                       getCampaignTypesForUser, getFieldsForUserForm│
│     • Admin.gs.js   → panel admin, roles, workflow, IA Gemini     │
│     • Config.gs.js  → constantes, mapas país, DERIVED_FIELDS      │
│     • Helpers.gs.js → utilidades (fechas ES, números a letras…)   │
│     • Setup.gs.js   → instalación/seed/migración                  │
└───────────────┬───────────────────────┬─────────────────────────┘
                │                       │
        ┌───────▼───────┐       ┌───────▼────────┐      ┌──────────────┐
        │ Google Sheets │       │  Google Drive  │      │  Gemini API  │
        │ (1 spreadsheet)│      │ plantillas/Docs│      │ (UrlFetchApp)│
        └───────────────┘       └────────────────┘      └──────────────┘
```

- **Frontend:** SPA de una sola página (`WebApp.Html`), navegación por
  mostrar/ocultar `<div>` (sin router). Tailwind (CDN), Font Awesome 6.4.0,
  fuente Nunito. Todo en **español**.
- **Backend:** funciones Apps Script invocadas vía `google.script.run`
  (devuelven strings JSON). Persisten en Sheets, generan Docs y los mueven a
  carpetas de Drive por país.
- **IA:** importación de plantillas y chatbot vía **Gemini**
  (`gemini-flash-latest`, llamado con `UrlFetchApp`, key en Script Property).
- **Sin build, sin tests, sin dependencias npm en runtime** (npm aquí es solo
  para clasp).

---

## 3. Contenido del repositorio

| Archivo | Líneas aprox. | Descripción |
|---|---|---|
| `Código.js` | ~745 | **Core/motor.** `doGet` (sirve `WebApp`), `processWebPayload` (genera el documento), `getCampaignTypesForUser`, `getFieldsForUserForm`, `coreEngineV2`, `mapWebToEngine`. |
| `Admin.gs.js` | ~1156 | **Panel admin**: equipo/roles, workflow de plantillas, carpetas Drive, e **IA Gemini** (`analyzeTextForPlaceholders`, `createTemplateFromWizard`, `fetchGoogleDocContent`). |
| `Config.gs.js` | ~438 | Constantes globales, `TW_CONFIG`, mapas país (`COUNTRY_FOLDERS`), `DERIVED_FIELDS` (~50 funciones de placeholders), `LEGAL_DEFAULTS_MAP`, `MESES_ES`. |
| `Setup.gs.js` | ~1276 | Funciones de **instalación/seed/migración** (crear hojas, plantillas CO, carpetas, `migrateV33`…). |
| `Helpers.gs.js` | ~152 | Utilidades: fechas/horas en español, `numeroALetras`, `_getOrCreateSheet`, `_sheetToObjects`, `setPublicViewPermissions`, etc. |
| `WebApp.Html` | ~5192 | **SPA del generador de T&C** (UI usuario + panel admin + wizard IA). Título "RappiMind \| Generador de T&C". |
| `Propuesta de ajuste del front` | ~4146 | **Rediseño propuesto (NO desplegar).** Borrador con 2 errores de sintaxis JS que lo dejan no funcional. Ver [§14](#14-roadmap-y-propuesta-de-rediseño). |
| `appsscript.json` | — | Manifiesto real del servidor (timezone, runtime V8, web app `executeAs: USER_DEPLOYING` / `access: DOMAIN`). |
| `.clasp.json`, `.claspignore`, `package.json`, `.gitignore` | — | Tooling de sincronización clasp. |
| `.github/workflows/clasp-push.yml` | — | Push automático a Apps Script (se activa al configurar el secret `CLASPRC_JSON`). |
| `docs/DEPLOY_CLASP.md` | — | **Runbook de sincronización GitHub ⇄ Apps Script.** |

---

## 4. Flujo del usuario final (generar un T&C)

1. **Landing** (`#landingPage`) → "Comenzar Ahora".
2. **Modal de responsabilidad** (`#responsibilityModal`): 3 checkboxes
   obligatorios. ⚠️ Hoy uno dice *"Actualmente solo disponible para Colombia"* —
   contradice el piloto multi-país (corregir).
3. **Selector de país** (`#countrySelector`): tarjetas por país. Al elegir,
   `selectCountry(code, name, currency, symbol, flag)` fija moneda/símbolo y abre
   el formulario.
4. **Formulario** (`#mainForm`, 4 secciones con barra de progreso):
   - **1. Información:** email (`@rappi.com`), tipo de campaña (tarjetas cargadas
     del backend), nombre interno.
   - **2. Configuración:** comercio aliado + territorio, fechas y horas de
     vigencia.
   - **3. Variables** (dinámicas, "GOD MODE"): campos según el tipo de campaña,
     traídos de `getFieldsForUserForm(typeName, country)`. Soporta dependencias
     condicionales, tooltips, validaciones.
   - **4. Restricciones** (opcional): compra mínima, máx. órdenes, medios de pago,
     segmento de usuario, CC, condiciones especiales.
   - Autosave del borrador a `localStorage` (debounce 500 ms, TTL 24 h) + vista
     previa en vivo (desktop).
5. **"GENERAR DOCUMENTO"** → `handleFormSubmit` valida; si hay premio físico
   entregado por el organizador, exige el modal de *Acuerdo de Transferencia de
   Datos* (Ley 1581 CO). Luego `submitForm()` arma el `payload` y llama
   **`google.script.run.processWebPayload(JSON.stringify(payload))`**.
6. **Éxito:** modal con enlace al Doc generado (`#docLink`), instrucciones de
   publicación en Squarespace, captura de feedback. Se añade al historial local.
   **Error:** `alert()` con el mensaje del servidor.
7. Extras: **chatbot Gemini** (`askGemini`), historial local (`localStorage`),
   botón Demo (`fillTestValues`).

> El generador corre del lado del servidor con la identidad del despliegue y
> comparte el Doc resultante (`setPublicViewPermissions` → "cualquiera con el
> enlace puede ver"). Revisar implicaciones de privacidad ([§11](#11-seguridad-y-permisos)).

---

## 5. Panel de administración, roles y flujo de aprobación

### Roles (jerarquía)

`ROLE_HIERARCHY = { owner: 4, admin: 3, editor: 2, viewer: 1 }`

| Capacidad | Rol mínimo |
|---|---|
| Ver panel admin | viewer |
| Crear/editar plantillas y campos | editor |
| Aprobar / activar / eliminar plantillas, gestionar carpetas | admin |
| Gestionar equipo (alta/baja/roles) | owner |

- **Fuente de roles:** hoja `Admin_Team` (solo filas `status='active'`). El gate
  real es `_requireRole(minRole)`.
- `ADMIN_EMAILS_LIST = ['juan.gallego@rappi.com']` es el allowlist de bootstrap,
  pero **solo lo respeta `getUserRole`, no `_requireRole`**. ⚠️ El primer `owner`
  debe existir como fila activa en `Admin_Team` (ver [§12](#12-problemas-conocidos-y-riesgos)).

### Workflow de plantillas (estado en `Template_Registry.status`)

```
(nueva) ──► draft ──submit──► pending_review ──► approved ──► active
                                   │
                                   └──► rejected
   (admin+ puede crear directo en 'active')   (active ⇄ inactive vía toggle)
```

- `editor` crea borradores y los envía a revisión; `admin` aprueba/rechaza/activa.
- Al activarse una plantilla, `_ensureCampaignTypeActive()` activa también su
  tipo de campaña (para que sea visible a usuarios comerciales).
- ⚠️ **No hay notificaciones reales**: `_notifyAdmins` es un *stub* (solo
  `Logger.log`). El revisor no se entera de los envíos.

### Wizard de importación con IA (Gemini)

4 pasos: **Cargar** un Google Doc (`fetchGoogleDocContent`) → **Analizar**
(`analyzeTextForPlaceholders`, detecta variables `{{...}}` con Gemini, 2–5 min) →
**Confirmar** placeholders (editar/filtrar por confianza) → **Guardar**
(`createTemplateFromWizard` crea el Doc-plantilla, lo mueve a la carpeta del país
y registra campos en `Template_Fields`).

### Funciones admin (`google.script.run`)

`adminGetCurrentUser`, `adminGetTeam`, `adminAddTeamMember`,
`adminUpdateTeamMember`, `adminRemoveTeamMember`, `adminGetFolderStructure`,
`adminGetTemplates`, `adminSaveTemplate`, `adminSubmitForReview`,
`adminApproveTemplate`, `adminRejectTemplate`, `adminToggleTemplate`,
`adminDeleteTemplate`, `adminGetFields`, `adminSaveField`, `adminDeleteField`,
`adminGetCampaignTypes`, `adminGetCountrySettings`, `adminGetLogs`,
`adminGetApprovalLog`, `getUserRole`, `analyzeTextForPlaceholders`,
`createTemplateFromWizard`, `fetchGoogleDocContent`.

> Todas devuelven **string JSON** → el cliente hace `JSON.parse`. (Hay dos
> convenciones de respuesta y `getUserRole` devuelve un string plano; ver
> [§12](#12-problemas-conocidos-y-riesgos)).

---

## 6. Modelo de datos

### Spreadsheet A — DB del Generador (`AUDIT_SHEET_ID = 1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI`)

| Hoja | Para qué | Columnas (clave) |
|---|---|---|
| `Template_Registry` | Catálogo de plantillas (1 fila por país × tipo) | `country_code, country_name, campaign_type, template_doc_id, version, status, currency_code, currency_symbol, legal_owner, last_updated, notes, submitted_by, submitted_date, approved_by, approved_date, rejected_by, vertical` |
| `Template_Fields` | **Esquema dinámico** del formulario | `field_id, country_code, campaign_type, placeholder, label_es, field_type, icon, required, section, validation_rule, options, default_value, tooltip, depends_on, order, group, format_as, canonical_field_id` |
| `Admin_Team` | Usuarios + roles | `email, name, role, added_by, added_date, status, notes` |
| `Approval_Log` | Auditoría de acciones | `timestamp, actor, action, details` |
| `Campaign_Types` | Catálogo de dinámicas de campaña | `type_id, type_name, description, parent_type, processing_mode, icon, color, status, countries, created_by, created_date` |
| `Country_Settings` | **Config legal/moneda por país** | `country_code, country_name, legal_country, currency_name, currency_code, currency_symbol, legal_entity, timezone, jurisdiction_text, applicable_law, legal_url` (+ columnas de `LEGAL_DEFAULTS_MAP` que hoy **no crea ningún setup**) |
| `Respuestas_Audit_V2` | Auditoría de generación | `timestamp, email, docUrl, type, country, shop` |

Plantillas (Google Docs) creadas por el setup: `Template_CO_Cashback`,
`Template_CO_Concurso` (prosa legal colombiana con tokens `{{PLACEHOLDER}}`).

> El "Legal Team Tracker" (otro spreadsheet `19eR-…` y un `Codigo.gs` con hojas
> `Tracking Activo`/`Historial`/`Proyectos`/`Equipos`) que aparecía en versiones
> viejas del repo **no pertenece a este proyecto de Apps Script** y se eliminó al
> sincronizar. Si ese tracker sigue vivo, está en otro proyecto/Script ID aparte.

---

## 7. Multi-país / Localización (CRÍTICO para el piloto)

**Estado: arquitectura para 9 países, pero solo Colombia operativo.**

- **Países en datos/UI:** `CO, MX, PE, AR, CL, EC, UY, CR, BR` (9).
  - Selector de usuario: 8 activos + **Brasil deshabilitado** ("Requiere template
    PT").
  - **Perú (PE) es inconsistente:** seleccionable para el usuario, pero falta en
    varios mapas del panel admin (`FLAG_MAP`, `COUNTRY_NAMES`) y en el loop de
    `setupTemplateFolders()` (que solo crea 8 carpetas, sin PE ni AR).
  - La landing dice "8 Países"; el selector muestra 9. Unificar.
- **Solo CO está sembrado:** plantillas (`Template_CO_*`), campos
  (`seedColombiaFields`) y las columnas legales de `Country_Settings`
  (`jurisdiction_text`, `applicable_law`, `legal_url`) **solo se llenan para CO**.
- **Idioma:** **solo español** en todo (`MESES_ES`, `label_es`, prosa de las
  plantillas, `DERIVED_FIELDS`). **Brasil necesitaría portugués, que no existe.**
- **Supuestos colombianos incrustados en código/plantillas:**
  - Locale `'es-CO'` hardcodeado en formato de números (`toLocaleString('es-CO')`).
  - `numeroALetras` solo en español; `LISTA_MUNICIPIOS` solo de Colombia.
  - Moneda literal "pesos M/Cte" / "$" en la prosa de las plantillas (ignora
    `currency_symbol` de `Country_Settings`).
  - Ley de gobierno "leyes de **Colombia**", **Ley 1581 de 2012**, **SIC**,
    **RAPPI S.A.S.**, URLs `legal.rappi.com.co/colombia/…` hardcodeadas en los Docs.
  - `NIT_ORGANIZADOR` (NIT es ID fiscal colombiano; otros países usan RFC/RUC/CUIT/CNPJ).

### Onboarding de un país nuevo (p. ej. PE)

1. **`Country_Settings`:** completar la fila del país, incluidas las columnas
   legales (`jurisdiction_text`, `applicable_law`, `legal_url` + las de
   `LEGAL_DEFAULTS_MAP`, que hay que **crear a mano**).
2. **`TW_CONFIG.COUNTRY_FOLDERS`** (`Config.gs`): asegurar el código (PE ya está).
3. **Carpetas Drive:** añadir el país al array de `setupTemplateFolders()` (faltan
   PE y AR) y re-ejecutar, o crearlas a mano.
4. **Redactar los Google Docs** de plantilla del país (Cashback/Concurso) con la
   ley local, autoridad de datos, entidad legal, URLs y moneda correctas —
   **y quitar los literales colombianos** que no leen de `Country_Settings`.
5. **`Template_Registry`:** añadir filas `active` del país apuntando a esos Docs,
   con `currency_code`/`currency_symbol` correctos.
6. **`Template_Fields`:** los campos sembrados son `country_code='ALL'` (aplican
   solo); verificar etiqueta de ID fiscal y opciones específicas.
7. **Cambios de código para correctitud** (no solo datos): generalizar el locale
   `'es-CO'`, `numeroALetras`, lógica de territorio, y (para BR) soporte de
   portugués. **Nada de esto existe hoy.**

---

## 8. Configuración y secretos

### Script Properties (Proyecto → Configuración del proyecto → Propiedades del script)

| Propiedad | Uso | Quién la crea |
|---|---|---|
| `TEMPLATES_FOLDER_ID` | Carpeta raíz de plantillas en Drive | `setupTemplateFolders()` |
| `GEMINI_API_KEY` | **API key de Gemini** para IA (importación + chatbot) | **Manual** (ningún setup la crea) |

### Constantes clave (`Config.gs`)

| Constante | Valor | Nota |
|---|---|---|
| `ADMIN_EMAILS_LIST` | `['juan.gallego@rappi.com']` | Owner de bootstrap (único, hardcoded). |
| `LEGAL_AUDIT_EMAIL` | `['juan.gallego@rappi.com']` | Definido pero **no usado**. |
| `AUDIT_SHEET_ID` | `1Ki9FvHGk…LixqZI` | Spreadsheet A (DB del generador). |
| `DRIVE_FOLDER_ID` | `''` | **Vacío/muerto**; el real está en `TEMPLATES_FOLDER_ID`. |
| `TEMPLATES_ROOT_NAME` | `'RappiMind_Templates'` | Nombre de la carpeta raíz. |

> 🔐 **`GEMINI_API_KEY` es un secreto vivo**: está (correctamente) en Script
> Properties, no en el código. No lo subas a git ni al README. Rótalo si se
> expone. ⚠️ `callGeminiForAnalysis` hace `Logger.log` de la respuesta cruda de
> Gemini (posible fuga en logs — revisar).

---

## 9. Puesta en marcha (setup inicial)

> Ejecutar desde el editor de Apps Script. **Solo en un entorno nuevo.** Varias
> funciones **NO son idempotentes** (ver aviso abajo).

1. `setupTemplateEngine()` — crea plantillas CO + registro + campos base.
   **Ejecutar EXACTAMENTE UNA VEZ.**
2. `setupAdminSystem()` — crea `Admin_Team` (te siembra como `owner`),
   `Approval_Log`, carpetas Drive y guarda `TEMPLATES_FOLDER_ID`.
3. `setupCampaignTypes()`
4. `setupCountrySettings()` — siembra los 9 países (columnas base).
5. `upgradeTemplateFieldsFormatting()`
6. `upgradeRegistryVertical()`
7. `seedMissingFields()`
8. `migrateV33()` y luego `verifyV33()` (chequeo de salud).
9. **Manual:** configurar `GEMINI_API_KEY` en Script Properties; aplicar el edit
   pendiente a `processWebPayload` que indica el log de `setupTemplateEngine`.

> ⚠️ **`setupTemplateEngine()` / `_seedRegistry` NO son idempotentes**:
> re-ejecutarlos crea Docs y filas de registro **duplicados**. `migrateV33` y los
> `setup*` con guardas sí son seguros de re-correr.

---

## 10. Despliegue y sincronización (clasp)

La sincronización **GitHub ⇄ Apps Script** se hace con `clasp`. Toda la guía
operativa (login, pull-first, push, GitHub Action, troubleshooting) está en:

➡️ **[docs/DEPLOY_CLASP.md](docs/DEPLOY_CLASP.md)**

Resumen:

```bash
npm install                 # instala clasp (fijado en package.json)
npx clasp login             # autenticación OAuth (token en ~/.clasprc.json)
npx clasp status            # verifica el scriptId

npx clasp pull              # traer cambios Apps Script → repo (luego commit + git push)
npm run push:force          # subir cambios repo → Apps Script (clasp push --force)
```

- **Modelo:** GitHub → Apps Script (push), con **GitHub Action** automático en
  merge a `main` (`.github/workflows/clasp-push.yml`). El repo ya está
  reconciliado con el servidor, así que el push es seguro.
- Para activar el automático: carga el secret `CLASPRC_JSON` en el repo (ver el
  runbook). Mientras no exista el secret, el Action falla sin tocar nada.

---

## 11. Seguridad y permisos

- **Scopes OAuth** (en `appsscript.json`): `spreadsheets`, `drive`, `documents`,
  `script.external_request` (Gemini), `userinfo.email`.
- **Despliegue web app:** `executeAs: USER_DEPLOYING`, `access: DOMAIN`
  (solo `@rappi.com`). Coherente con la identificación del visitante por
  `Session.getActiveUser().getEmail()`. **Verificar contra la config real tras el
  pull.**
- ⚠️ **`setXFrameOptionsMode(ALLOWALL)`** en el `doGet`: el dashboard es
  embebible en cualquier iframe (superficie de clickjacking). El endpoint
  `?page=api` expone todos los datos como JSON.
- ⚠️ **Docs públicos:** `setPublicViewPermissions` deja los T&C generados como
  "cualquiera con el enlace puede ver". Validar con Legal/Privacidad.
- ⚠️ **Funciones admin sin control de rol:** `adminToggleTemplate`,
  `adminDeleteTemplate`, `analyzeTextForPlaceholders`, `fetchGoogleDocContent`
  **no llaman `_requireRole`** → cualquier usuario autenticado que las invoque
  puede activar/borrar plantillas o leer Docs por ID. **Corregir antes del piloto.**
- **PII:** nombres/emails de empleados en `getDefaultEquipos` (Tracker) y en
  `Admin_Team`.

---

## 12. Problemas conocidos y riesgos

**Backend / arquitectura**
- **Chatbot y feedback rotos:** el frontend llama a `askGemini` y `saveFeedback`,
  pero esas funciones **no están definidas en el backend** → fallan en runtime.
- **Nombres de archivo del servidor poco limpios** (`Admin.gs.js`, `Setup.gs.js`,
  `Código.js` con acento): funcionan, pero conviene normalizarlos a futuro.
- **Sin tests ni ambientes (dev/prod) separados;** un único Spreadsheet de datos.

**Multi-país**
- Solo CO sembrado y probado; supuestos colombianos hardcodeados; sin portugués
  para BR; PE/AR inconsistentes (carpetas y mapas admin).

**Seguridad** — ver [§11](#11-seguridad-y-permisos) (funciones admin sin gate,
docs públicos, ALLOWALL).

**Datos / correctitud**
- `setupTemplateEngine`/`_seedRegistry` **no idempotentes** (duplican).
- `ADMIN_EMAILS_LIST` vs `Admin_Team`: el primer owner debe estar como fila
  activa en `Admin_Team` o no podrá hacer acciones de owner.
- `adminRemoveTeamMember`/`adminUpdateTeamMember` **no revocan** el acceso a la
  carpeta de Drive ya compartida.
- `adminRejectTemplate` registra el motivo pero **no lo guarda** en la fila.
- Direccionamiento de filas **posicional** (`index+2`) → ediciones concurrentes
  pueden afectar la fila equivocada (sin locking).
- `LEGAL_DEFAULTS_MAP` referencia columnas de `Country_Settings` que **ningún
  setup crea** → valores legales en blanco hasta crearlas a mano.
- `buildAnalysisPrompt` **trunca el texto a 15.000 caracteres** para la IA.
- Varios `catch {}` **silencian errores** (sharing, move, logs).

**Frontend**
- La pestaña **"Actividad" (logs) del admin lanza error**: el JS referencia
  `#panel-logs` pero ese bloque HTML **no existe** en `WebApp.Html`.
- Copy contradictorio: modal de responsabilidad dice "solo Colombia"; landing
  dice "8 Países" vs 9 en el selector.
- Manejo de errores por `alert()` (aceptable para piloto interno).

---

## 13. Checklist pre-lanzamiento del piloto

**Bloqueantes (verificar en el proyecto DESPLEGADO):**
- [ ] Confirmar que en Apps Script existen y funcionan `processWebPayload`,
      `getCampaignTypesForUser`, `getFieldsForUserForm`, `askGemini`,
      `saveFeedback`. (No están en el repo.)
- [ ] Confirmar cuál `doGet`/HTML sirve realmente el **generador** (el del repo
      sirve el Tracker).
- [ ] **Una generación end-to-end por país piloto**: país → tipo → campos →
      generar → abrir Doc → verificar moneda, fechas y prosa legal correctas.
- [ ] `Campaign_Types`, `Template_Fields`, `Country_Settings` sembrados para
      **cada** país piloto (no solo CO).

**Despliegue / acceso / seguridad:**
- [ ] Web app con `executeAs` correcto y `access` restringido al dominio Rappi
      (no "Cualquiera").
- [ ] Sembrar `Admin_Team` con los emails/roles correctos (incluido el primer
      owner como fila activa).
- [ ] Añadir `_requireRole` a `adminToggleTemplate`, `adminDeleteTemplate`,
      `analyzeTextForPlaceholders`, `fetchGoogleDocContent`.
- [ ] Auditar `setPublicViewPermissions` / `_shareFolderWithMember` (no
      sobre-compartir).
- [ ] `GEMINI_API_KEY` en Script Properties, con cuota suficiente; quitar el
      `Logger.log` de la respuesta cruda.

**Aislamiento de datos:**
- [ ] Verificar el ruteo de carpetas por país (`_moveTemplateToFolder`): que la
      campaña del país A caiga en la carpeta del país A.
- [ ] Confirmar que los IDs hardcodeados apuntan a **producción**, no a copias.

**UX / resiliencia:**
- [ ] Guarda explícita "no hay plantilla para este país" (hoy cae a CO/COP en
      silencio — riesgoso para un documento legal).
- [ ] Corregir copy multi-país (modal de responsabilidad, contador de países).
- [ ] Arreglar la pestaña "Actividad" del admin (`#panel-logs` ausente).

---

## 14. Roadmap y propuesta de rediseño

El archivo **`Propuesta de ajuste del front`** es un rediseño visual (tema
oscuro, "v2.4", nuevas pantallas: hub de acciones, "T&C Generales", flujo de
modificar campaña, tour guiado). **No desplegarlo todavía:**

- Tiene **2 errores de sintaxis JS** que dejan los `<script>` inertes (un
  `toggle(` sin cerrar en `fillDemo()` y un `else` suelto); un bloque admin quedó
  pegado en su propio `<script>` sin fusionar.
- **Elimina el motor dinámico** (`getCampaignTypesForUser`/`getFieldsForUserForm`)
  y hardcodea los formularios Cashback/Concurso → los tipos/campos creados por el
  admin **dejarían de aparecer** al usuario.
- Vuelve a hardcodear Colombia/COP en las pantallas de usuario.

**Recomendación:** mantener `WebApp.Html` para el piloto; tratar la propuesta
como spike de diseño v-next (arreglar sintaxis, fusionar admin, reintroducir el
motor dinámico, y solo entonces probar).

### Mejoras sugeridas (post-piloto)

- Internacionalización real (i18n) + portugués para BR.
- Sacar la prosa colombiana hardcodeada de las plantillas hacia `Country_Settings`.
- Notificaciones reales de revisión (`_notifyAdmins`).
- Idempotencia en el setup; locking/transacciones en escrituras de Sheets.
- Pruebas end-to-end por país y manejo de errores no basado en `alert()`.

---

*Documentación generada a partir de una revisión exhaustiva del código
(backend, admin/workflows, configuración/datos, frontend y análisis de brechas).*
