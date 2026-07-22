# Integración RappiMind ⇄ Campaign Manager (CRM)

> Análisis del 2026-07-22, tras la reunión "Alineación nuevo flujo CRM <> Legal"
> (2026-07-21: Juan C. Gallego, Andrés F. Cerinza, Iván C. Patiño) y la revisión
> completa del código compartido por el equipo CRM (ver
> `reference/crm-campaign-manager/`). **Estado: propuesta — pendiente de
> decisiones (§6) antes de implementar.**

**Objetivo:** que la solicitud de oferta a CRM y la generación del T&C sean **un
solo proceso** con trazabilidad completa: `ticket CAM-#### ↔ T&C (Doc) ↔ link
publicado (Squarespace) ↔ Global Offer ID`.

---

## 1. Qué es el Campaign Manager (lo que ya tiene el equipo CRM)

Web app de Apps Script (proyecto de Valentina Bobadilla) con router por páginas
(`?page=form|status|admin`; `home.html` ES el formulario) y 2 módulos:

- **Campañas** — 7 tipos de oferta: `cashback`, `percentage`, `free_shipping`,
  `value`, `offer_by_product`, `service_fee` (+ selector `rappicreditos` que
  cambia de módulo). Form de 11 secciones: país/owner, tipo, fechas/horas,
  brand & stores (vertical, brand_id XOR store_ids, ciudades), descuento
  (%/valor, mín. compra, máx. descuento, máx. órdenes), segmentación,
  cobertura horaria (`days_hours`), esquema de pago (full_rappi / full_aliado →
  % alianza + ID alianza + orden de compra), config cashback (redención),
  config adicional (PRIME, métodos de pago, cc_type, BIN, **link T&C**) y
  budget/squad/strategy/descripción. Incluye "cargar campaña similar como base".
- **Rappicreditos** — carga de créditos a usuarios específicos: archivo
  user_id/monto (o ingreso manual), tope por pedido, vigencia (días o fecha),
  cobertura de redención, budget, pago/alianza.

**Flujo actual:** form → fila en su Sheet (`Base` 55 col / `RappiCreditos` 26
col, spreadsheet `1yzcRT…`) → **CSV por ticket** en Drive (`1pZ03…`, insumo para
montar la oferta en el CMS) → **Slack** (mensaje + hilo con detalle + mención al
revisor por país; `THREAD_TS` guardado para responder en hilo los cambios de
estado) → Andy monta la oferta → registra **Global Offer ID** (el admin exige
GO-ID para pasar a `Aprobado`) → estados `Pendiente → En Revisión → Aprobado →
Completado / Rechazado`.

**Datos clave para nosotros:**
- Ya existe la columna **`terminos_y_condiciones` (linkTyC)** — hoy manual y
  opcional. Es el punto de enganche natural del T&C.
- Sus **8 países** (`co,mx,ar,cl,pe,uy,ec,cr`) = los nuestros sin BR.
- Su form ya captura **casi todo lo que pide nuestro T&C de Cashback**.

## 2. Hallazgos técnicos (revisión del código)

| # | Hallazgo | Impacto |
|---|---|---|
| 1 | 🔴 **Token de bot de Slack hardcodeado** en `Code.gs:7` (quedó expuesto al compartir el código) | **Rotarlo ya** y moverlo a Script Properties. En nuestra copia de referencia está redactado. |
| 2 | `updateStatus`, `updateGlobalOffer`, `getAllTickets` **sin control de rol** | Mismo problema-clase que nuestras funciones admin; al fusionar, gatear con `_requireRole`/`Admin_Team`. |
| 3 | Su `doGet` + `getScriptUrl` + nombres genéricos (`getSheet`, `openSheet`) | **Colisionarían** con `Código.js` si se importan tal cual → renombrar/namespacear y unificar router. |
| 4 | `TIMEZONE = America/Mexico_City` | Nuestro proyecto usa hora de Bogotá; definir una sola para timestamps de tickets. |
| 5 | `generateTicketId` sin `LockService` | Carrera posible con envíos simultáneos (riesgo bajo; fácil de blindar al fusionar). |
| 6 | `setXFrameOptionsMode(ALLOWALL)` también en su `doGet` | Mismo pendiente de seguridad que ya tenemos. |
| 7 | Segmento **"Otro" = texto libre** | Al ticket CRM puede ir; **al T&C no** (regla legal: solo opciones pre-aprobadas). |
| 8 | No tenemos su `appsscript.json` | Confirmar scopes/`executeAs`/`access` de su despliegue. Los nuestros ya cubren lo necesario (Sheets, Drive, `external_request` para Slack, email). |

## 3. Mapeo de campos (ticket CRM → T&C Cashback)

Cobertura sorprendentemente alta — su form alimenta directo nuestro motor:

| Campaign Manager | RappiMind (canónico → placeholder/derived) | Nota |
|---|---|---|
| `country` (`co`) | `countryCode` (`CO`) | upper + guardarraíl A1 |
| `brand_name` / `brand_id` | `shopName` → `TIENDA_DISPLAY`/`REF_TIENDA`/`DEFINICION_TIENDA` | |
| `discount` | % cashback → `TEXTO_PORCENTAJE` | |
| `minimo_compra` | compra mínima → `UMBRAL_LETRAS/NUM` | |
| `maximo_descuento` | tope → `TOPE_LETRAS` | confirmar semántica: por usuario/por orden |
| `max_ordenes_usuario` | máx. órdenes → `TEXTO_ORDENES` | |
| `presupuesto` (budget) | presupuesto → `PRESUPUESTO_LETRAS` | |
| `segmentacion` | → `TEXTO_SEGMENTO` | ⚠️ "Otro" no puede fluir al T&C (§2.7) |
| `metodos_pago` + `cc_type`/`bin` | → `TEXTO_METODO_PAGO` | mapear valores (`cc`,`dc`,`cash`,`rappi_pay`…) |
| `prime` | segmento PRIME | |
| `fecha/hora_inicio/fin` | vigencia → `FECHA_*`/`HORA_*` (date_legal) | |
| `cashback_days_to_end` o fechas redención | → `TEXTO_VIGENCIA_CREDITOS` | |
| `stores_redencion_si/no`, `store_types_redencion` | → `TEXTO_LUGAR_REDENCION` | |
| `store_ids` / ciudades / cobertura | territorio → `TEXTO_TERRITORIO` | default "todo el territorio nacional" |
| `days_hours` | franjas horarias | hoy sin derived propio → condiciones especiales |
| `pago`, `pct_alianza`, `id_alianza`, `id_orden_compra` | — (solo ticket) | ⚠️ pregunta legal: ¿`full_aliado` cambia el Organizador? |
| `vertical`, `squad`, `strategy`, `descripcion` | — (solo ticket) | |
| `terminos_y_condiciones` (linkTyC) | ⬅ **aquí escribimos el link del T&C** | ver §5 (Doc vs Squarespace) |

**Lo que su form NO tiene y nuestro T&C sí pide:** Organizador (razón social +
ID fiscal — los campos de FASE C; `id_alianza` no es un NIT), condiciones
especiales, nombre interno de campaña (derivable de brand + ticket). El email
del solicitante sí lo capturan (sesión + owner).

## 4. Arquitectura propuesta (recomendación)

**Fusión en RappiMind** — un solo web app (la visión de la reunión):

1. Importar los 4 archivos renombrados: `Crm.gs.js` (backend, sin su `doGet`),
   `CrmForm.html`, `CrmStatus.html`, `CrmAdmin.html`.
2. **Router unificado** en `Código.js#doGet`: sin `?page` → generador RappiMind
   (`WebApp`); `?page=solicitud|status|admin-crm` → páginas CRM. `navTo()` y
   `getScriptUrl()` se conservan.
3. `SLACK_BOT_TOKEN` → **Script Property** (nunca en código). Canal/sheet
   parametrizados (dev/prod).
4. **Sección legal** dentro de `CrmForm.html` (visible cuando el tipo genera
   T&C): confirmaciones de responsabilidad + Organizador (si aplica) +
   condiciones especiales opcionales.
5. **Encadenamiento del envío** (tipo `cashback` + país habilitado por A1):
   `submitForm` → `processWebPayload` (motor RappiMind con guardarraíles A1/A2)
   → si OK → `saveRequest` con `linkTyC`/`TC_DOC_URL` = URL del Doc → CSV +
   Slack (el hilo de Slack sale ya **con el link del T&C**). Los demás tipos y
   países pasan como ticket normal sin T&C (nada se les rompe a marketing).
6. **Trazabilidad**: columna nueva `TC_DOC_URL` en su hoja (ADD-only al final,
   no rompe índices) + `ticketId` en nuestra `Respuestas_Audit_V2`. El campo
   `terminos_y_condiciones` queda para el **link público** (Squarespace) —
   conecta con P6 del roadmap.
7. Gates de rol al fusionar: `updateStatus`/`updateGlobalOffer` con
   `_requireRole` (equipo CRM entra a `Admin_Team` como rol propio o editor).

**Plan por fases:**
- **F0 — previos:** rotar token Slack (equipo CRM); copia del proyecto/Sheet de
  Vale para /dev; definir decisiones de §6.
- **F1:** importar + router + Script Properties, apuntando a Sheet/canal de
  pruebas. Su app de prod no se toca.
- **F2:** sección legal + encadenamiento T&C→ticket (Cashback CO).
- **F3:** pruebas e2e en /dev con Andy/Iván (incluye correr antes el runbook
  FASE C para que el form pida Organizador, y A2=on).
- **F4:** switch a producción (su Sheet + su canal Slack) coordinado con
  Valentina; su app vieja queda como respaldo/redirección.
- **F5:** extensiones — más países (según A1), Rappicreditos con T&C de
  créditos, leyendas (P5), tracking Squarespace (P6).

## 5. Decisiones de diseño abiertas (menores)

- **`linkTyC`: ¿Doc o link publicado?** Propuesta: `TC_DOC_URL` = Doc (auto,
  trazabilidad) y `terminos_y_condiciones` = URL pública de Squarespace,
  adjuntable después desde la consulta del ticket (P6).
- **Segmento "Otro":** al T&C va texto genérico aprobado (o se exige opción
  pre-aprobada); el texto libre solo queda en el ticket.
- **Timezone** única para timestamps (propuesta: America/Bogota).
- **Cashback fuera de CO en v1:** pasa como ticket sin T&C (marca "T&C
  pendiente — país no habilitado") en vez de bloquear.

## 6. Decisiones que necesita tomar Juan (bloqueantes para empezar)

1. **Arquitectura:** ¿Fusión en RappiMind (recomendado), o solo puente de datos
   (dos apps, RappiMind escribe el ticket), o montar copia /dev y decidir?
2. **Destino de los tickets en v1:** ¿Sheet copia /dev (recomendado) o directo
   al Sheet de producción del CRM? (Lo segundo exige acceso de edición +
   coordinación con Valentina desde el día 1.)
3. **Alcance v1 (viernes):** ¿Cashback CO con T&C auto + resto de tipos como
   ticket sin T&C (recomendado)? ¿Entra Rappicreditos? ¿Dejamos el mapeo
   multi-país listo?
4. **Orden del flujo:** ¿T&C primero y solo si sale bien se crea el ticket
   (recomendado, el "deber ser" de la reunión), ticket primero y T&C después, o
   en paralelo tolerante a fallos?

**Verificaciones externas (CRM / Legal):**
- ¿Ya está el grupo con Valentina y la copia del proyecto? ¿Tienes acceso al
  spreadsheet `1yzcRT…` y a la carpeta CSV `1pZ03…`?
- Avisar a Vale/Andy del **token de Slack expuesto** → rotar + Script Property.
- Legal: en `full_aliado`, ¿el Organizador del T&C sigue siendo Rappi o pasa a
  ser el aliado (y de dónde salen razón social + NIT)?
- Confirmar semántica exacta de `maximo_descuento` (por usuario vs por orden)
  para `TOPE_LETRAS`.
- Su `appsscript.json` (scopes/`executeAs`/`access`) para replicar el despliegue.
