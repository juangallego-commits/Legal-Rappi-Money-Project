# Revisión integral — RappiMind (Motor Legal Rappi)

> Corte: **2026-07-22** · Rama: `claude/revision-integral-proceso` · Basada en el
> estado **real del código** en el repo (7 archivos sincronizados con Apps Script
> vía clasp) + `README.md` + runbooks. Verificado con la suite de tests y lectura
> directa del backend.

**TL;DR** — El proyecto está **más maduro** de lo que sugería el `README.md`
anterior: la tanda de trabajo **FASE A–D** (ya en `main`) cerró varios riesgos
legales/estructurales y añadió **tests (31/31 en verde)**. Sigue siendo un
**piloto solo-Colombia**. Lo que falta para operar Cashback con "Organizador" es
**operacional** (correr un runbook y encender una propiedad), no código. Antes de
abrir el piloto multi-país quedan pendientes de **seguridad** (5 funciones admin
sin control de rol) y **2 features rotas** (chatbot y feedback llaman a un backend
inexistente).

---

## 1. Objetivos del proyecto

- **Producto:** RappiMind / *Motor Legal Rappi*. Permite a equipos comerciales
  (KAMs) y legales **generar T&C de campañas promocionales** (Cashback, Concurso)
  en ~2 minutos, rellenando plantillas legales de Google Docs por país —sin
  redactar desde cero— y entregando un Doc listo para publicar (Squarespace /
  promos.rappi.com).
- **Alcance objetivo:** multi-país LATAM (9 países en la arquitectura), panel de
  administración con **roles y workflow de aprobación** de plantillas, e
  **importación asistida por IA (Gemini)**.
- **Principio rector (legal):** el KAM **no escribe texto libre** en campos
  sensibles; todo por *dropdowns/chips* pre-aprobados. El formateo legal complejo
  (números a letras, jurisdicción, ley aplicable, montos) lo resuelve el motor de
  forma determinista.
- **Norte de producto (roadmap V4):** selector de vertical (Delivery vs **Travel**,
  entidad legal distinta), limpieza del panel admin, **banco de campos**
  reutilizables, UX pregunta-por-pregunta, **leyendas** post-generación para piezas
  de marketing, y tracking del link de Squarespace.

---

## 2. Arquitectura y estado del repositorio

- **Plataforma:** Google Apps Script (runtime V8). Frontend SPA (`WebApp.Html`) +
  backend en `.gs`, invocado por `google.script.run`. Base de datos = **un único
  Spreadsheet** (`1Ki9FvHGk…LixqZI`). IA por `UrlFetchApp` → Gemini.
- **7 archivos reales**, sincronizados con el servidor vía **clasp** (GitHub =
  fuente de verdad; push automático a Apps Script en merge a `main` una vez cargado
  el secret `CLASPRC_JSON`):

  | Archivo | Rol |
  |---|---|
  | `Código.js` | Core/motor: `doGet`, `processWebPayload`, `coreEngineV2`, generación, guardarraíles FASE A1/A2/B/D |
  | `Admin.gs.js` | Panel admin, roles, workflow, IA Gemini, writer FASE C |
  | `Config.gs.js` | Constantes, `DERIVED_FIELDS`, `LEGAL_DEFAULTS_MAP`, `FIELD_CATALOG`, mapas país |
  | `Helpers.gs.js` | Utilidades (fechas/números ES, sheets, permisos) — **limpio** |
  | `Setup.gs.js` | Instalación/seed/migración |
  | `WebApp.Html` | SPA (usuario + admin + wizard) |
  | `appsscript.json` | Manifiesto (scopes, web app) |

- **`doGet` sirve `WebApp`** con título "Motor Legal Rappi" (el generador). El
  "Legal Team Tracker" obsoleto fue eliminado en la reconciliación con el servidor.
- **`Propuesta de ajuste del front`** (rediseño v2.4): **NO desplegar** — tiene
  errores de sintaxis y elimina el motor dinámico. Está excluido de `clasp push`
  (`.claspignore`), así que es seguro que viva en el repo. Tratar como spike de
  diseño.

> ⚠️ **Límite de esta revisión:** los **datos sembrados** (filas de
> `Template_Registry`, `Country_Settings`, `Template_Fields`, `Admin_Team`) viven
> en el Spreadsheet, **no en git**. El estado por país de §5 se basa en la
> documentación/skill, no en inspección directa de la hoja — **verificar en la hoja
> antes del piloto.**

---

## 3. Lo que se entregó recientemente (FASE A–D) — ya en `main`

Progreso real desde la última foto del README. **Todo cubierto por tests.**

| Fase | Qué hace | Estado |
|---|---|---|
| **A1 — Guardarraíl de país** | `_validateCountryLegal()` aborta la generación **antes** de crear el Doc si el país no tiene fila en `Country_Settings`, si falta una columna legal requerida, o si hay marcadores `[VERIFICAR`. Lanza *"País no habilitado… Contacta a Legal"*. **Elimina el peligroso fallback silencioso a CO/COP.** | ✅ Cableado en `coreEngineV2` |
| **A2 — Validación pre-entrega** | `A2_ABORT` si el documento final conserva marcadores `{{...}}` o `[ ... ]` sin resolver → **no publica un T&C inválido**. Gated por Script Property `RAPPIMIND_A2='on'`. | ⚙️ Codificado, **apagado por defecto** |
| **B — Bloques opcionales** | Sentinelas `[[?TOKEN]]…[[/?]]`: si el token va vacío borra el bloque **completo** (evita restos tipo *"por , identificada con ."*). Determinista, sin heurística. | ✅ |
| **C — Catálogo de campos + writer** | `FIELD_CATALOG` determinista + `previewFieldDerivation` (dry-run) / `applyFieldDerivation` (**ADD-only, UPSERT idempotente** por llave `country+type+placeholder`). Puebla los 2 campos de **"Organizador"** (razón social + ID fiscal) en Cashback **sin tocar** filas `ALL` ni `Concurso`. | ⚙️ Codificado, **no ejecutado en prod** |
| **D — Contrato de derivados** | Normaliza salida (capitaliza, agrega punto, idempotente, respeta `? !`). | ✅ |
| **Fix UI** | El form agrupa **"Organizador" primero** (evita headers duplicados/intercalado). | ✅ |

**Tests:** `npm test` → **31/31 en verde** (catálogo de campos, scope/idempotencia
del writer, bloques opcionales, contrato de derivados, detección de residuales A2).
El repo pasó de *"sin tests"* a tener una red de seguridad real para la lógica
legal-sensible.

---

## 4. Temas pendientes y riesgos (verificado en código)

### 🔴 Bloqueantes de seguridad (antes del piloto)

1. **5 funciones admin sin control de rol** — no llaman `_requireRole`, así que
   cualquier usuario `@rappi.com` autenticado que las invoque por
   `google.script.run` podría ejecutarlas:
   - `adminToggleTemplate(index, status)` → **activa/desactiva** cualquier plantilla.
   - `adminDeleteTemplate(index)` → **borra** una fila del registro (`deleteRow`).
   - `analyzeTextForPlaceholders`, `fetchGoogleDocContent`, `createTemplateFromWizard`
     (wizard IA) → leen/analizan Docs por ID y crean plantillas.
   - *(Nota: `previewFieldDerivation`/`applyFieldDerivation` tampoco gatean, pero
     son de uso desde el editor, no desde la UI → riesgo menor.)*
   → **Fix:** añadir `_requireRole('admin')` (o `'editor'`) al inicio de cada una.

2. **Docs T&C públicos por enlace** — `setPublicViewPermissions` deja cada T&C
   generado como *"cualquiera con el enlace puede ver"*. **Validar con
   Privacidad/Legal** la política de compartición.

3. **`setXFrameOptionsMode(ALLOWALL)`** en `doGet` — el dashboard es embebible en
   cualquier iframe (superficie de clickjacking). Evaluar restringir.

### 🟠 Features rotas (visibles al usuario)

4. **Chatbot y feedback caídos** — `WebApp.Html` llama a `askGemini` y
   `saveFeedback`, pero **ninguna existe en el backend** (confirmado: no están en
   ningún `.gs`). Esos dos botones fallan en runtime. → Implementar el backend o
   **retirar los botones** para el piloto. *(La generación de documentos sí
   funciona.)*

### 🟡 Correctitud / operación

5. **FASE C/A2 no ejecutadas en prod** — el form aún **no pide "Organizador"** en
   Cashback hasta correr `applyFieldDerivation`, y A2 no valida hasta encender
   `RAPPIMIND_A2='on'`. Es el **próximo paso operativo** (ver §6, runbook FASE C).
6. **Setup no idempotente** — `setupTemplateEngine`/`_seedRegistry` **duplican**
   Docs y filas si se re-ejecutan. Correr una sola vez por entorno.
7. **Bootstrap de rol** — `ADMIN_EMAILS_LIST` solo lo respeta `getUserRole`, no
   `_requireRole`. El **primer `owner` debe existir como fila `active` en
   `Admin_Team`** o no podrá hacer acciones de owner.
8. **CI apagado** — el push automático a Apps Script (`clasp-push.yml`) no corre
   hasta cargar el secret `CLASPRC_JSON` en el repo.

### 🟢 Multi-país / copy (para el piloto)

9. **Solo Colombia completo y probado.** Según la config/skill: MX/PE/UY con
   `Country_Settings` parcial; CL/AR/EC/CR con campos `[VERIFICAR]`; **BR bloqueado**
   (requiere plantilla en **portugués**, que no existe). **A verificar en la hoja.**
   El guardarraíl A1 ahora **bloquea** generar en un país sin config legal completa
   (bueno: falla seguro en vez de emitir un T&C inválido).
10. **Supuestos colombianos hardcodeados** — locale `es-CO`, `numeroALetras` solo
    ES, `NIT`, Ley 1581, SIC, RAPPI S.A.S., URLs `.com.co` en las plantillas. Para
    otros países hay que redactar Docs locales (no reusar el de CO) y generalizar
    código.
11. **Copy inconsistente** — el modal de responsabilidad aún dice *"solo disponible
    para Colombia"* (`WebApp.Html:761`). Menor, pero corregir para el piloto.

---

## 5. Correcciones al README/skill (cosas que **ya no** aplican)

Para no perder tiempo persiguiendo pendientes ya resueltos:

- ❌ *"Sin tests"* → ✅ **31 tests en verde**.
- ❌ *"Fallback silencioso a CO/COP"* → ✅ **bloqueado por FASE A1**.
- ❌ *"Pestaña Actividad rompe por `#panel-logs` ausente"* → ✅ **ya no hay
  referencias a `panel-logs`** (0 en `WebApp.Html`).
- ❌ *"`Helper.gs` es un duplicado roto de Admin"* (nota de la skill) → ✅
  `Helpers.gs.js` está **limpio** (utilidades correctas). La skill está
  desactualizada en ese punto.
- ❌ *"Landing dice 8 Países vs 9"* → el literal *"8 Países"* ya no aparece.

---

## 6. Próximos pasos (priorizados)

**P0 — Inmediato (desbloquea "Organizador" en Cashback; es operacional, no código)**
1. En Apps Script **/dev**, correr el runbook `docs/RUNBOOK_FASE_C.md`:
   `previewFieldDerivation('Cashback','ALL')` (dry-run) → `applyFieldDerivation(...)`
   → correr **otra vez** para probar idempotencia (`added:0`).
2. Encender `RAPPIMIND_A2='on'` en Propiedades del script (/dev primero).
3. **Prueba end-to-end Cashback CO:** país → tipo → campos (con grupo Organizador)
   → generar → abrir Doc → verificar **cero `{{}}` y cero `[ ]`**, moneda/fechas y
   jurisdicción de CO correctas. Repetir en **prod** cuando pase.

**P1 — Antes del piloto (seguridad + features rotas)**
4. Añadir `_requireRole` a las 5 funciones admin sin gate (§4.1).
5. Definir política de compartición de los Docs generados (§4.2) con Privacidad.
6. Implementar `askGemini` + `saveFeedback` en backend **o** retirar los botones.

**P2 — Habilitar multi-país (por país piloto)**
7. Completar `Country_Settings` del país (incl. columnas legales) y **validar con
   Legal local**; corregir copy multipaís.
8. Redactar/registrar **plantillas locales** (no reusar el Doc de CO); confirmar el
   ruteo de carpetas por país.

**P3 — Higiene / infraestructura**
9. Idempotencia del setup; cargar `CLASPRC_JSON` para activar el CI; normalizar
   nombres de archivo del servidor.

**P4 — Roadmap V4 (post-piloto)**
10. Selector de vertical Delivery/Travel · banco de campos · UX pregunta-a-pregunta
    · leyendas post-generación · tracking Squarespace · i18n + portugués (BR).

---

## 7. Veredicto

- **Estado:** piloto **Colombia-only funcional**, con base técnica **notablemente
  más sólida** tras FASE A–D (guardarraíles legales + tests).
- **Camino más corto a valor:** ejecutar **P0** (runbook FASE C + A2) → deja Cashback
  CO redondo con el "Organizador" y validación anti-marcadores.
- **Puerta al piloto multi-país:** cerrar **P1** (seguridad + features rotas) y **P2**
  (datos legales por país). El mayor riesgo hoy no es el motor —es **exposición de
  funciones admin** y **datos legales por país sin validar**.

---

*Revisión generada a partir de lectura directa del backend (`Código.js`,
`Admin.gs.js`, `Config.gs.js`, `Helpers.gs.js`), ejecución de la suite de tests,
runbooks (`RUNBOOK_FASE_C.md`, `DEPLOY_CLASP.md`) y el historial de git (FASE A–D).*
