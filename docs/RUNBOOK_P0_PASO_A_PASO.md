# Runbook P0 — montar /dev y activar Organizador + A2 (paso a paso)

Guía ejecutable para Juan (sin asumir experiencia técnica). Al terminar tendrás:
un **RappiMind /dev** aislado de producción, el form de Cashback pidiendo
**Organizador** (razón social + NIT del aliado), la validación **A2** activa
(no se entrega un T&C con marcadores sin resolver), y el catálogo de **T&C
generales** sembrado.

> Tiempo estimado: 30–40 min. Todo pasa en TU cuenta; producción no se toca.

---

## Parte 1 — Crear el proyecto Apps Script /dev (una vez)

1. Abre https://script.new (crea un proyecto vacío). Nómbralo **"RappiMind DEV"**.
2. Copia el **Script ID**: Configuración del proyecto (⚙️) → "ID de secuencia
   de comandos". Guárdalo.

## Parte 2 — Subir el código del repo a /dev

En tu computador (donde ya usaste clasp antes):

```bash
git clone https://github.com/juangallego-commits/Legal-Rappi-Money-Project.git
cd Legal-Rappi-Money-Project
git checkout claude/revision-integral-proceso-7sbn7k   # la rama con F1
npm install
npx clasp login                       # si no lo has hecho en esta máquina
```

Edita el archivo **`.clasp.json`** y reemplaza el `scriptId` por el de
**RappiMind DEV** (Parte 1). ⚠️ No hagas commit de ese cambio (es local).

```bash
npx clasp push --force
```

En el editor de Apps Script de DEV deben aparecer **11 archivos**: Código,
Admin.gs, Config.gs, Helpers.gs, Setup.gs, WebApp, Crm.gs, CrmForm, CrmStatus,
CrmAdmin y appsscript.json.

## Parte 3 — Aislar la base de datos (copia de la DB de RappiMind)

1. Abre el spreadsheet de RappiMind (DB de prod `1Ki9FvHGk…`) → **Archivo →
   Hacer una copia** → nómbrala "RappiMind DB — DEV". Copia su ID (lo que va
   entre `/d/` y `/edit` en la URL).
2. En **RappiMind DEV** → Configuración del proyecto → **Propiedades del
   script** → añadir:

| Property | Valor |
|---|---|
| `RAPPIMIND_DB_ID` | ID de la copia "RappiMind DB — DEV" |
| `CRM_SPREADSHEET_ID` | `1XR5zWNj5cd3vDLCEeaOX6mqjPliknJ9gHhO6YVzrzuE` (tu copia del Sheet del CRM) |
| `CRM_CSV_FOLDER_ID` | `1T39sOtnG1rd4UCeJZtSs6VDoMCv3Cb5D` (tu carpeta de pruebas) |
| `GEMINI_API_KEY` | (la misma que tienes en prod, si quieres probar el wizard) |

> Con `RAPPIMIND_DB_ID` configurado, TODO lo que haga /dev (plantillas, campos,
> tickets del generador) pasa en la COPIA. Sin esa property, apuntaría a prod.
> Slack queda sin configurar por ahora → el módulo CRM simplemente no notifica
> (todo lo demás funciona). Cuando tengas token: `CRM_SLACK_BOT_TOKEN` y
> `CRM_SLACK_CHANNEL`.

3. ⚠️ La copia de la DB apunta a las MISMAS plantillas de Google Docs (los IDs
   viajan en `Template_Registry`). Generar documentos está bien (hace
   `makeCopy`); solo no borres plantillas desde /dev.

## Parte 4 — Sembrar catálogo de T&C generales (nuevo)

En el editor de **DEV**: selector de función → **`setupTcGenerales`** →
Ejecutar. Debe loguear `added: 13`. (Re-correrlo no duplica.)

## Parte 5 — FASE C: campos de Organizador en Cashback

En el editor de **DEV**, ejecuta en este orden (selector de función → Ejecutar):

1. **`previewFieldDerivation('Cashback','ALL')`** — para pasarle argumentos:
   pega esto en una función temporal o córrela desde una función wrapper:
   ```js
   function _p0_preview() { Logger.log(previewFieldDerivation('Cashback','ALL')); }
   ```
   Esperado en el log: `adds = 2` (organizerLegalName, organizerTaxId),
   `wouldModify 0`, `deletes 0`. **Si difiere → PARAR y avisarme.**
2. ```js
   function _p0_apply() { Logger.log(applyFieldDerivation('Cashback','ALL')); }
   ```
   Esperado: `written: { added: 2, updated: 0 }`.
3. Corre `_p0_apply` **otra vez** → esperado `added: 0` (idempotencia ✓).

## Parte 6 — Activar A2 y probar end-to-end

1. Propiedades del script → añadir **`RAPPIMIND_A2` = `on`**.
2. Implementar → **Nueva implementación** → App web (o "Probar
   implementación") → abrir la URL de /dev.
3. **Prueba 1 (generador):** sin `?page` → Colombia → Cashback → el form debe
   mostrar el grupo **"Organizador"** → llenar → Generar → el Doc debe salir
   con **cero `{{}}` y cero `[ ]`**. Si queda un marcador, verás
   `A2_ABORT: …` con el token culpable (eso es el guardarraíl haciendo su
   trabajo, no un bug).
4. **Prueba 2 (módulo CRM):** `…?page=crm` → crear una solicitud `value`
   dummy → verifica la fila en TU copia del Sheet del CRM y el CSV en tu
   carpeta. `…?page=crm-status` → busca el ticket. `…?page=crm-admin` →
   cambia el estado.

## Parte 7 — Cuando /dev esté verde → producción

1. En el proyecto de **prod**: correr `setupTcGenerales`, `_p0_preview` /
   `_p0_apply` (mismos pasos), y `RAPPIMIND_A2='on'`.
2. Verificar que la plantilla **activa** de Cashback CO en `Template_Registry`
   sea la v2 (la que tiene `{{ORGANIZADOR}}`).
3. Una generación real de humo y listo.

---

**Checklist rápido**

- [ ] Proyecto "RappiMind DEV" creado y `clasp push` con 11 archivos
- [ ] Copia de la DB + `RAPPIMIND_DB_ID`
- [ ] Properties CRM (`CRM_SPREADSHEET_ID`, `CRM_CSV_FOLDER_ID`)
- [ ] `setupTcGenerales` → added: 13
- [ ] FASE C preview → adds:2 → apply → re-apply added:0
- [ ] `RAPPIMIND_A2=on` + e2e Cashback CO con Organizador y doc limpio
- [ ] Módulo CRM probado (`?page=crm`, `crm-status`, `crm-admin`)
- [ ] Réplica a prod
